# -*- coding: utf-8 -*-
"""PP-OCRv6 OCR 工作进程（运行于 openvino venv，JSON-lines 协议）。

默认加载 PP-OCRv6（models/PP-OCRv6/{det,rec}/*.xml），缺失时回退 PP-OCRv5；
可用环境变量 LOCAL_OCR_MODELS / LOCAL_OCR_DICT 强制指定。

为什么不直接用 skill 的 OVOCREngine：
  skill 的 OVDetector/OVRecognizer 在 __init__ 中调用 output(0).get_any_name()，
  而 PP-OCRv5 的 det/rec OpenVINO 模型输出张量没有名字，直接抛
  RuntimeError: Check '!get_names().empty()' failed ... tensor.cpp:102
  本模块改用 compile_model + 按索引取输出（list(res.values())[0]），绕开该缺陷。

  另外 skill 的 detect() 把 32 倍数 resize 后的轮廓坐标当成原图坐标用，
  图像边长不是 32 倍数时框会整体偏移；这里做了尺度映射修正。

协议（stdin/stdout，一行一个 JSON，UTF-8）：
  启动就绪后先推一条无 id 事件：{"ok":true,"event":"ready"}
  请求 {"op":"ping"}                     -> {"ok":true}
  请求 {"op":"ocr","path":"...", ...}    -> {"ok":true,"boxes":[...]}
  请求 {"op":"encode","path":"...", ...} -> {"ok":true,"image":"<base64 jpeg>","width":..,"height":..}
      rectify=true 时输出强制透视矫正后的矩形 sheet（视觉识别用）
      clip=[x1,y1,x2,y2] 只编码该矩形区域（供外部模型识别单个单元格；
        坐标与 path 同坐标系，即 table_grid 返回的 warp_path 上用网格坐标）
      scale=float 编码前放大倍数（默认 1.0；小字单元格建议 2.0）
  请求 {"op":"table_grid","path":"..."} -> {"ok":true,"grid":true,"warp_path":...,
        "hlines":[..],"vlines":[..],"header":[cells],"cells":[cells]}
      检测表格网格并透视矫正（小字实拍表格用）；无可用网格时 grid=false 并给 reason
  请求 {"op":"render_pdf","path":"..","dpi":300}
                                      -> {"ok":true,"pages":[{"path","width","height"},...]}
  请求 {"op":"shutdown"}                 -> {"ok":true} 然后退出
  请求可带 "id"，响应原样回带。

ocr 参数：
  path        图片路径（中文路径安全：按字节读取后 imdecode）
  clip        [x1,y1,x2,y2] 只识别原图该矩形区域（供补行重识别用）
  scale       识别前放大倍数（默认 1.0，配合 clip 做局部放大重识别）
  det_thresh  检测二值化阈值（默认 0.3）
  boxes 返回：{"text":..,"score":..,"bbox":[x1,y1,x2,y2]}，坐标均为原图坐标，按阅读序排列
"""
import base64
from concurrent.futures import ThreadPoolExecutor, as_completed
import json
import math
import os
import re
import sys
import tempfile
import threading
from pathlib import Path

import numpy as np

DET_LIMIT = 2560  # 检测输入最长边限制，超出则等比缩小（坐标仍映射回原图）
RENDER_MAX_EDGE = 3600  # PDF 渲染长边上限：A4@300DPI≈3508px，留少量余量
REC_BATCH = 32    # 识别批大小
REC_INFER_WORKERS = 4
FIND_DET_THRESH = 0.22
FIND_BOX_THRESH = 0.38
TAIL_BLANK = 18384  # 字库 18383 行 + 索引0 blank = 18384 项，模型 18385 类，末尾为备用 blank
DET_DB_UNCLIP_RATIO = 2.0  # DB 后处理 unclip 扩张比例（PaddleOCR det 默认 2.0）
DET_DB_BOX_THRESH = 0.5    # 框内概率均值过滤阈值（PaddleOCR 默认 0.6，取 0.5 保召回）

# 方向校正：整页转 90° 时用"领域词 + 运单号"判断哪一侧才是正的
# （转反了的画面读出来多是形近乱码，字符量再多也不能当证据）。
_ORIENT_WORDS = (
    "日期", "单号", "序号", "承运商", "商品", "数量", "签收", "完好", "签字",
    "交接", "清单", "退仓", "打印", "收货", "运单", "包裹", "备注", "短少",
    "实收", "名称", "破损", "发货", "仓库", "退货", "签收人",
)
_ORIENT_TOKEN_RE = re.compile(r"(?:R\d{2}T)?YT\d{8,}", re.I)


def log(msg: str) -> None:
    """日志只走 stderr，stdout 严格保持 JSON 协议。"""
    try:
        sys.stderr.write(f"[ocr_worker] {msg}\n")
        sys.stderr.flush()
    except Exception:
        pass


def load_image(path: str) -> np.ndarray:
    """按字节读取后解码，避免 cv2.imread 对中文路径失败。"""
    import cv2

    data = Path(path).read_bytes()
    arr = np.frombuffer(data, dtype=np.uint8)
    img = cv2.imdecode(arr, cv2.IMREAD_COLOR)
    if img is None:
        raise ValueError(f"无法解码图片: {path}")
    return img


class OcrEngine:
    """PP-OCRv6 / v5 检测 + 识别（自写薄封装，绕开 skill 的 output-name 缺陷）。"""

    def __init__(self, models_dir: str, dict_path: str, device: str = "CPU"):
        import openvino as ov

        models_dir = Path(models_dir)
        det_xml = str(models_dir / "det" / "det.xml")
        rec_xml = str(models_dir / "rec" / "rec.xml")
        for p in (det_xml, rec_xml):
            if not Path(p).exists():
                raise FileNotFoundError(f"模型不存在: {p}")
        core = ov.Core()
        # The finder runs many independent small images. Ask OpenVINO for a
        # throughput-oriented CPU schedule with enough requests for the worker
        # pool; unsupported runtimes fall back to their default compilation.
        try:
            perf = {ov.properties.hint.performance_mode: "THROUGHPUT",
                    ov.properties.hint.num_requests: 8}
            self.det_model = core.compile_model(det_xml, device, perf)
            self.rec_model = core.compile_model(rec_xml, device, perf)
        except Exception as exc:
            log(f"吞吐优化不可用，使用默认推理设置：{exc}")
            self.det_model = core.compile_model(det_xml, device)
            self.rec_model = core.compile_model(rec_xml, device)
        # A CompiledModel's implicit request is shared by __call__. Fast batch
        # search uses several CPU threads, so give each thread its own request.
        self._infer_local = threading.local()
        self.chars = self._load_dict(dict_path)
        self._prep = None   # (path, 归一化后整图, minfo, 可复用boxes, 可复用texts) 缓存
        self._rect = None   # (path, 强制透视矫正图) 缓存，仅供视觉编码
        self._orient = None # (path, 旋转角, 摆正图, boxes, texts) 方向校正缓存
        log(f"模型就绪 device={device} 字库={len(self.chars)} 项")

    def _infer(self, model, tensor, key: str):
        """Run inference with a request owned by this thread."""
        requests = getattr(self._infer_local, "requests", None)
        if requests is None:
            requests = {}
            self._infer_local.requests = requests
        request = requests.get(key)
        if request is None:
            request = model.create_infer_request()
            requests[key] = request
        result = request.infer({0: tensor})
        return np.asarray(list(result.values())[0])

    @staticmethod
    def _load_dict(path: str):
        """字库：索引 0 为 blank，其后为字库行（保留行首/行尾原样，仅去换行符）。"""
        chars = ["blank"]
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                chars.append(line.strip("\n").strip("\r\n"))
        return chars

    # ── 检测 ─────────────────────────────────────────────

    def detect(self, image: np.ndarray, det_thresh: float = 0.3,
               unclip_ratio: float = DET_DB_UNCLIP_RATIO,
               box_thresh: float = DET_DB_BOX_THRESH):
        """返回轴对齐文本框列表 [x1,y1,x2,y2]（原图坐标）。

        DB 后处理与 PaddleOCR 对齐：阈值化 -> 轮廓 -> 框分数过滤 -> unclip 扩张。
        unclip 不可省：DB 概率图收缩核只覆盖文字中部，直接取轮廓会把字符
        上下沿切掉（实测纯净行文字高 18px，无 unclip 的框只有 9px）。
        venv 无 pyclipper，用 minAreaRect 四边外扩近似（文本轮廓接近矩形）。
        """
        import cv2

        orig_h, orig_w = image.shape[:2]
        h, w = orig_h, orig_w
        if max(h, w) > DET_LIMIT:
            s = DET_LIMIT / float(max(h, w))
            image = cv2.resize(
                image,
                (max(32, int(round(w * s))), max(32, int(round(h * s)))),
                interpolation=cv2.INTER_AREA,
            )
            h, w = image.shape[:2]
        target_h = max(32, int(math.ceil(h / 32)) * 32)
        target_w = max(32, int(math.ceil(w / 32)) * 32)
        resized = cv2.resize(image, (target_w, target_h), interpolation=cv2.INTER_LINEAR)
        tensor = resized.astype(np.float32) / 255.0
        mean = np.array([0.485, 0.456, 0.406], dtype=np.float32).reshape(1, 1, 3)
        std = np.array([0.229, 0.224, 0.225], dtype=np.float32).reshape(1, 1, 3)
        tensor = (tensor - mean) / std
        tensor = np.transpose(tensor, (2, 0, 1))
        tensor = np.expand_dims(tensor, 0).astype(np.float32)

        out = self._infer(self.det_model, tensor, "det")
        prob = out[0]
        if prob.ndim == 3:
            prob = prob[0]
        bitmap = (prob > det_thresh).astype(np.uint8)

        contours, _ = cv2.findContours(bitmap, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        # 轮廓坐标在 32 倍数 resize 后的尺度，映射回原图
        fx = orig_w / float(target_w)
        fy = orig_h / float(target_h)

        boxes = []
        for cnt in contours:
            if cv2.contourArea(cnt) < 50:
                continue
            rect = cv2.minAreaRect(cnt)
            if min(rect[1]) < 5:
                continue
            if box_thresh > 0 and self._box_score(prob, cnt) < box_thresh:
                continue
            expanded = self._unclip_rect(cnt, rect, unclip_ratio)
            if expanded is None:
                continue
            pts = cv2.boxPoints(expanded)
            x1 = max(0, int(round(float(np.min(pts[:, 0])) * fx)))
            y1 = max(0, int(round(float(np.min(pts[:, 1])) * fy)))
            x2 = min(orig_w, int(round(float(np.max(pts[:, 0])) * fx)))
            y2 = min(orig_h, int(round(float(np.max(pts[:, 1])) * fy)))
            if x2 - x1 < 5 or y2 - y1 < 5:
                continue
            boxes.append([x1, y1, x2, y2])
        return boxes

    @staticmethod
    def _box_score(prob: np.ndarray, cnt) -> float:
        """轮廓在概率图上的均值（对应 PaddleOCR box_score_fast）。"""
        import cv2

        mask = np.zeros(prob.shape, dtype=np.uint8)
        cv2.drawContours(mask, [cnt], -1, 1, cv2.FILLED)
        n = int(mask.sum())
        if n == 0:
            return 0.0
        return float(prob[mask > 0].sum()) / n

    @staticmethod
    def _unclip_rect(cnt, rect, ratio: float):
        """unclip 扩张：distance = area*ratio/perimeter，四边各外扩 distance。"""
        import cv2

        peri = cv2.arcLength(cnt, True)
        area = cv2.contourArea(cnt)
        if peri <= 1e-6 or area <= 1e-6:
            return None
        distance = area * ratio / peri
        (cx, cy), (bw, bh), ang = rect
        return (cx, cy), (bw + 2.0 * distance, bh + 2.0 * distance), ang

    # ── 识别 ─────────────────────────────────────────────

    def _prep_crop(self, crop: np.ndarray):
        """单框预处理 -> [3,48,W] 张量（W 动态，不足 16 按 16）。"""
        import cv2

        ch, cw = crop.shape[:2]
        ratio = 48.0 / ch
        tw = max(int(cw * ratio), 16)
        tw = min(tw, 2048)  # 防御极端长条
        resized = cv2.resize(crop, (tw, 48), interpolation=cv2.INTER_LINEAR)
        t = resized.astype(np.float32) / 127.5 - 1.0
        return np.transpose(t, (2, 0, 1))

    def _decode_row(self, ids, probs):
        """CTC 解码：blank=0 折叠重复；索引越界（含 18384 尾部 blank）防御性跳过。"""
        out_ids, out_probs = [], []
        prev = -1
        for idx, prob in zip(ids, probs):
            i = int(idx)
            if i == 0 or i >= len(self.chars) or i == prev:
                prev = i
                continue
            out_ids.append(i)
            out_probs.append(float(prob))
            prev = i
        if not out_ids:
            return "", 0.0
        text = "".join(self.chars[i] for i in out_ids)
        return text, float(np.mean(out_probs))

    @staticmethod
    def _strip_borders(crop: np.ndarray):
        """剥离 crop 四边的表格边框线；剥完为空返回 None，无边框返回原 crop。

        unclip 扩张后 det 框常触到单元格竖线，rec 会把边框读成 '[' 前缀，
        并诱发字符混淆（0->a、9->e）。判定：边缘连续的近全黑列/行
        （暗像素占比>0.85）视为边框线，线宽<=5px 且线后紧跟>=2 列近白
        （占比<0.1）才剥——文字竖笔画罕见同时满足这三条。
        """
        import cv2

        gray = cv2.cvtColor(crop, cv2.COLOR_BGR2GRAY)
        dark = gray < 100
        h, w = dark.shape
        if h < 6 or w < 6:
            return crop
        col_r = dark.mean(axis=0)
        row_r = dark.mean(axis=1)

        def edge_count(ratios):
            """返回 [前端线宽, 后端线宽]，不是边框记 0。"""
            n = len(ratios)
            out = []
            for seq in (ratios, ratios[::-1]):
                k = 0
                while k < n and seq[k] > 0.85:
                    k += 1
                if k == 0 or k > 5:
                    out.append(0)
                    continue
                g = 0
                while k + g < n and seq[k + g] < 0.1:
                    g += 1
                out.append(k if g >= 2 else 0)
            return out

        left, right = edge_count(col_r)
        top, bottom = edge_count(row_r)
        x1, x2 = left, w - right
        y1, y2 = top, h - bottom
        if x2 - x1 < 2 or y2 - y1 < 2:
            return None  # 基本是边框残框
        if left == 0 and right == 0 and top == 0 and bottom == 0:
            return crop
        return crop[y1:y2, x1:x2]

    def _recognize_crops(self, crops, batch_size: int = REC_BATCH):
        """Batch recognition for crops collected from one or many images."""
        results = [("", 0.0)] * len(crops)
        tensors, valid = [], []
        for i, crop in enumerate(crops):
            if crop.size == 0:
                continue
            crop = self._strip_borders(crop)
            if crop is None or crop.size == 0 or min(crop.shape[:2]) < 2:
                continue  # 纯边框残框
            tensors.append(self._prep_crop(crop))
            valid.append(i)

        # 宽度分桶（64px 粒度）：每张裁剪补到本桶上界、批只在桶内切、
        # 批宽=桶宽。单张裁剪的识别输入与批组合无关——单图/批量结果
        # 构造性一致（宽度门放宽后混批 padding 曾让两口径分叉 72 vs 74）。
        batch_size = max(1, int(batch_size or REC_BATCH))
        buckets = {}
        for i, t in enumerate(tensors):
            ow = t.shape[2]
            bw = ((ow - 1) // 64 + 1) * 64
            if ow != bw:
                pad = np.zeros((t.shape[0], t.shape[1], bw), dtype=t.dtype)
                pad[:, :, :ow] = t
                t = pad
            buckets.setdefault(bw, []).append((i, t, ow))
        batches = []
        for bw in sorted(buckets):
            items = buckets[bw]
            for start in range(0, len(items), batch_size):
                part = items[start:start + batch_size]
                batches.append(([i for i, _, _ in part],
                                [t for _, t, _ in part],
                                [w for _, _, w in part]))

        def infer_decode(indexes, chunk, widths):
            mw = max(t.shape[2] for t in chunk)
            # Padding is used only to align a batch; decoding is limited to
            # each crop's valid time axis below.
            batch = np.zeros((len(chunk), 3, 48, mw), dtype=np.float32)
            for j, t in enumerate(chunk):
                batch[j, :, :, : t.shape[2]] = t
            logits = self._infer(self.rec_model, batch, "rec")
            decoded = []
            for j, t in enumerate(chunk):
                # rec 骨干网络横向 8 倍下采样，输出时间步约 width / 8
                # （conv 边界使真实步数可能为 ceil）。只解码当前裁剪图的
                # 有效时间轴（桶补零不算），不能让批内右侧 padding 产生
                # 字符；用 ceil 是防止 .5 边界（Python round 银行家舍入）
                # 把末位字符的时间步吞掉（实测 669d4d7e YT…291 丢末位 1）。
                valid_steps = min(logits.shape[1], max(1, int(math.ceil(widths[j] / 8.0))))
                row = logits[j, :valid_steps]
                ids = np.argmax(row, axis=1)
                probs = np.max(row, axis=1)
                decoded.append((valid[indexes[j]], self._decode_row(ids, probs)))
            return decoded

        if len(batches) == 1:
            for result_index, decoded in infer_decode(*batches[0]):
                results[result_index] = decoded
        elif batches:
            # Keep the recognition model busy across independent batches.
            # Each thread owns its own OpenVINO InferRequest.
            with ThreadPoolExecutor(
                max_workers=min(REC_INFER_WORKERS, len(batches)),
                thread_name_prefix="waybill-rec",
            ) as pool:
                pending = {
                    pool.submit(infer_decode, indexes, chunk, widths): None
                    for indexes, chunk, widths in batches
                }
                for future in as_completed(pending):
                    for result_index, decoded in future.result():
                        results[result_index] = decoded
        return results

    def recognize(self, image: np.ndarray, boxes):
        """Batch OCR boxes from one image (empty boxes return ("", 0.0))."""
        crops = [image[y1:y2, x1:x2] for x1, y1, x2, y2 in boxes]
        return self._recognize_crops(crops)

    def _decide_rotation(self, path: str, img, boxes):
        """判断整页是否被转了 90°，必要时返回摆正结果。

        返回 (角度, 摆正图, boxes, texts)；不需要旋转时为 (0, None, None, None)。
        判据用检测框形状：横排文字的框是"宽扁"的（中位高/宽约 0.3-0.5），
        整页转了 90° 时 DB 只能按单个字符出框，框接近方形（实测 0.75-1.0），
        且每百万像素的框数明显偏多（135-181 vs 正常 43）。命中可疑条件才做
        模型验证（两个旋转方向各 det+rec 一次），其余页面只多一次形状统计。

        方向判定优先看"哪一侧读出了表格里的领域词/运单号"（转反了文字会
        上下颠倒，读出来的多是形近乱码，量再多也不算证据）。两个方向都
        没有领域证据时保持原方向——密排小字的正常表格也会出现"方框多"的
        形状，不能据此乱转（实测 7fc33779 这类 200 行四栏表会被误判）。
        """
        if self._orient is not None and self._orient[0] == path:
            return self._orient[1], self._orient[2], self._orient[3], self._orient[4]
        result = (0.0, None, None, None)
        h, w = img.shape[:2]
        density = len(boxes) / max(0.1, (h * w) / 1e6)
        aspects = [(b[3] - b[1]) / max(1, b[2] - b[0]) for b in boxes]
        med_aspect = float(np.median(aspects)) if aspects else 0.0
        if len(boxes) >= 8 and med_aspect >= 0.65 and density >= 80:
            import cv2

            cands = []
            for ang in (90, 270):
                rot = cv2.rotate(
                    img,
                    cv2.ROTATE_90_CLOCKWISE if ang == 90 else cv2.ROTATE_90_COUNTERCLOCKWISE,
                )
                r_boxes = self.detect(rot)
                texts = self.recognize(rot, r_boxes)
                joined = "".join(t for t, _s in texts if t.strip())
                words = sum(joined.count(word) for word in _ORIENT_WORDS)
                tokens = len(_ORIENT_TOKEN_RE.findall(joined))
                hi = sum(len(t) for t, s in texts if t.strip() and s >= 0.8)
                chars = sum(len(t) for t, s in texts if t.strip() and s >= 0.5)
                cands.append((words * 3 + tokens * 2, hi, chars, ang, rot, r_boxes, texts))
            cands.sort(key=lambda item: (item[0], item[1], item[2]), reverse=True)
            best, second = cands[0], cands[1]
            clear = best[0] >= 6 and best[0] > second[0]
            # 文字量兜底必须建立在"确实读出了领域词/运单号"之上：
            # 正着的密排小字表格转 90° 后会变成形近乱码，碎片字符的置信度
            # 反而更高（实测 849c34ef 领域证据 0 -> 0、高置信字符 46 -> 103），
            # 只看字符量会把正着的多栏表转反，网格检测随之失败、整页丢单号。
            by_text = (best[0] > 0 and best[1] >= 100
                       and best[1] >= int(second[1] * 1.5)
                       and best[2] >= int(second[2] * 1.3))
            if clear or by_text:
                log(f"方向校正: 版面疑似转了 90°（框密度 {density:.0f}/MP、"
                    f"中位高宽比 {med_aspect:.2f}），"
                    f"采用{'顺时针' if best[3] == 90 else '逆时针'} 90°"
                    f"（领域证据 {second[0]} -> {best[0]}，高置信字符 {second[1]} -> {best[1]}）")
                result = (float(best[3]), best[4], best[5], best[6])
            else:
                log(f"方向校正: 版面疑似转了 90°，但两个方向都没有领域证据"
                    f"（领域证据 {second[0]} / {best[0]}，高置信字符 {second[1]} / {best[1]}），"
                    f"保持原方向")
        self._orient = (path, *result)
        return result

    def decide_rotation(self, path: str) -> int:
        """供 table_grid 等非 OCR 操作使用：返回该图需要的旋转角（0/90/270）。"""
        img = load_image(path)
        boxes = self.detect(img)
        angle, _img, _b, _t = self._decide_rotation(path, img, boxes)
        return int(angle)

    # ── 全流程 ───────────────────────────────────────────

    def _prepared(self, path: str):
        """按路径缓存文档几何归一化：透视矫正（实测更优才采用）-> 旋转去斜兜底。

        返回 (处理后整图, minfo, reuse_boxes, reuse_texts)。reuse_boxes/texts
        是最终图像上已完成的检测框与识别结果：透视矫正被采用时即矫正图的
        结果；未做几何变换时即原图结果（_prepared 内补识别，与 ocr() 自跑的
        结果逐字节一致）。整页 90° 方向校正时复用摆正图的识别结果；去斜
        小角度旋转会重采样图像，此时不提供复用（置 None），由 ocr() 在旋转
        后图像上重新检测识别，保证识别质量。

        复用省掉的是重复推理：实测 200 行四栏照可省一次 detect+recognize
        （约 6s）；检测/识别与返回的 box 坐标、以及后续 clip 补行请求，
        全部在该归一化坐标系中进行，天然一致。
        """
        if self._prep is not None and self._prep[0] == path:
            return self._prep[1], self._prep[2], self._prep[3], self._prep[4]
        import cv2

        img = load_image(path)
        minfo = None
        if Path(path).name.startswith("img2xlsx_grid_"):
            # table_grid 生成的 warp 已完成透视矫正；再次归一化会让 OCR 框
            # 与 hlines/vlines 处于不同坐标系，造成串行、串列和局部裁剪错位。
            raw_boxes = self.detect(img)
            reuse_texts = self.recognize(img, raw_boxes)
            self._prep = (path, img, None, raw_boxes, reuse_texts)
            return img, None, raw_boxes, reuse_texts
        raw_boxes = self.detect(img)
        angle, upright, up_boxes, up_texts = self._decide_rotation(path, img, raw_boxes)
        if angle and upright is not None:
            img = upright
            raw_boxes = up_boxes
        quad = _find_doc_quad(cv2.cvtColor(img, cv2.COLOR_BGR2GRAY))
        if angle:
            reuse_boxes, reuse_texts = up_boxes, up_texts
        else:
            reuse_boxes, reuse_texts = None, None
        if quad is not None:
            warped, m = _quad_warp(img, quad)
            if warped is not None:
                w_boxes = self.detect(warped)
                if len(w_boxes) >= max(6, 0.6 * len(raw_boxes)):
                    w_texts = self.recognize(warped, w_boxes)
                    r_texts = self.recognize(img, raw_boxes)
                    w_med = float(np.median([s for t, s in w_texts if t.strip()] or [0]))
                    r_med = float(np.median([s for t, s in r_texts if t.strip()] or [0]))
                    if w_med >= r_med + 0.005:
                        img, minfo = warped, ("persp", m)
                        reuse_boxes, reuse_texts = w_boxes, w_texts
                        log(f"几何归一化: 透视矫正（识别质量 {r_med:.3f} -> {w_med:.3f}）")
                    else:
                        log(f"几何归一化: 检测到纸张边界，但透视矫正后识别变差"
                            f"（{r_med:.3f} -> {w_med:.3f}），保留原图")
                        # 原图胜出：识别结果已在手，仅当后续不做旋转才可复用
                        reuse_boxes, reuse_texts = raw_boxes, r_texts
        if minfo is None:
            ang = _skew_from_boxes(raw_boxes)
            if abs(ang) >= 0.8:
                h, w = img.shape[:2]
                m = cv2.getRotationMatrix2D((w / 2.0, h / 2.0), ang, 1.0)
                img = cv2.warpAffine(
                    img, m, (w, h), flags=cv2.INTER_CUBIC,
                    borderMode=cv2.BORDER_CONSTANT, borderValue=(255, 255, 255),
                )
                minfo = ("rotate", m, ang)
                log(f"几何归一化: 文档倾斜 {ang:.1f}°，已旋转矫正")
                reuse_boxes, reuse_texts = None, None   # 图像已变，不能复用
            elif reuse_boxes is None:
                # 无几何变换：补识别原图框供复用（与 ocr() 自行检测识别一致）
                reuse_boxes, reuse_texts = raw_boxes, self.recognize(img, raw_boxes)
        self._prep = (path, img, minfo, reuse_boxes, reuse_texts)
        return img, minfo, reuse_boxes, reuse_texts

    def ocr(self, req: dict):
        path = req["path"]
        clip = req.get("clip")
        scale = float(req.get("scale") or 1.0)
        if scale <= 0:
            scale = 1.0
        det_thresh = float(req.get("det_thresh") or 0.3)
        unclip_ratio = float(req.get("unclip_ratio") or DET_DB_UNCLIP_RATIO)
        box_thresh = float(req.get("box_thresh") or DET_DB_BOX_THRESH)

        img_full, _minfo, cached_boxes, cached_texts = self._prepared(path)
        H, W = img_full.shape[:2]
        ox = oy = 0
        if clip:
            # clip 坐标与既往返回的 box 同坐标系（归一化后）
            x1, y1, x2, y2 = [int(v) for v in clip]
            x1, y1 = max(0, x1), max(0, y1)
            x2, y2 = min(W, x2), min(H, y2)
            if x2 <= x1 or y2 <= y1:
                return []
            img = img_full[y1:y2, x1:x2]
            ox, oy = x1, y1
        else:
            img = img_full
        eff_scale = scale
        if abs(scale - 1.0) > 1e-6:
            img = _resize(img, scale)

        # 无 clip/缩放且检测参数为默认时，直接复用 _prepared 已算出的
        # 检测+识别（同一图像、同一参数，结果一致），省掉一次重复推理
        if (clip is None and abs(scale - 1.0) <= 1e-6
                and cached_boxes is not None
                and det_thresh == 0.3
                and unclip_ratio == DET_DB_UNCLIP_RATIO
                and box_thresh == DET_DB_BOX_THRESH):
            boxes, texts = cached_boxes, cached_texts
        else:
            boxes = self.detect(img, det_thresh, unclip_ratio, box_thresh)
            texts = self.recognize(img, boxes)

        # 小字号自适应：框数增加时采用放大结果；框数相同则比较文字置信度，
        # 避免“放大后读对了，但因为没多出文本框而被丢弃”。
        if boxes and float(np.median([b[3] - b[1] for b in boxes])) < 15.0 \
                and max(img.shape[:2]) <= 3200:
            up = _resize(img, 2.0)
            b2 = self.detect(up, det_thresh, unclip_ratio, box_thresh)
            t2 = self.recognize(up, b2)
            old_scores = [s for t, s in texts if t.strip()]
            new_scores = [s for t, s in t2 if t.strip()]
            n_old, n_new = len(old_scores), len(new_scores)
            med_old = float(np.median(old_scores or [0.0]))
            med_new = float(np.median(new_scores or [0.0]))
            if n_new > n_old or (n_new == n_old and n_new > 0
                                 and med_new >= med_old + 0.005):
                log(f"文本过小（中位高 <15px），采用放大重识别: "
                    f"{n_old} -> {n_new} 框，置信度 {med_old:.3f} -> {med_new:.3f}")
                boxes, texts, eff_scale = b2, t2, scale * 2.0

        out = []
        for (x1, y1, x2, y2), (text, score) in zip(boxes, texts):
            if not text.strip():
                continue
            # 映射回归一化原图坐标：除以总放大倍数 + 加 clip 原点
            bx1 = int(round(ox + x1 / eff_scale))
            by1 = int(round(oy + y1 / eff_scale))
            bx2 = int(round(ox + x2 / eff_scale))
            by2 = int(round(oy + y2 / eff_scale))
            bx1, by1 = max(0, bx1), max(0, by1)
            bx2, by2 = min(W, bx2), min(H, by2)
            out.append({"text": text, "score": round(score, 4), "bbox": [bx1, by1, bx2, by2]})
        # 阅读序：先 y 后 x
        out.sort(key=lambda b: (b["bbox"][1], b["bbox"][0]))
        return out

    @staticmethod
    def _search_candidates(boxes):
        """Keep long text lines and nearby adjacent boxes that may form one code."""
        chosen = set()
        for i, (x1, y1, x2, y2) in enumerate(boxes):
            w, h = x2 - x1, y2 - y1
            if h >= 7 and w >= 40 and w / max(1, h) >= 4.0:
                chosen.add((x1, y1, x2, y2))

        # DB sometimes splits a printed tracking number into characters. Join
        # horizontally adjacent boxes on the same line and OCR the combined crop.
        rows = []
        for box in sorted(boxes, key=lambda b: ((b[1] + b[3]) / 2, b[0])):
            x1, y1, x2, y2 = box
            cy, h = (y1 + y2) / 2, y2 - y1
            row = next((r for r in rows
                        if abs(cy - r["cy"]) <= max(7, 0.55 * max(h, r["h"]))), None)
            if row is None:
                rows.append({"cy": cy, "h": h, "boxes": [box]})
            else:
                row["boxes"].append(box)
                row["cy"] = sum((b[1] + b[3]) / 2 for b in row["boxes"]) / len(row["boxes"])
                row["h"] = max(row["h"], h)
        for row in rows:
            line = sorted(row["boxes"], key=lambda b: b[0])
            segment = []
            for box in line:
                if segment:
                    prev = segment[-1]
                    h = max(prev[3] - prev[1], box[3] - box[1])
                    gap = box[0] - prev[2]
                    overlap = min(prev[3], box[3]) - max(prev[1], box[1])
                    if gap > max(18, 1.6 * h) or overlap < -0.25 * h:
                        if len(segment) >= 5:
                            merged = (min(b[0] for b in segment), min(b[1] for b in segment),
                                      max(b[2] for b in segment), max(b[3] for b in segment))
                            if ((merged[2] - merged[0]) >= 80
                                    and (merged[2] - merged[0]) / max(1, merged[3] - merged[1]) >= 4.0):
                                chosen.add(merged)
                        segment = []
                segment.append(box)
            if len(segment) >= 5:
                merged = (min(b[0] for b in segment), min(b[1] for b in segment),
                          max(b[2] for b in segment), max(b[3] for b in segment))
                if ((merged[2] - merged[0]) >= 80
                        and (merged[2] - merged[0]) / max(1, merged[3] - merged[1]) >= 4.0):
                    chosen.add(merged)
            for left, right in zip(line, line[1:]):
                h = max(left[3] - left[1], right[3] - right[1])
                gap = right[0] - left[2]
                overlap = min(left[3], right[3]) - max(left[1], right[1])
                if gap <= max(18, 1.6 * h) and overlap >= -0.25 * h:
                    merged = (min(left[0], right[0]), min(left[1], right[1]),
                              max(left[2], right[2]), max(left[3], right[3]))
                    if ((merged[2] - merged[0]) >= 80
                            and (merged[2] - merged[0]) / max(1, merged[3] - merged[1]) >= 4.0):
                        chosen.add(merged)
        candidates = list(chosen)
        areas = {box: max(0, box[2] - box[0]) * max(0, box[3] - box[1])
                 for box in candidates}
        compact = []
        for big in candidates:
            big_area = max(1, areas[big])
            duplicate = False
            for small in candidates:
                small_area = areas[small]
                if small == big or small_area >= big_area * 0.72:
                    continue
                iw = max(0, min(big[2], small[2]) - max(big[0], small[0]))
                ih = max(0, min(big[3], small[3]) - max(big[1], small[1]))
                if iw * ih / max(1, small_area) >= 0.82:
                    duplicate = True
                    break
            if not duplicate:
                compact.append(big)
        return sorted(compact, key=lambda b: (b[1], b[0]))

    def _prepare_fast_find(self, path: str, max_edge: int):
        """Detect likely waybill lines and return crops without recognizing them.

        检测在降采样图上做（快），但识别裁剪取原图分辨率：小图上 7-8px 高的
        单号行放大到 48px 全是插值模糊，实测 840 小图裁剪命中率仅 26%。
        识别张量宽度只取决于裁剪长宽比（高度统一归一化到 48），原图裁剪
        不增加推理量，只提升像素质量。
        """
        import cv2

        orig = load_image(path)
        orig_h, orig_w = orig.shape[:2]
        scale = min(1.0, float(max_edge) / max(orig_h, orig_w))
        image = orig
        if scale < 1.0:
            image = cv2.resize(orig, (max(1, round(orig_w * scale)),
                                      max(1, round(orig_h * scale))),
                               interpolation=cv2.INTER_AREA)
        h, w = image.shape[:2]
        raw_boxes = self.detect(image, det_thresh=FIND_DET_THRESH,
                                box_thresh=FIND_BOX_THRESH)
        candidates = self._search_candidates(raw_boxes)
        variants = [(image, candidates, 0)]
        # On a rotated page DB tends to return character-sized boxes. Re-detect
        # the two portrait orientations only when the first pass found no usable
        # long lines; ordinary upright pages stay on the one-pass fast path.
        landscape_page = orig_w > orig_h * 1.15
        if len(candidates) < 2 or landscape_page:
            for angle, code in ((90, cv2.ROTATE_90_CLOCKWISE),
                                (270, cv2.ROTATE_90_COUNTERCLOCKWISE)):
                rotated = cv2.rotate(image, code)
                rot_boxes = self.detect(rotated, det_thresh=FIND_DET_THRESH,
                                        box_thresh=FIND_BOX_THRESH)
                rot_candidates = self._search_candidates(rot_boxes)
                if rot_candidates:
                    variants.append((rotated, rot_candidates, angle))

        prepared = []
        full_by_angle = {0: orig}
        for variant, boxes, angle in variants:
            variant_width = variant.shape[1]
            # 宽度门只按页宽相对口径（0.28）：绝对上限 180 在竖幅手机照
            # （1080x1920，检测宽 945）上把整页单号行全判超宽弃掉，实测
            # 3c7a2af9 88 个候选全弃、只留 1 个救援裁剪 → 0/44 命中。
            max_crop_width = max(140, int(round(0.28 * variant_width)))
            if angle not in full_by_angle:
                code = (cv2.ROTATE_90_CLOCKWISE if angle == 90
                        else cv2.ROTATE_90_COUNTERCLOCKWISE)
                full_by_angle[angle] = cv2.rotate(orig, code)
            full = full_by_angle[angle]
            fh, fw = full.shape[:2]
            for box in boxes:
                x1, y1, x2, y2 = box
                if angle == 90:
                    mapped = [round(y1 / scale), round((h - x2) / scale),
                              round(y2 / scale), round((h - x1) / scale)]
                elif angle == 270:
                    mapped = [round((w - y2) / scale), round(x1 / scale),
                              round((w - y1) / scale), round(x2 / scale)]
                else:
                    mapped = [round(v / scale) for v in box]
                mapped = [max(0, mapped[0]), max(0, mapped[1]),
                          min(orig_w, mapped[2]), min(orig_h, mapped[3])]
                # 检测框映射回原图分辨率再裁剪。垂直留 10% 行高余量；
                # 水平不对称：右侧留 60% 行高（小图检测框常切掉单号末位，
                # 实测 840 下 YT1234567890123 读成 YT123456789012），
                # 左侧只留 15%（余量太大会把单元格竖线卷进来吞掉首字符，
                # 实测 c2c7b955 左余量 12px 时 YT2549… 读成 T2549…）
                pad_y = max(1, int(round(0.10 * (y2 - y1))))
                pad_l = max(1, int(round(0.15 * (y2 - y1))))
                pad_r = max(2, int(round(0.60 * (y2 - y1))))
                fx1 = max(0, int(round((x1 - pad_l) / scale)))
                fy1 = max(0, int(round((y1 - pad_y) / scale)))
                fx2 = min(fw, int(round((x2 + pad_r) / scale)))
                fy2 = min(fh, int(round((y2 + pad_y) / scale)))
                crop = full[fy1:fy2, fx1:fx2]
                if crop.size == 0:
                    continue
                prepared.append({"crop": crop.copy(), "bbox": mapped,
                                 "rotation": angle,
                                 "max_crop_width": max_crop_width,
                                 "det_width": x2 - x1})
        return prepared

    def fast_find(self, path: str, max_edge: int = 1680):
        """Fast, targeted OCR for one image; coordinates map to the source image."""
        results = self.fast_find_batch([path], max_edge=max_edge, workers=1)
        if not results or not results[0].get("ok"):
            raise RuntimeError((results[0] if results else {}).get("error", "识别失败"))
        return results[0].get("boxes") or []

    def fast_find_batch(self, paths: list[str], max_edge: int = 1680,
                        workers: int = 8):
        """Run targeted OCR for several images concurrently using thread-local requests."""
        if not paths:
            return []
        try:
            import cv2

            cv2.setNumThreads(1)
        except Exception:
            pass
        workers = max(1, min(int(workers or 1), len(paths), 8,
                             os.cpu_count() or 1))
        prepared = [None] * len(paths)
        with ThreadPoolExecutor(max_workers=workers, thread_name_prefix="waybill-ocr") as pool:
            pending = {pool.submit(self._prepare_fast_find, path, max_edge): i
                       for i, path in enumerate(paths)}
            for future in as_completed(pending):
                index = pending[future]
                try:
                    prepared[index] = future.result()
                except Exception as exc:
                    prepared[index] = exc

        out = [{"ok": True, "boxes": []} if not isinstance(item, Exception)
               else {"ok": False, "error": str(item)} for item in prepared]
        crops, owner = [], []
        for image_index, item in enumerate(prepared):
            if isinstance(item, Exception):
                continue
            candidates = [candidate for candidate in item
                          if candidate["det_width"] <= candidate["max_crop_width"]]
            # Keep one rescue crop for unusually narrow/zoomed images where the
            # only line candidate is wider than the normal tracking-number band.
            if not candidates and item:
                candidates = [min(item, key=lambda c: c.get(
                    "det_width", c["crop"].shape[1]))]
            for candidate in candidates:
                crops.append(candidate["crop"])
                owner.append((image_index, candidate))
        try:
            # The recognizer accepts a dynamic batch. A larger batch cuts the
            # per-call overhead substantially when searching many images.
            texts = self._recognize_crops(crops, batch_size=REC_BATCH)
        except Exception as exc:
            for image_index, item in enumerate(prepared):
                if not isinstance(item, Exception):
                    out[image_index] = {"ok": False, "error": str(exc)}
            return out
        for (image_index, candidate), (text, score) in zip(owner, texts):
            if not text.strip():
                continue
            out[image_index]["boxes"].append({
                "text": text, "score": round(score, 4), "bbox": candidate["bbox"],
                "rotation": candidate["rotation"],
            })
        return out

    def _rectified(self, path: str):
        """强制透视矫正版整图（供视觉识别用）：检测到纸张四边形就拉正，
        找不到（或四边形退化/过小）时退回流 _prepared 的归一化图。

        与 _prepared 分开缓存：OCR 坐标系不能被矫正图污染（clip 补行、
        交叉校验的几何定位都以 OCR 归一化坐标为基准）。视觉模型相反——
        带桌面背景、透视倾斜的原图会让它长时间"找纸"而 read timeout
        （实测 200 行四栏照 3 种尺寸全部 90s 超时；矫正后 154s 识别成功），
        干净的矩形 sheet 才是它擅长的输入。

        四边形过小（误检：桌面阴影/深色区块等）时矫正图没有意义，
        要求拉正后边长不低于原图长边的 40%，否则退回归一化图。
        """
        if self._rect is not None and self._rect[0] == path:
            return self._rect[1]
        import cv2

        img = load_image(path)
        orig_edge = max(img.shape[:2])
        boxes = self.detect(img)
        angle, upright, _b, _t = self._decide_rotation(path, img, boxes)
        if angle and upright is not None:
            img = upright
        quad = _find_doc_quad(cv2.cvtColor(img, cv2.COLOR_BGR2GRAY))
        if quad is not None:
            warped, _m = _quad_warp(img, quad)
            if warped is not None and max(warped.shape[:2]) >= 0.4 * orig_edge:
                self._rect = (path, warped)
                return warped
        normalized, _minfo, _b, _t = self._prepared(path)
        self._rect = (path, normalized)
        return normalized

    def encode(self, req: dict) -> dict:
        """把图片编码为 JPEG base64，供外部（云端）模型识别。

        主进程（打包后的 exe）没有 cv2，图片预处理必须在本子进程内完成。
        rectify=True 时优先输出透视矫正后的矩形 sheet（去背景、拉正）；
        clip=[x1,y1,x2,y2] 时只编码该矩形（坐标与 path 同坐标系）——
        外部模型只需要存疑单元格的图，绝不把整张照片发出去；
        scale 在裁剪后再放大，小字单元格建议 2.0 以上。
        """
        import cv2

        path = req["path"]
        limit = int(req.get("limit") or 2048)
        quality = int(req.get("quality") or 90)
        if req.get("rectify"):
            img = self._rectified(path)
        else:
            img, _minfo, _b, _t = self._prepared(path)
        clip = req.get("clip")
        if clip:
            H, W = img.shape[:2]
            x1, y1, x2, y2 = [int(v) for v in clip]
            x1, y1 = max(0, x1), max(0, y1)
            x2, y2 = min(W, x2), min(H, y2)
            if x2 <= x1 or y2 <= y1:
                raise RuntimeError(f"clip 区域为空：{clip}")
            img = img[y1:y2, x1:x2]
        scale = float(req.get("scale") or 1.0)
        if scale > 1.05:
            h, w = img.shape[:2]
            img = cv2.resize(img, (max(1, int(round(w * scale))), max(1, int(round(h * scale)))),
                             interpolation=cv2.INTER_CUBIC)
        h, w = img.shape[:2]
        longest = max(h, w)
        if longest > limit > 0:
            shrink = limit / float(longest)
            img = cv2.resize(
                img, (max(1, int(round(w * shrink))), max(1, int(round(h * shrink)))),
                interpolation=cv2.INTER_AREA,
            )
            h, w = img.shape[:2]
        ok, buf = cv2.imencode(".jpg", img, [cv2.IMWRITE_JPEG_QUALITY, quality])
        if not ok:
            raise RuntimeError("JPEG 编码失败")
        return {
            "image": base64.b64encode(buf.tobytes()).decode("ascii"),
            "width": int(w),
            "height": int(h),
        }


def _resize(img: np.ndarray, scale: float) -> np.ndarray:
    import cv2

    h, w = img.shape[:2]
    nw = max(1, int(round(w * scale)))
    nh = max(1, int(round(h * scale)))
    interp = cv2.INTER_CUBIC if scale > 1.0 else cv2.INTER_AREA
    return cv2.resize(img, (nw, nh), interpolation=interp)


# ---------------------------------------------------------------- 文档几何归一化
#
# 手机实拍的文档照片常有两种几何问题，都会让"行聚类"失效：
#   a) 倾斜（俯拍歪了 1-5°）：同一行文字的 y 区间与相邻行互相重叠；
#   b) 透视（纸不在画面正中/侧拍）：行斜率随 x 变化。
# 在 OCR 前把图拉正：优先透视矫正（检测纸张四边形），失败则按文本框
# RANSAC 估出的倾角做旋转。输出的 box 坐标均在矫正后坐标系，
# fill_serial_gaps 的 clip 请求也按该坐标系解释（见 ocr()）。


def _find_doc_quad(gray: np.ndarray):
    """在灰度图上找纸张四边形（原图坐标的 4x2 float32），找不到返回 None。"""
    import cv2

    h, w = gray.shape[:2]
    scale = 1000.0 / max(h, w) if max(h, w) > 1000 else 1.0
    small = cv2.resize(gray, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA) if scale != 1.0 else gray
    blur = cv2.GaussianBlur(small, (5, 5), 0)
    _, bw = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    bw = cv2.morphologyEx(bw, cv2.MORPH_CLOSE, np.ones((5, 5), np.uint8))
    contours, _ = cv2.findContours(bw, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    area_img = small.shape[0] * small.shape[1]
    best, best_area = None, 0.0
    for cnt in contours:
        area = cv2.contourArea(cnt)
        if area < 0.25 * area_img or area > 0.97 * area_img:
            continue
        peri = cv2.arcLength(cnt, True)
        approx = None
        for eps in (0.02, 0.03, 0.05):          # 纸边微弯时放宽逐步重试
            approx = cv2.approxPolyDP(cnt, eps * peri, True)
            if len(approx) == 4:
                break
        if approx is None or len(approx) != 4 or not cv2.isContourConvex(approx):
            continue
        pts = approx.reshape(4, 2).astype(np.float32)
        sides = [np.linalg.norm(pts[i] - pts[(i + 1) % 4]) for i in range(4)]
        if min(sides) < 0.3 * (sum(sides) / 4):  # 排除细长/退化四边形
            continue
        if area > best_area:
            best_area, best = area, pts
    if best is None:
        return None
    return best / scale


def _order_quad(pts: np.ndarray) -> np.ndarray:
    """四点按 左上/右上/右下/左下 排序。"""
    s = pts.sum(axis=1)
    d = (pts[:, 1] - pts[:, 0]).ravel()
    tl, br = pts[np.argmin(s)], pts[np.argmax(s)]
    tr, bl = pts[np.argmin(d)], pts[np.argmax(d)]
    return np.array([tl, tr, br, bl], dtype=np.float32)


def _quad_warp(img: np.ndarray, quad: np.ndarray):
    """透视矫正到矩形。返回 (图, 3x3 矩阵)；目标边长超原图 1.5 倍部分按比例缩回。"""
    import cv2

    quad = _order_quad(quad)

    def dist(a, b):
        return float(np.linalg.norm(a - b))

    tw = max(dist(quad[0], quad[1]), dist(quad[3], quad[2]))
    th = max(dist(quad[0], quad[3]), dist(quad[1], quad[2]))
    if tw < 60 or th < 60:
        return None, None
    cap = 1.5 * max(img.shape[:2])              # 防止低分辨率纸面被过分放大
    shrink = min(1.0, cap / max(tw, th))
    tw, th = int(round(tw * shrink)), int(round(th * shrink))
    dst = np.array([[0, 0], [tw - 1, 0], [tw - 1, th - 1], [0, th - 1]], dtype=np.float32)
    m = cv2.getPerspectiveTransform(quad, dst)
    out = cv2.warpPerspective(img, m, (tw, th), flags=cv2.INTER_CUBIC,
                              borderMode=cv2.BORDER_CONSTANT, borderValue=(255, 255, 255))
    return out, m


# ---------------------------------------------------------------- 表格网格检测
#
# 手机实拍的表格照片上，PP-OCR 的文本框常把「日期 + 承运商 + 运单号」粘成一个
# 跨列长框，行聚类/列带聚类无从下手（表现为三列内容挤进一列、大量错行）。
# 根本解法是回到表格自身的结构：检测网格线 -> 透视矫正拉正 -> 按线切格 ->
# 每个框按中心点落格。合并单元格天然正确（该区间没有竖线穿过）。


def _line_mask(bw, horiz: bool, size_ref: int):
    """形态学提取横线（horiz=True）或竖线的二值掩膜。"""
    import cv2

    if horiz:
        k = max(11, int(size_ref * 0.06) | 1)
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (k, 1))
    else:
        k = max(11, int(size_ref * 0.025) | 1)
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, k))
    return cv2.morphologyEx(bw, cv2.MORPH_OPEN, kernel)


def _line_runs(mask, horiz: bool):
    """把线掩膜切成一段段线，返回 [{"pos","start","len"}]。

    horiz=True 时 pos 为 y、start/len 为 x 方向跨度；False 时相反。
    len 取段内最长连续亮跑，避免被垂直穿过的线撑宽。
    """
    proj = (mask > 0).sum(axis=1 if horiz else 0)
    n = len(proj)
    out = []
    i = 0
    while i < n:
        if proj[i] == 0:
            i += 1
            continue
        j = i
        while j < n and proj[j] > 0:
            j += 1
        band = mask[i:j] if horiz else mask[:, i:j]
        row_any = band.any(axis=0 if horiz else 1)
        best = cur = 0
        best_s = cur_s = 0
        for k, v in enumerate(row_any):
            if v:
                if cur == 0:
                    cur_s = k
                cur += 1
                if cur > best:
                    best, best_s = cur, cur_s
            else:
                cur = 0
        if best > 0:
            out.append({"pos": (i + j - 1) / 2.0, "start": best_s, "len": best})
        i = j
    return out


def _slant_vline_runs(bw, y0: float, y1: float, min_len: float):
    """抗残斜的竖线检测：小斜率剪切堆叠 + 形态学竖直核。

    矩形拉伸吸收不了梯形畸变，warp 上左半列线可斜 1° 以上，轴对齐竖直核
    (1,45) 要求 45px 窗口内绝对竖直，斜线整条被打碎（实测 16 列表只剩右半
    4 条竖线）。把图按斜率 k 剪切 x'=x-k*(y-cy) 后斜线变竖直，核即可检出；
    剪切锚在 cy，故检出 pos 直接就是 cy 处的 x，跨斜率聚类不会散。
    返回 [{"pos","start","len","k"}]，pos 为 cy 处 x，按 pos 升序、近距去重。
    """
    import cv2
    import numpy as np

    th = y1 - y0
    H, W = bw.shape[:2]
    if th < 120 or min_len < 40:
        return []
    cy = (y0 + y1) / 2.0
    out = []
    for k in (-0.06, -0.04, -0.02, 0.0, 0.02, 0.04, 0.06):
        if k == 0.0:
            sh = bw
        else:
            M = np.array([[1.0, -k, k * cy], [0.0, 1.0, 0.0]], dtype=np.float32)
            sh = cv2.warpAffine(bw, M, (W, H), flags=cv2.INTER_NEAREST,
                                borderValue=0)
        vm = _line_mask(sh, False, int(th))
        for s in _line_runs(vm, False):
            if s["len"] >= min_len:
                out.append({"pos": s["pos"], "start": s["start"],
                            "len": s["len"], "k": k})
    out.sort(key=lambda d: d["pos"])
    merged = []
    for d in out:
        if merged and d["pos"] - merged[-1]["pos"] <= 25.0:
            if d["len"] > merged[-1]["len"]:
                merged[-1] = d
        else:
            merged.append(d)
    return merged


def _cluster_positions(vals, gap: float):
    """相邻间隔 <= gap 的坐标合并为一个（取均值），返回升序列表。"""
    if not vals:
        return []
    vals = sorted(vals)
    groups = [[vals[0]]]
    for v in vals[1:]:
        if v - groups[-1][-1] <= gap:
            groups[-1].append(v)
        else:
            groups.append([v])
    return [sum(g) / len(g) for g in groups]


def _fill_missing_lines(lines, span: float, max_gap_ratio: float = 1.9):
    """按中位间距等分补齐漏检的网格线（浅色线/断线常见）。

    lines: 升序坐标；span: 区间总长（仅用于限制插入数量）。
    """
    if len(lines) < 2:
        return list(lines)
    gaps = [b - a for a, b in zip(lines, lines[1:])]
    gaps_sorted = sorted(gaps)
    med = gaps_sorted[len(gaps_sorted) // 2]
    if med <= 0:
        return list(lines)
    out = [lines[0]]
    for a, b in zip(lines, lines[1:]):
        gap = b - a
        if gap > med * max_gap_ratio and gap < span:
            n = int(round(gap / med))
            for k in range(1, n):
                out.append(a + gap * k / n)
        out.append(b)
    return out


def _line_gap_regularity(lines) -> float:
    """线间距规整度：与中位距偏差 <=30% 的间距占比。

    漏检造成的行合并会让间距混入整倍行高（实测 36/78/111 混杂，规整度 0.83）；
    完整网格或整体等距漏检的间距都是均匀的（1.00）。用于判断纸张候选网格
    自身是否漏行，线太少时无从判断，返回 1.0 不干预。
    """
    gaps = [b - a for a, b in zip(lines, lines[1:])]
    if len(gaps) < 3:
        return 1.0
    med = sorted(gaps)[len(gaps) // 2]
    if med <= 0:
        return 1.0
    return sum(1 for g in gaps if abs(g - med) <= 0.3 * med) / len(gaps)


def _keep_dense_run(lines, tol: float = 0.6):
    """保留最长的"等距连续段"，丢掉离群的稀疏线。

    warp 区域比表格大一圈（为了带上标题/脚注），标题与签名区的文字基线可能
    被误判成网格线；这些离群线与表格主体的行高差异明显。表格主体是最长的
    一段"相邻间距都接近中位行高"的连续线串，只保留它。
    """
    if len(lines) < 3:
        return list(lines)
    gaps = [b - a for a, b in zip(lines, lines[1:])]
    med = sorted(gaps)[len(gaps) // 2]
    if med <= 0:
        return list(lines)
    best = (0, 1)
    start = 0
    for i, g in enumerate(gaps):
        if abs(g - med) > tol * med:
            if i + 1 - start > best[1] - best[0]:
                best = (start, i + 1)
            start = i + 1
    if len(gaps) + 1 - start > best[1] - best[0]:
        best = (start, len(gaps) + 1)
    return lines[best[0]:best[1]]


def _line_skew_deg(bw) -> float:
    """二值图上"长直线"的主倾角（度，归一化到 (-45,45]）。

    只认长线（>=0.2*宽、缝隙小）：表格框线动辄半张纸长，文字行/纸纹连不成
    这么长的段，天然被排除。大倾斜俯拍时线整体同倾，中位数即表格倾角；
    摆正后长横线水平，返回值接近 0。既用于旋转摆正，也用于切格前质检
    （明显没摆正就拒绝，斜线只会捞出文字行假线——实测 5 条等距假横线 +
    11 条乱距竖线切成 4 行 x 10 列）。
    """
    import cv2
    import math

    W = bw.shape[1]
    lines = cv2.HoughLinesP(255 - bw, 1, math.pi / 180, threshold=100,
                            minLineLength=max(80.0, 0.2 * W), maxLineGap=20)
    if lines is None:
        return 0.0
    items = []
    for x1, y1, x2, y2 in lines[:, 0]:
        a = math.degrees(math.atan2(y2 - y1, x2 - x1))
        while a > 45:
            a -= 90
        while a <= -45:
            a += 90
        items.append((a, math.hypot(x2 - x1, y2 - y1)))
    if not items:
        return 0.0
    # 只取最长的 1/3：表格框线动辄贯通半张纸，桌沿/折痕只有零星几条
    # （实测混入后全体中位数被拉偏一半）。取长段的中位数抗单条异常。
    items.sort(key=lambda t: -t[1])
    top = items[:max(3, len(items) // 3)]
    angs = sorted(a for a, _ in top)
    return angs[len(angs) // 2]


def _angle_line_score(bw_small, ang: float, w: int, h: int) -> int:
    """二值图旋转 ang 度后的长横/竖线计数（摆正时最大）。

    在已二值化的小图上用 INTER_NEAREST 旋转（毫秒级），比每角重跑
    CLAHE+自适应阈值快一个量级；形态学核是轴的，斜线会被打碎成短段，
    只有摆正角才能检出贯通长线——这正是我们要的信号。
    """
    import cv2

    M = cv2.getRotationMatrix2D((w / 2.0, h / 2.0), ang, 1.0)
    rot = cv2.warpAffine(bw_small, M, (w, h), flags=cv2.INTER_NEAREST,
                         borderValue=0)
    hm = _line_mask(rot, True, w)
    vm = _line_mask(rot, False, h)
    hs = _line_runs(hm, True)
    vs = _line_runs(vm, False)
    nh = len([s for s in hs if s["len"] >= 0.3 * w])
    nv = len([s for s in vs if s["len"] >= 0.15 * h])
    return nh * 2 + nv


def _scan_best_angle(img, span: float = 15.0):
    """扫描旋转角，返回 (best_angle, best_score, second_score)。

    先粗扫（步长 2°）再细扫（步长 0.5°）。第二高分用于判断"摆正是否有
    明显收益"：若最佳角只比 0° 好一点，说明表格本就接近轴对齐，不必旋转
    （旋转必然引入插值模糊与白边）。
    """
    import cv2

    h, w = img.shape[:2]
    s = 700.0 / max(h, w)
    if s < 1.0:
        small = cv2.resize(img, (max(1, int(w * s)), max(1, int(h * s))),
                           interpolation=cv2.INTER_AREA)
    else:
        small = img
    sh, sw = small.shape[:2]
    g = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8)).apply(
        cv2.cvtColor(small, cv2.COLOR_BGR2GRAY))
    bw = cv2.adaptiveThreshold(g, 255, cv2.ADAPTIVE_THRESH_MEAN_C,
                               cv2.THRESH_BINARY_INV, 31, 20)
    scores = {}
    for a10 in range(int(-span * 5), int(span * 5) + 1, 10):   # 粗扫步长 2°
        a = a10 / 5.0
        scores[a] = _angle_line_score(bw, a, sw, sh)
    best = max(scores, key=scores.get)
    for a10 in range(int((best - 2) * 10), int((best + 2) * 10) + 1, 5):  # 细扫 0.5°
        a = a10 / 10.0
        scores[a] = _angle_line_score(bw, a, sw, sh)
    best = max(scores, key=scores.get)
    ranked = sorted(scores.values(), reverse=True)
    return best, scores[best], (ranked[1] if len(ranked) > 1 else 0)


def _grid_line_families(bw):
    """Hough 长线段聚成网格线族，返回 (hlines, vlines)。

    hlines: [(angle, y_at_cx)] 按 y 升序（近水平的表格横线）
    vlines: [(angle, x_at_cy)] 按 x 升序（近竖直的表格竖线）
    只保留 >=2 段支持的线（单段多为纸纹/字迹），抗离群。
    梯形畸变的俯拍上网格线整体同倾，按角度分族后各自聚成平行线组，
    这是四角透视矫正的输入。
    """
    import math
    from collections import defaultdict

    import cv2

    H, W = bw.shape[:2]
    min_len = max(120.0, 0.18 * min(W, H))
    segs = cv2.HoughLinesP(255 - bw, 1, math.pi / 180, threshold=60,
                           minLineLength=min_len, maxLineGap=20)
    if segs is None:
        return [], []
    cx, cy = W / 2.0, H / 2.0
    hor, ver = [], []
    for x1, y1, x2, y2 in segs[:, 0]:
        dx, dy = float(x2 - x1), float(y2 - y1)
        if dx == 0.0 and dy == 0.0:
            continue
        a = math.degrees(math.atan2(dy, dx))
        while a > 45.0:
            a -= 90.0
        while a <= -45.0:
            a += 90.0
        mx, my = (x1 + x2) / 2.0, (y1 + y2) / 2.0
        if abs(a) <= 25.0:
            # 近水平：外推到 x=cx 处的 y
            hor.append((a, my + math.tan(math.radians(a)) * (cx - mx)))
        else:
            # 近竖直：换成"与竖直方向的夹角"，外推到 y=cy 处的 x
            va = a - 90.0 if a > 0 else a + 90.0
            ver.append((va, mx + math.tan(math.radians(va)) * (cy - my)))

    def cluster(items):
        by_ang = defaultdict(list)
        for a, c in items:
            by_ang[int(round(a / 4.0))].append((a, c))
        out = []
        for group in by_ang.values():
            group.sort(key=lambda t: t[1])
            cur = [group[0]]
            for a, c in group[1:]:
                if c - cur[-1][1] <= 40.0:
                    cur.append((a, c))
                else:
                    if len(cur) >= 2:
                        out.append((sum(x[0] for x in cur) / len(cur),
                                    sum(x[1] for x in cur) / len(cur)))
                    cur = [(a, c)]
            if len(cur) >= 2:
                out.append((sum(x[0] for x in cur) / len(cur),
                            sum(x[1] for x in cur) / len(cur)))
        out.sort(key=lambda t: t[1])
        return out

    return cluster(hor), cluster(ver)


def _intersect_hv(h, v, W, H):
    """横线 (ah, yh) 与竖线 (av, xv) 的交点；平行/退化返回 None。"""
    import math

    ah, yh = h
    av, xv = v
    px, py = W / 2.0, yh
    dh = (math.cos(math.radians(ah)), math.sin(math.radians(ah)))
    ang_v = math.radians(av + 90.0)
    dv = (math.cos(ang_v), math.sin(ang_v))
    qx, qy = xv, H / 2.0
    det = dh[0] * (-dv[1]) - (-dv[0]) * dh[1]
    if abs(det) < 1e-6:
        return None
    rx, ry = qx - px, qy - py
    t = (rx * (-dv[1]) - (-dv[0]) * ry) / det
    return (px + t * dh[0], py + t * dh[1])


def _table_quad_from_lines(img):
    """由网格线族估计表格四角（原图坐标 4x2 float32），失败返回 None。

    取最上/最下横线与最左/最右竖线，两两求交得四角。极端线取到纸边也没关系：
    warp 后 _grid_on_image 会重新按长线裁到表格区域。
    """
    import cv2
    import numpy as np

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    gray = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8)).apply(gray)
    bw = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C,
                               cv2.THRESH_BINARY_INV, 31, 12)
    H, W = bw.shape[:2]
    hor, ver = _grid_line_families(bw)
    if len(hor) < 2 or len(ver) < 2:
        return None
    top, bot = hor[0], hor[-1]
    lef, rig = ver[0], ver[-1]
    if bot[1] - top[1] < 60 or rig[1] - lef[1] < 60:
        return None
    pts = []
    for h, v in ((top, lef), (top, rig), (bot, rig), (bot, lef)):
        p = _intersect_hv(h, v, W, H)
        if p is None or not (np.isfinite(p[0]) and np.isfinite(p[1])):
            return None
        pts.append(p)
    quad = np.array(pts, dtype=np.float32)
    # 合理性：四角必须都在图内（留 10% 余量），且四边形不能退化成细条
    if (quad[:, 0] < -0.1 * W).any() or (quad[:, 0] > 1.1 * W).any() \
            or (quad[:, 1] < -0.1 * H).any() or (quad[:, 1] > 1.1 * H).any():
        return None
    sides = [float(np.linalg.norm(quad[i] - quad[(i + 1) % 4])) for i in range(4)]
    if min(sides) < 0.2 * max(sides) or min(sides) < 60:
        return None
    return quad


def image_quality_metrics(img) -> dict:
    """窗口化模糊度 + 高光占比（只读质量元数据，不改识别）。

    把图切成 96x96 小块（步长 48），跳过空白块，取拉普拉斯方差最高的
    20% 块的中位数。大片留白不会再把整图方差拉低，更接近人眼对
    “有字区域是否清楚”的判断。高光占比按全图亮度 >235 像素比例。
    """
    import cv2

    try:
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        h, w = gray.shape[:2]
        tile, step = 96, 48
        vals = []
        if h >= tile and w >= tile:
            for y in range(0, h - tile + 1, step):
                for x in range(0, w - tile + 1, step):
                    t = gray[y:y + tile, x:x + tile]
                    if float(t.mean()) < 10 or float(t.std()) < 5:
                        continue
                    vals.append(float(cv2.Laplacian(t, cv2.CV_64F).var()))
        vals.sort(reverse=True)
        k = max(8, int(len(vals) * 0.2))
        if vals:
            blur = float(np.median(vals[:k]))
        else:
            blur = float(cv2.Laplacian(gray, cv2.CV_64F).var())
        return {"blur_var": round(blur, 1),
                "highlight_ratio": round(float((gray > 235).mean()), 4)}
    except Exception:
        return {"blur_var": None, "highlight_ratio": None}


_DESKEW_MAX_DEG = 7.0
# 残余倾角反斜阈值；默认关闭（99）：实测对"测试图片1"这类表格把表头行整体
# 推歪，16 列还原退化成乱码；0683 的收益暂不抵消该回归。调优后可用
# IMG2DOC_DESKEW_MIN_DEG=0.6 重新启用，诊断/回滚仍可设 99 关闭。
_DESKEW_MIN_DEG = float(os.environ.get("IMG2DOC_DESKEW_MIN_DEG", "99"))


def _residual_skew_deg(warped) -> float:
    """表格残余倾斜角（度，向右下为正）。只统计近水平的长线段。

    矩形拉伸只吸收透视缩放、不吸收旋转：纸面相对取景框摆歪时，warp 上的
    行线仍整体倾斜 1-3°。轴对齐的横线形态核要求一个核长内不偏行，斜线会被
    打碎或糊成宽带；相邻行线投影粘连后只剩最平的几条能检出，等距补线再按
    假中位行距编出一整套假行（实测 0683：真实行距 47px 被补成 73.7px，一格
    挤进两个物理行，序号/单号整体错行）。这里先用长横线的主倾角把残余旋转
    量出来，供 _deskew_warp 反斜。
    """
    import cv2
    import numpy as np

    h, w = warped.shape[:2]
    s = min(1.0, 900.0 / float(max(h, w)))
    small = warped if s >= 1.0 else cv2.resize(
        warped, (max(1, int(w * s)), max(1, int(h * s))),
        interpolation=cv2.INTER_AREA)
    gray = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8)).apply(
        cv2.cvtColor(small, cv2.COLOR_BGR2GRAY))
    bw = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C,
                               cv2.THRESH_BINARY_INV, 31, 12)
    sh, sw = bw.shape[:2]
    segs = cv2.HoughLinesP(bw, 1, np.pi / 360, threshold=50,
                           minLineLength=max(60.0, 0.25 * sw), maxLineGap=18)
    if segs is None:
        return 0.0
    items = []
    for x1, y1, x2, y2 in segs[:, 0]:
        a = float(np.degrees(np.arctan2(y2 - y1, x2 - x1)))
        while a > 45:
            a -= 90
        while a <= -45:
            a += 90
        if abs(a) <= _DESKEW_MAX_DEG:
            items.append((a, float(np.hypot(x2 - x1, y2 - y1))))
    if len(items) < 5:                  # 证据不足（无表格线/手写短笔画）不摆正
        return 0.0
    items.sort()
    total = sum(wt for _, wt in items)
    acc = 0.0
    med = items[len(items) // 2][0]
    for a, wt in items:                 # 按长度加权中位，长行线权重更大
        acc += wt
        if acc >= total / 2.0:
            med = a
            break
    return med if abs(med) >= _DESKEW_MIN_DEG else 0.0


def _deskew_warp(warped):
    """按残余倾角旋转摆正 warp 图，返回 (图, 3x3 矩阵, 角度)。

    不摆正时返回单位阵与 0；矩阵把摆正前的 warp 坐标映射到摆正后坐标，
    调用方把它左乘到 quad_matrix 上即可保持 to_original 逆映射正确。
    """
    import cv2
    import numpy as np

    angle = _residual_skew_deg(warped)
    if abs(angle) < _DESKEW_MIN_DEG:
        return warped, np.eye(3, dtype=np.float64), 0.0
    h, w = warped.shape[:2]
    m = cv2.getRotationMatrix2D((w / 2.0, h / 2.0), angle, 1.0)
    cos, sin = abs(m[0, 0]), abs(m[0, 1])
    nw = int(round(h * sin + w * cos))
    nh = int(round(h * cos + w * sin))
    m[0, 2] += nw / 2.0 - w / 2.0      # 扩画布，避免旋转裁掉表格边缘
    m[1, 2] += nh / 2.0 - h / 2.0
    out = cv2.warpAffine(warped, m, (nw, nh), flags=cv2.INTER_LINEAR,
                         borderValue=(255, 255, 255))
    m3 = np.array([[m[0, 0], m[0, 1], m[0, 2]],
                   [m[1, 0], m[1, 1], m[1, 2]],
                   [0.0, 0.0, 1.0]], dtype=np.float64)
    return out, m3, angle


def _grid_metrics(warped, target_long_edge: int = 2200,
                  max_rows: int = 90, max_cols: int = 26, scale: float | None = None,
                  gap_ratio: float = 1.9, h_ratio: float = 0.55, bin_c: int = 12):
    """在（已摆正的）图上检测网格线并切格；没摆正/不规整返回 None。

    返回 dict 与 detect_table_grid 的成功返回同构（含 warp 临时图路径）。
    scale 为调用方算好的放大倍数（按表格区域算，见 _grid_on_image）；None 时
    按本图尺寸兜底。
    """
    import cv2

    wh, ww = warped.shape[:2]
    if max(ww, wh) < 200:
        return None
    # 图像质量指标（窗口化，避开大片留白）：只读元数据，不改识别。
    qmetrics = image_quality_metrics(warped)
    blur_var = qmetrics.get("blur_var")
    highlight_ratio = qmetrics.get("highlight_ratio")
    # 把表格拉大到 target_long_edge，小字变清晰；上限 4000 防内存爆炸
    if scale is None:
        scale = min(target_long_edge / float(max(ww, wh)), 4000.0 / float(max(ww, wh)))
    scale = max(scale, 1.0)
    if abs(scale - 1.0) > 1e-6:
        warped = cv2.resize(warped, (max(1, int(round(ww * scale))),
                                     max(1, int(round(wh * scale)))),
                            interpolation=cv2.INTER_CUBIC)
        ww, wh = warped.shape[1], warped.shape[0]

    # 在摆正后的图上重新定位网格（透视已被吸收，线应是直的）。
    # C=12 是实测最优：放到 20 会漏检淡线（实测 23 条横线掉到 17 条、竖线
    # 整体偏移）；旋转插值导致的糊线由 detect_table_grid 的角度扫描兜底——
    # 摆正后才进这里，不需要为斜图放宽阈值。
    wgray = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8)).apply(
        cv2.cvtColor(warped, cv2.COLOR_BGR2GRAY))
    wbw = cv2.adaptiveThreshold(wgray, 255, cv2.ADAPTIVE_THRESH_MEAN_C,
                                cv2.THRESH_BINARY_INV, 31, bin_c)
    hmask = _line_mask(wbw, True, ww)
    hsegs = _line_runs(hmask, True)
    hy = _cluster_positions([s["pos"] for s in hsegs if s["len"] >= h_ratio * ww],
                            gap=max(6, ww * 0.01))
    # 顺序很重要：先按中位行高等分补齐漏检的横线（浅色线/断线常见），
    # 再去掉离群的稀疏线（标题/签名区的文字基线）。反过来的话，漏检造成的大
    # 间距会被当成分割点，把表格主体切成几段只保留一段（实测 25 行只剩 12 行）。
    # 注意这里不能给 hy 补 0/wh-1 假边界：那会造出"首条真实线 - 0"的大间距，
    # 被 _fill_missing_lines 等分插出一串假行（实测多出 5 行）。
    hy_raw = len(hy)
    hy = _fill_missing_lines(hy, wh, max_gap_ratio=gap_ratio)
    hy_fill = len(hy)
    hy = _keep_dense_run(hy)
    # 竖线预跑一遍拿表格 y 范围：列分隔线贯通整表，其垂直跨度即表格真实上下
    # 沿，用作横线外推的边界（否则向上铺不进表头、向下铺进桌影垃圾带）。
    _th0 = max(1.0, (hy[-1] - hy[0]) if len(hy) >= 2 else float(wh))
    _vm = _line_mask(wbw, False, int(_th0))
    _vs = [s for s in _line_runs(_vm, False) if s["len"] >= 0.35 * _th0]
    if _vs:
        # 分位抗离群：表外竖带（纸边/折痕/邻页阴影）会从 y=0 起长段，
        # min/max 会把表格上下沿撑到整图，横线外推随之铺进摘要区插假行。
        starts = sorted(s["start"] for s in _vs)
        ends = sorted(s["start"] + s["len"] for s in _vs)
        k = max(0, len(starts) // 10)
        vy0 = starts[k]
        vy1 = ends[len(ends) - 1 - k]
    else:
        vy0, vy1 = 0.0, float(wh)
    # 表格外的假横线（摘要/标题文字行被形态学核捞成线，且与行高等距时
    # _keep_dense_run 丢不掉）会把表头顶到摘要区、等分插值再补假行。
    # 竖线跨度即表格真实上下沿，用它夹掉跨度外的横线。
    if _vs and len(hy) >= 3:
        hg0 = [b - a for a, b in zip(hy, hy[1:])]
        tol = 0.6 * sorted(hg0)[len(hg0) // 2]
        hy = [y for y in hy if vy0 - tol <= y <= vy1 + tol]
    # 表格 x 跨度（必须在横线外推之前算）：取检出横线线段的 min(start) 与
    # max(end)。不能用"行带暗像素左右端点的中位数"——表格上方的手写标题是
    # 居中文字，会把左边界从 0 抬到 49，进而把真正的左边界当纸边裁掉。
    # 断线（左侧碎成短段）只会让个别 start 偏大，min/max 不受影响。
    # 只统计表格 y 跨度内的横线：摘要文字行横跨整页，会把 hx0/hx1 撑到
    # 页边，补首尾时多切出左右两个空列（实测 16 列表变 18 列）。
    hspan = [s for s in hsegs if s["len"] >= 0.15 * ww
             and vy0 - 2 <= s["pos"] <= vy1 + 2]
    hx0 = min((s["start"] for s in hspan), default=0.0)
    hx1 = max((s["start"] + s["len"] for s in hspan), default=float(ww))
    # 横线外推：阴影/淡线会让半张表的行线整段漏检（实测 11 行表只检出 7 条），
    # 漏检区间没有"内部大间距"可插，_fill_missing_lines 无能为力。按中位行距
    # 向上下外推，外推行带内须有文字（内容门控，阈值放到 0.006：表头字少、
    # 实测占比仅 ~1.9%，0.02 会把表头行整行漏掉），且不越过竖线的垂直跨度
    # （空白纸边/桌影自然停在表格外）。
    if len(hy) >= 3:
        hgaps = [b - a for a, b in zip(hy, hy[1:])]
        hmed = sorted(hgaps)[len(hgaps) // 2]
        if hmed > 4:
            x_lo, x_hi = int(0.05 * ww), max(int(0.05 * ww) + 1, int(0.95 * ww))
            y_top = max(8.0, vy0 - 0.5 * hmed)
            y_bot = min(float(wh) - 8, vy1 + 0.5 * hmed)

            def _band_ink(y1, y2):
                a, b = int(max(0, y1)), int(min(wh, y2))
                if b - a < 4:
                    return 0.0
                return float((wbw[a:b, x_lo:x_hi] > 0).mean())

            for _ in range(80):
                nxt = hy[-1] + hmed
                if nxt > y_bot or _band_ink(hy[-1], nxt) < 0.006:
                    break
                hy.append(round(nxt, 1))
            for _ in range(80):
                nxt = hy[0] - hmed
                if nxt < y_top or _band_ink(nxt, hy[0]) < 0.006:
                    break
                hy.insert(0, round(nxt, 1))
            # 末行被 warp 下沿裁断（底框线落在图像外）时，外推会被 y_bot 拦住，
            # 停在上一条线，整行数据丢失（实测 12 行表只切出 11 行，漏最后一行
            # YT1234567890123）。图像下沿与上一条线相距接近一个行高、且该行带内
            # 确有文字时，用下沿做虚拟底边把末行圈回来；空白 pad/脚注区墨迹
            # 不足，不会误加。
            if (len(hy) >= 3
                    and 0.6 * hmed <= wh - hy[-1] <= 1.6 * hmed
                    and _band_ink(hy[-1], float(wh)) >= 0.004):
                hy.append(float(wh))
    # 竖线核高与长度阈值都按"表格主体高度"算，不能用整图高度：warp 比表格高
    # 一圈（为了带上标题/脚注），按整图高度设阈值会把 683px 的真实列线误杀
    # （实收|短少数之间的线），导致最后两列粘在一起。
    # 长度阈值 0.35 兼顾"够长的真线"与"右侧两条短列线"（实收/短少数只有半表高）；
    # 取 0.55 会连它们也一起漏掉，实测 25 行表最后两列粘成一格。
    # 用"外推前的检出跨度"而不是外推后的：外推会把表头上方标题带算进来、
    # table_h 从 900 撑到 1221，长度阈值随之变严，真实列线会被整条滤掉
    # （实测 9 条竖线只剩 4 条）。
    table_h = _th0
    vmask = _line_mask(wbw, False, int(table_h))
    vsegs = _line_runs(vmask, False)
    # 竖线裁剪到表格 x 跨度：纸边/折痕在表格外也能形成长竖线，横线只跨表格
    # 真实宽度，之外的竖线一律视为纸边丢弃。
    vsegs = [s for s in vsegs if hx0 - 8 <= s["pos"] <= hx1 + 8]
    # 聚类间距放到 25px：双重边框/折痕会产生 13-15px 的近距重复线（实测
    # 266/279、2184/2199），不合并会把一列切成两个空列。
    vx = _cluster_positions([s["pos"] for s in vsegs if s["len"] >= 0.35 * table_h],
                            gap=25.0)
    # 列边界补首尾：用表格真实外沿，不是 warp 边界（warp 右侧有白边，补
    # ww-1 会多切一个空列）。附近已有线（间距 <= pad）就不补，否则会把
    # 2184.5 与右外沿 2197 并列成两个空列。
    pad = 4
    if not vx or vx[0] > hx0 + pad:
        vx.insert(0, hx0)
    if not vx or hx1 - vx[-1] > pad:
        vx.append(hx1)

    if len(hy) < 4 or len(vx) < 3:
        return None
    if len(hy) - 1 > max_rows or len(vx) - 1 > max_cols:
        return None

    def cells_of(ylo, yhi):
        out = []
        for a, b in zip(ylo, yhi):
            for c, d in zip(vx, vx[1:]):
                if b - a < 6 or d - c < 6:
                    continue
                out.append([int(round(c)), int(round(a)), int(round(d)), int(round(b))])
        return out

    header = cells_of(hy[:2], hy[1:3]) if len(hy) >= 3 else []
    body = cells_of(hy[1:-1], hy[2:])

    fd, tmp = tempfile.mkstemp(prefix="img2xlsx_grid_", suffix=".png")
    os.close(fd)
    cv2.imwrite(tmp, warped)
    return {
        "grid": True,
        "warp_path": tmp,
        "warp_w": int(ww),
        "warp_h": int(wh),
        "hlines": [round(v, 1) for v in hy],
        "vlines": [round(v, 1) for v in vx],
        "header": header,
        "cells": body,
        "n_rows": max(0, len(hy) - 2),
        "_h_raw": hy_raw, "_h_fill": hy_fill,
        "blur_var": blur_var, "highlight_ratio": highlight_ratio,
        "n_cols": max(0, len(vx) - 1),
    }


def _keystone_rewarp(warped, m, scale, res, target_long_edge, max_rows, max_cols):
    """矩形拉伸后再吸收一次梯形畸变；切格数严格变多才采用。

    一次 warp 若走矩形拉伸（Hough 线族被阴影/糊线带偏时四角候选缺席），
    竖线左半斜、右半直，轴对齐核只检出右半 → 列坍塌。这里用抗斜竖线的
    最外两条 + 表格上下沿求四角再透视一次，线真正摆正后切格恢复全列。
    轴对齐图斜率≈0 不进入；切格数不增则保持原结果，行为与旧版一致。
    """
    import cv2
    import numpy as np

    hy = res.get("hlines") or []
    if len(hy) < 4:
        return None
    # res 的 hlines 是放大后坐标，warped 还是放大前：先统一到放大后坐标系，
    # 否则窗口越界、抗斜竖线空手而归（实测修复路径静默失效）。
    if abs(scale - 1.0) > 1e-6:
        hh, ww = warped.shape[:2]
        warped = cv2.resize(warped, (max(1, int(round(ww * scale))),
                                     max(1, int(round(hh * scale)))),
                            interpolation=cv2.INTER_CUBIC)
    y0, y1 = float(hy[0]), float(hy[-1])
    gray = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8)).apply(
        cv2.cvtColor(warped, cv2.COLOR_BGR2GRAY))
    bw = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C,
                               cv2.THRESH_BINARY_INV, 31, 12)
    sl = _slant_vline_runs(bw, y0, y1, 0.35 * (y1 - y0))
    if len(sl) < 6:
        bw8 = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C,
                                    cv2.THRESH_BINARY_INV, 31, 8)
        sl = _slant_vline_runs(bw8, y0, y1, 0.35 * (y1 - y0))
    if len(sl) < 6 or max(abs(d["k"]) for d in sl) < 0.006:
        return None
    left, right = sl[0], sl[-1]
    cy = (y0 + y1) / 2.0

    def xat(d, y):
        return d["pos"] + d["k"] * (y - cy)

    quad = np.array([[xat(left, y0), y0], [xat(right, y0), y0],
                     [xat(right, y1), y1], [xat(left, y1), y1]], dtype=np.float32)
    w2, m2 = _quad_warp(warped, quad)
    if w2 is None or min(w2.shape[:2]) < 150:
        return None
    s2 = max(1.0, min(target_long_edge / float(max(w2.shape[:2])),
                      4000.0 / max(w2.shape[:2])))
    res2 = _grid_metrics(w2, target_long_edge=target_long_edge,
                         max_rows=max_rows, max_cols=max_cols, scale=s2)
    if not res2:
        return None
    n1 = len(res["hlines"]) * len(res["vlines"])
    n2 = len(res2["hlines"]) * len(res2["vlines"])
    if n2 <= n1:
        try:
            Path(res2["warp_path"]).unlink(missing_ok=True)
        except OSError:
            pass
        return None
    try:
        Path(res["warp_path"]).unlink(missing_ok=True)
    except OSError:
        pass
    res2["quad_matrix"] = (m2 @ np.array(m, dtype=np.float64)).tolist()
    res2["quad_scale"] = float(scale * s2)
    return res2


def _grid_on_image(img, target_long_edge: int = 2200,
                   max_rows: int = 90, max_cols: int = 26,
                   paper_box: tuple | None = None):
    """单张图跑网格检测：长线定位表格区域 -> 矩形拉伸（左扩 2 列）-> 切格。

    paper_box=(x0,y0,x1,y1) 为纸张外接框（原图坐标）。实拍照片里键盘键帽、
    桌面边缘会形成短横/竖线，矩形拉伸后可能把纸张外的区域卷进 warp（实测
    一张倾斜运单被切成 57x9，输出首行成了键盘字母）。给了纸张框就只接受
    框内的长线，并把拉伸区域夹在纸张内；纸张检测失败时传 None，行为与旧版
    完全一致。

    成功返回 _grid_metrics 结果，任一环节不合格返回 None。
    """
    import cv2

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    gray = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8)).apply(gray)
    H, W = gray.shape

    # 长线阈值两级放宽：先严（防桌面/纸边混入），不够再松；对比阈值 C 也两级
    # （12 -> 20）：旋转/缩放插值会把细网格线糊淡，严阈值下一条长线都不剩。
    for C in (12, 20):
        bw = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C,
                                   cv2.THRESH_BINARY_INV, 31, C)
        for hr, vr in ((0.30, 0.15), (0.15, 0.08)):
            hmask = _line_mask(bw, True, W)
            vmask = _line_mask(bw, False, H)
            hlong = [s for s in _line_runs(hmask, True) if s["len"] >= hr * W]
            vlong = [s for s in _line_runs(vmask, False) if s["len"] >= vr * H]
            if paper_box is not None:
                px0, py0, px1, py1 = paper_box
                hlong = [s for s in hlong if py0 <= s["pos"] <= py1]
                vlong = [s for s in vlong if px0 <= s["pos"] <= px1]
            # 轴对齐核漏掉斜列线时（梯形畸变残斜），用抗斜竖线补左边界，
            # 否则 x0 分位落在内列线上、左扩 2 列圈不回被裁的整列。
            if len(vlong) < 8 and len(hlong) >= 3:
                sy0 = min(s["pos"] for s in hlong)
                sy1 = max(s["pos"] for s in hlong)
                sl = _slant_vline_runs(bw, sy0, sy1, vr * H)
                if len(sl) < 6:
                    # 阴影区淡线 C=12 提取不全，放宽到 8 再试一次
                    bw8 = cv2.adaptiveThreshold(
                        gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C,
                        cv2.THRESH_BINARY_INV, 31, 8)
                    sl8 = _slant_vline_runs(bw8, sy0, sy1, vr * H)
                    sl = sorted(sl + sl8, key=lambda d: d["pos"])
                    merged = []
                    for d in sl:
                        if merged and d["pos"] - merged[-1]["pos"] <= 25.0:
                            if d["len"] > merged[-1]["len"]:
                                merged[-1] = d
                        else:
                            merged.append(d)
                    sl = merged
                if paper_box is not None:
                    sl = [d for d in sl if px0 <= d["pos"] <= px1]
                vlong = vlong + sl
            if len(hlong) >= 3 and len(vlong) >= 2:
                break
        if len(hlong) >= 3 and len(vlong) >= 2:
            break
    if len(hlong) < 3 or len(vlong) < 2:
        return None

    ys = sorted(s["pos"] for s in hlong)
    xs = sorted(s["pos"] for s in vlong)
    # 用 5%/95% 分位抗离群（纸边、桌沿等更长的线）
    y0, y1 = ys[max(0, int(len(ys) * 0.05))], ys[min(len(ys) - 1, int(len(ys) * 0.95))]
    x0, x1 = xs[max(0, int(len(xs) * 0.05))], xs[min(len(xs) - 1, int(len(xs) * 0.95))]

    if paper_box is not None:
        px0, py0, px1, py1 = paper_box
        x0, x1 = max(x0, px0 + 2), min(x1, px1 - 2)
        y0, y1 = max(y0, py0 + 2), min(y1, py1 - 2)
    if x1 - x0 < 80 or y1 - y0 < 60:
        return None

    # 上下各扩 15%：表格上方的标题、下方的脚注/签名要一起进 warp 图，
    # 否则这些文字会被裁掉（精确复制模式要还原整页，不只是表格）。
    pad_y = 0.15 * (y1 - y0)
    y0e, y1e = max(0.0, y0 - pad_y), min(float(H - 1), y1 + pad_y)
    # 左侧也要扩：x0 来自检出竖线的 5% 分位，而表格左缘的窄列（序号列、
    # 日期列）竖线短、易漏检——漏检后 x0 直接落在内列线上，warp 会把整列
    # 裁在外面（实测 14 行表只剩 6 列，日期/序号整列消失）。用检出竖线的
    # 中位间距估计一个列宽、左侧扩 2 列把漏检的边缘列圈回 warp；右侧不扩：
    # 右缘是表格最粗的框线极少漏检（旧版右侧两列一直完整），扩了反而会把
    # 表格外沿的浅线圈进来多切出一个空列。白边对切格/OCR 无害（无框即空）。
    col_gaps = [b - a for a, b in zip(xs, xs[1:])]
    med_col = sorted(col_gaps)[len(col_gaps) // 2] if col_gaps else x1 - x0
    x0e = max(0.0, x0 - 2.0 * med_col)
    if paper_box is not None:
        # 纸张框外的表头标题/签名不会被裁掉（它们在纸内），但键盘、桌面、
        # 手指这些纸张外的内容必须挡在 warp 之外。
        px0, py0, px1, py1 = paper_box
        x0e, y0e = max(x0e, px0), max(y0e, py0)
        x1, y1e = min(x1, px1), min(y1e, py1)
    quad = np.array([[x0e, y0e], [x1, y0e], [x1, y1e], [x0e, y1e]], dtype=np.float32)
    # 放大倍数按"表格区域"算（不是 warp 尺寸）：warp 带了上下 15% pad 与左扩
    # 2 列，按它算会把缩放系数算小，warp 尺寸与坐标整体偏移（实测 photo2 的
    # hlines 从 [157.5,1213.5] 漂到 [180.5,1402.0]）。
    scale = min(target_long_edge / float(max(x1 - x0, y1 - y0)),
                4000.0 / max(H, W))
    scale = max(scale, 1.0)
    warped, m = _quad_warp(img, quad)
    if warped is None:
        return None
    res = _grid_metrics(warped, target_long_edge=target_long_edge,
                        max_rows=max_rows, max_cols=max_cols, scale=scale)
    if res:
        # 高亮回原图需要：warp 坐标 = scale * (m * 候选图坐标)，把换算要用的
        # 透视矩阵与缩放系数带回 detect_table_grid 组合成完整映射。
        res["quad_matrix"] = m.tolist()
        res["quad_scale"] = float(scale)
        better = _keystone_rewarp(warped, m, scale, res, target_long_edge,
                                  max_rows, max_cols)
        if better is not None:
            res = better
        else:
            # 残余倾斜只作为“本候选内部的备选”，不参与候选择优：若提前反斜，
            # 有斜行线的坏候选会多检出假行，在“行×列”得分上反超正确候选
            # （实测 815b 纸张候选因此从 44x17 变成 45x17 且序号粘进单号格）。
            # 这里先按原图切格，再由 detect_table_grid 在选定主候选后比较行距
            # 规整度决定是否换成反斜结果。
            skew = _residual_skew_deg(warped)
            if abs(skew) >= _DESKEW_MIN_DEG:
                dwarp, deskew_m, _ = _deskew_warp(warped)
                alt = _grid_metrics(dwarp, target_long_edge=target_long_edge,
                                    max_rows=max_rows, max_cols=max_cols,
                                    scale=scale)
                if alt:
                    alt["quad_matrix"] = (deskew_m @ np.asarray(
                        m, dtype=np.float64)).tolist()
                    alt["quad_scale"] = float(scale)
                    res["deskew_alt"] = alt
    return res


def detect_table_grid(path: str, target_long_edge: int = 2200,
                      max_rows: int = 90, max_cols: int = 26, rotate: float = 0):
    """检测表格网格并做透视矫正。

    返回 dict（含 warp 临时图路径与单元格矩形，均为矫正图坐标）：
      {"grid": true, "warp_path", "warp_w", "warp_h",
       "hlines": [...], "vlines": [...],
       "header": [[x1,y1,x2,y2], ...], "cells": [[...], ...]}
    未检测到可用网格时返回 {"grid": False, "reason": ...}。

    大倾斜俯拍（纸张斜在桌上）先旋转摆正再检测：矩形拉伸纠正不了旋转变换，
    warp 上的线还是斜的，横/竖线检测只能捞出文字行的假线（实测切成 4 行 x
    10 列）。摆正角用二值图角度扫描（_scan_best_angle）确定——形态学核是
    轴的，斜线被打碎成短段，只有摆正角能检出贯通长线，扫描分数即信号；
    纸张四边形（_find_doc_quad）在复杂背景下经常找不到，不可依赖，Hough
    估角会被桌沿/纸边带偏（实测 3° 的表格估成 19°），均不作为选角依据。
    原图与旋转图都跑，按切格数择优。rotate=90/270 表示整页先按该方向转正
    （由 OcrEngine 的方向校正判定，调用方不要另做判断）。
    """
    import cv2

    img = load_image(path)
    orig_h, orig_w = img.shape[:2]
    # img→原图坐标映射（90/270 旋转时把旋转图坐标换回原图）
    if int(rotate or 0) == 90:          # ROTATE_90_CLOCKWISE
        base_to_orig = np.array([[0, 1, 0], [-1, 0, orig_h - 1], [0, 0, 1]], dtype=np.float64)
    elif int(rotate or 0) == 270:       # ROTATE_90_COUNTERCLOCKWISE
        base_to_orig = np.array([[0, -1, orig_w - 1], [1, 0, 0], [0, 0, 1]], dtype=np.float64)
    else:
        base_to_orig = np.eye(3, dtype=np.float64)
    if int(rotate or 0) in (90, 270):
        img = cv2.rotate(
            img, cv2.ROTATE_90_CLOCKWISE if int(rotate) == 90
            else cv2.ROTATE_90_COUNTERCLOCKWISE,
        )

    # 纸张外接框：只用于约束原图候选（键帽/桌面等纸张外纹理集中在原图）。
    # 旋转/线族候选的坐标系不同，不做约束，避免把正确结果挡掉。
    paper_box = None
    paper = None
    try:
        paper = _find_doc_quad(cv2.cvtColor(img, cv2.COLOR_BGR2GRAY))
        if paper is not None:
            paper_box = (float(paper[:, 0].min()), float(paper[:, 1].min()),
                         float(paper[:, 0].max()), float(paper[:, 1].max()))
            log(f"网格检测：纸张约束 {paper_box}")
    except Exception as exc:  # noqa: BLE001 - 纸张检测失败不应影响原流程
        log(f"网格检测：纸张检测跳过（{exc}）")

    candidates = [(0.0, img)]
    # 每个候选图坐标 → img 坐标的逆变换（与 candidates 一一对应）
    cand_to_img = [np.eye(3, dtype=np.float64)]
    # 候选 0：网格线族四角透视矫正。俯拍梯形畸变（上宽下窄）下，任何矩形
    # 拉伸都修不掉——warp 后表头与数据行仍斜着排（实测同一列表头 cy=166、
    # 表尾 cy=562），切格会把表头挤到第 1 数据行、序号前几行整行丢失。
    # 由 Hough 长线段聚成横/竖两族、取最外线求交得四角，能真正吸收梯形畸变。
    quad = _table_quad_from_lines(img)
    if quad is not None:
        warped, m_quad = _quad_warp(img, quad)
        if warped is not None and min(warped.shape[:2]) >= 150:
            candidates.append((-1.0, warped))
            cand_to_img.append(np.linalg.inv(m_quad))
            log("网格检测：已用网格线族四角做透视矫正")
    ang, best_sc, second_sc = _scan_best_angle(img)
    # 摆正收益明显才旋转（避免给本就轴对齐的图引入插值模糊与白边）
    if abs(ang) >= 0.5 and best_sc > second_sc + 2:
        h, w = img.shape[:2]
        M = cv2.getRotationMatrix2D((w / 2.0, h / 2.0), ang, 1.0)
        rot = cv2.warpAffine(img, M, (w, h), flags=cv2.INTER_LINEAR,
                             borderValue=(255, 255, 255))
        candidates.append((ang, rot))
        m3 = np.array([[M[0, 0], M[0, 1], M[0, 2]],
                       [M[1, 0], M[1, 1], M[1, 2]],
                       [0, 0, 1]], dtype=np.float64)
        cand_to_img.append(np.linalg.inv(m3))
        log(f"网格检测：角度扫描得摆正角 {ang:.1f}°（长线 {second_sc} -> {best_sc}）")

    def box_in_candidate(box, cand_to_img):
        """把 img 坐标的纸张框变换到候选图坐标，供 _grid_on_image 裁剪。"""
        if box is None:
            return None
        inv = np.linalg.inv(cand_to_img)
        xs, ys = [], []
        for px, py in ((box[0], box[1]), (box[2], box[1]),
                       (box[2], box[3]), (box[0], box[3])):
            v = inv @ np.array([px, py, 1.0], dtype=np.float64)
            if abs(v[2]) < 1e-9:
                return None
            xs.append(v[0] / v[2])
            ys.append(v[1] / v[2])
        return (min(xs), min(ys), max(xs), max(ys))

    results = []
    for ci, (_a, cand) in enumerate(candidates):
        res = _grid_on_image(cand, target_long_edge=target_long_edge,
                             max_rows=max_rows, max_cols=max_cols,
                             paper_box=box_in_candidate(paper_box, cand_to_img[ci]))
        if res:
            log(f"网格候选 {ci}：{res.get('n_rows')} 行 × {res.get('n_cols')} 列，"
                f"warp {res.get('warp_w')}x{res.get('warp_h')}")
            results.append((ci, res))
    if not results:
        return {"grid": False, "reason": "网格不规整或倾斜未能摆正"}
    # 择优：真实网格切的格子数（行 x 列）远多于斜线假线切出的碎片
    # （实测倾斜未摆正只有 3 行 x 10 列、摆正后 11 行 x 7 列）；并列取
    # 靠前候选（原图优先，旋转图带白边不占优）。
    results.sort(key=lambda t: (-(len(t[1]["hlines"]) * len(t[1]["vlines"])), t[0]))
    ci, best = results[0]
    for _c, r in results[1:]:
        try:
            Path(r["warp_path"]).unlink(missing_ok=True)
            alt_r = r.get("deskew_alt")
            if alt_r:
                Path(alt_r["warp_path"]).unlink(missing_ok=True)
        except OSError:
            pass

    # 主候选选定后，若行距明显不规整（残余倾斜把多行糊成一行后等距补线），
    # 且按残余倾角反斜重切的行距明显更规整，则换用反斜结果。该比较只作用
    # 在已选中的候选上，不改变候选择优，避免坏候选因旋转多出假行而反超。
    alt = best.get("deskew_alt")
    if alt is not None:
        best.pop("deskew_alt", None)
        reg_b = _line_gap_regularity(best["hlines"])
        reg_a = _line_gap_regularity(alt["hlines"])
        if alt["n_cols"] >= best["n_cols"] and reg_a > reg_b + 0.08:
            log(f"网格检测：残余倾斜反斜重切（原 {best['n_rows']}x{best['n_cols']} "
                f"规整 {reg_b:.2f} → {alt['n_rows']}x{alt['n_cols']} 规整 {reg_a:.2f}）")
            try:
                Path(best["warp_path"]).unlink(missing_ok=True)
            except OSError:
                pass
            best = alt
        else:
            try:
                Path(alt["warp_path"]).unlink(missing_ok=True)
            except OSError:
                pass

    # 纸张候选：先按纸张四边形做透视矫正，再直接切格。复杂背景下矩形拉伸
    # 会把键盘/桌面卷进表格区域（实测 57x9，首行是键盘字母）；纸张矫正图
    # 只含纸面，网格规模回到 50 行量级。只在列数明显多于纸张候选时采用，
    # 避免把本来就正确的候选换掉。
    paper_best = None
    paper_inv = None
    if paper is not None:
        try:
            pimg, pm = _quad_warp(img, paper)
            if pimg is not None and min(pimg.shape[:2]) >= 150:
                ph_, pw_ = pimg.shape[:2]
                ps = max(1.0, min(target_long_edge / max(pw_, ph_),
                                  4000.0 / max(pw_, ph_)))
                # 纸张候选不反斜：整张纸 warp 里表格只占一小块，Hough 主倾角
                # 会被纸面折痕/正文行带偏（实测 0683 估出 -4.87°），摆正后正文
                # 行距反而被补成一套均匀假网格（53 行 x 3 列）冒充表格。
                pres = _grid_metrics(pimg, target_long_edge=target_long_edge,
                                     max_rows=max_rows, max_cols=max_cols,
                                     scale=ps, h_ratio=0.30, bin_c=8)
                if pres:
                    pres["quad_matrix"] = np.eye(3, dtype=np.float64).tolist()
                    pres["quad_scale"] = float(ps)
                    paper_best = pres
                    paper_inv = np.linalg.inv(pm)
                    log(f"纸张候选：{pres.get('n_rows')} 行 × {pres.get('n_cols')} 列")
        except Exception as exc:  # noqa: BLE001 - 纸张候选失败不影响原候选
            log(f"网格检测：纸张候选跳过（{exc}）")

    use_paper = False
    # 纸张候选自己也会漏行（淡线/弯纸让行线检出一半，实测 53 行 x 9 列被
    # 换成 28 行 x 5 列）：漏行后间距混入整倍行高、列数也偏少，再拿列数规则
    # 去否决更完整的候选，会把两个物理行挤进一格、单号粘连截断。间距不规整
    # 即自身漏行的证据，此时不允许否决；间距均匀的纸张候选保持原否决路径
    # （实测 669d4d7e 靠它走下游几何回退才读全）。
    # 行数尺度相近时不能再拿列数规则否决基础候选：纸张四边形常裁进表格
    # 左缘（实测 8a2b/43e3/815b/fc03/6709/6db5/ac0c 的纸张候选只剩 4-13 列，
    # 基础候选 17-18 列才是完整表格，序号列/表头被纸张候选整列裁掉）。
    # 只有两候选行数差到 1.5 倍以上（实测 669d4d7e：基础 15 行、纸张 6 行）
    # 才说明基础候选是假网格，此时纸张否决后走下游几何回退才读全。
    # 调试/回退：IMG2DOC_NO_PAPER=1 可完全禁用纸张候选否决。
    row_gap = (max(best["n_rows"], paper_best["n_rows"]) /
               max(1, min(best["n_rows"], paper_best["n_rows"]))) if paper_best else 0.0
    if (paper_best is not None
            and best["n_cols"] > paper_best["n_cols"] * 1.2
            and _line_gap_regularity(paper_best["hlines"]) >= 0.9
            and row_gap >= 1.5
            and os.environ.get("IMG2DOC_NO_PAPER") != "1"):
        old_n = f"{best['n_rows']} 行 × {best['n_cols']} 列"
        try:
            Path(best["warp_path"]).unlink(missing_ok=True)
        except OSError:
            pass
        best = paper_best
        use_paper = True
        log(f"网格检测：改用纸张候选（原 {old_n} → "
            f"{best['n_rows']} 行 × {best['n_cols']} 列）")
    elif paper_best is not None:
        try:
            Path(paper_best["warp_path"]).unlink(missing_ok=True)
        except OSError:
            pass

    # warp（已按 scale 放大）坐标 → 原图坐标：依次经过
    #   warp_final →候选图 → img →原图
    try:
        inv_quad = np.linalg.inv(np.array(best["quad_matrix"], dtype=np.float64))
        s = float(best.get("quad_scale") or 1.0)
        inv_quad[:, 0] /= s          # p_cand = inv(m) * (p_warp / scale)
        inv_quad[:, 1] /= s
        base = paper_inv if use_paper else cand_to_img[ci]
        best["to_original"] = (base_to_orig @ base @ inv_quad).tolist()
    except Exception as exc:  # noqa: BLE001 - 映射失败不影响识别主流程
        log(f"网格检测：原图坐标映射不可用（{exc}）")
        best["to_original"] = None
    best["orig_w"] = int(orig_w)
    best["orig_h"] = int(orig_h)
    best.pop("quad_matrix", None)
    best.pop("quad_scale", None)
    best["rotated"] = (not use_paper) and ci > 0
    best["rotate"] = int(rotate or 0)
    best["paper"] = bool(use_paper)
    best["paper_area_ratio"] = None
    best["paper_angle"] = None
    if paper is not None:
        try:
            area = abs(cv2.contourArea(paper.reshape(-1, 1, 2)))
            best["paper_area_ratio"] = round(float(area) / float(orig_w * orig_h), 4)
            q = _order_quad(paper)
            tl, tr = q[0], q[1]
            best["paper_angle"] = round(
                float(np.degrees(np.arctan2(tr[1] - tl[1], tr[0] - tl[0]))), 1)
        except Exception:
            pass
    return best


def _skew_from_boxes(boxes) -> float:
    """用检测框中心的 RANSAC 共线估计文本行倾角（度，y 向下坐标系）。

    同一文本行的一串框中心共线，跨行连线不共线且方向随机；取内点最多的
    连线方向即为行方向。注意一行通常只有 4-10 个框，内点阈值不能按框总数
    的百分比设（否则多列表格永远达不到），定为 4。
    """
    rng = np.random.default_rng(0)
    pts = np.array([[(b[0] + b[2]) / 2.0, (b[1] + b[3]) / 2.0] for b in boxes], dtype=np.float32)
    n = len(pts)
    if n < 8:
        return 0.0
    # n<=350 全对采样（约 6 万对，numpy 向量化下很快），更多则随机抽样
    if n <= 350:
        ii, jj = np.triu_indices(n, k=1)
    else:
        idx = rng.integers(0, n, (1200, 2))
        ii, jj = idx[:, 0], idx[:, 1]
    d = pts[jj] - pts[ii]
    # 只保留"从左指向右"的连线：dx>0 保证 atan2 得到真实倾角（dx<0 会给出 ±180° 附近的假角）
    keep = (d[:, 0] > 60) & (np.abs(d[:, 1]) < 0.25 * d[:, 0])
    ii, jj, d = ii[keep], jj[keep], d[keep]
    if len(d) == 0:
        return 0.0
    best_inl, best_ang = 0, 0.0
    for p, q, dd in zip(pts[ii], pts[jj], d):
        ang = math.degrees(math.atan2(dd[1], dd[0]))
        u = dd / (math.hypot(dd[0], dd[1]) or 1.0)
        dist = np.abs((pts - p) @ np.array([-u[1], u[0]]))
        inl = int((dist < 4.5).sum())
        if inl > best_inl:
            best_inl, best_ang = inl, ang
    if best_inl < 4:
        return 0.0
    return best_ang


def render_pdf(pdf_path: str, dpi: int = 300) -> list[dict]:
    """把 PDF 每页渲染为临时 PNG，返回 [{path, width, height}]（原图坐标）。

    dpi 硬下限 300：低于该值小字号单号会糊，识别率显著下降。
    长边硬上限 RENDER_MAX_EDGE：A4 页面 300 DPI 本就约 3500 px，超大页面
    （如工程图纸）若不设上限，视觉接口传输与 OCR 耗时都会失控。
    渲染走 pypdfium2（Chrome PDFium 引擎）+ cv2 存 PNG，不依赖 PIL。
    临时文件由调用方负责清理。
    """
    dpi = max(300, int(dpi or 300))
    scale = dpi / 72.0
    import pypdfium2 as pdfium
    import cv2

    doc = pdfium.PdfDocument(pdf_path)
    n_pages = len(doc)
    if n_pages <= 0:
        raise RuntimeError("PDF 无有效页面")
    if n_pages > 200:
        doc.close()
        raise RuntimeError(f"PDF 页数过多（{n_pages} 页），请控制在 200 页以内")

    out_dir = Path(tempfile.mkdtemp(prefix="img2xlsx_pdf_"))
    pages = []
    capped = False
    for i in range(n_pages):
        bitmap = doc[i].render(scale=scale)
        arr = bitmap.to_numpy()
        bitmap.close()
        img = cv2.cvtColor(arr, cv2.COLOR_RGBA2BGR) if arr.shape[2] == 4 else arr
        h, w = img.shape[:2]
        if max(h, w) > RENDER_MAX_EDGE:
            ratio = RENDER_MAX_EDGE / max(h, w)
            img = cv2.resize(img, (int(w * ratio), int(h * ratio)),
                             interpolation=cv2.INTER_AREA)
            capped = True
            h, w = img.shape[:2]
        out_path = out_dir / f"page_{i + 1:04d}.png"
        if not cv2.imwrite(str(out_path), img):
            raise RuntimeError(f"第 {i + 1} 页渲染失败")
        pages.append({"path": str(out_path), "width": int(w), "height": int(h)})
    doc.close()
    msg = f"PDF 渲染完成：{n_pages} 页，{dpi} DPI"
    if capped:
        msg += f"（超大页面已缩至长边 {RENDER_MAX_EDGE}px）"
    log(msg + f"，输出 {out_dir}")
    return pages


def default_models_dir() -> str:
    """模型目录优先级：环境变量 > 便携包自带 > 本机 venv 默认路径。

    便携包结构：肥鸭文档转换\\models\\PP-OCRv6\\{det,rec}\\*.xml
    ocr_worker.py 在 onefile 下位于 _MEIPASS，脚本同级即便携包根目录。
    同一层级下优先 PP-OCRv6（识别更准），缺失时回退 PP-OCRv5。
    """
    env = os.environ.get("LOCAL_OCR_MODELS")
    if env:
        return env
    candidates = []
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(Path(meipass) / "models")
    try:
        candidates.append(Path(sys.executable).resolve().parent / "models")
    except Exception:
        pass
    candidates.append(Path(__file__).resolve().parent / "models")
    for base in candidates:
        for name in ("PP-OCRv6", "PP-OCRv5"):
            p = base / name
            if (p / "det" / "det.xml").is_file():
                return str(p)
    home = Path(os.environ.get("USERPROFILE", str(Path.home())))
    for name in ("PP-OCRv6", "PP-OCRv5"):
        p = home / ".openvino" / "models" / "mineru-ov-models" / name
        if (p / "det" / "det.xml").is_file():
            return str(p)
    return str(home / ".openvino" / "models" / "mineru-ov-models" / "PP-OCRv5")


def default_dict_path(models_dir: str = "") -> str:
    """字库路径：与模型目录配套（PP-OCRv6 -> ppocrv6_dict.txt，v5 -> v5 字库）。"""
    env = os.environ.get("LOCAL_OCR_DICT")
    if env:
        return env
    base = Path(__file__).resolve().parent / "assets"
    if "PP-OCRv6" in str(models_dir):
        p = base / "ppocrv6_dict.txt"
        if p.is_file():
            return str(p)
    return str(base / "ppocrv5_dict.txt")


def main() -> int:
    argv = sys.argv[1:]
    models_dir = default_models_dir()
    dict_path = default_dict_path(models_dir)
    device = "CPU"
    engine_name = "PP-OCRv6" if "PP-OCRv6" in str(models_dir) else "PP-OCRv5"
    i = 0
    while i < len(argv):
        a = argv[i]
        if a == "--models" and i + 1 < len(argv):
            models_dir = argv[i + 1]; i += 2
        elif a == "--dict" and i + 1 < len(argv):
            dict_path = argv[i + 1]; i += 2
        elif a == "--device" and i + 1 < len(argv):
            device = argv[i + 1]; i += 2
        else:
            i += 1

    try:
        sys.stdin.reconfigure(encoding="utf-8")
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

    engine = None

    def respond(obj):
        sys.stdout.write(json.dumps(obj, ensure_ascii=False) + "\n")
        sys.stdout.flush()

    try:
        engine = OcrEngine(models_dir, dict_path, device)
    except Exception as e:
        log(f"初始化失败: {e}")
        respond({"ok": False, "event": "ready", "error": str(e)})
        return 2
    respond({"ok": True, "event": "ready", "engine": engine_name,
             "chars": len(engine.chars)})

    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            req = json.loads(line)
        except Exception:
            respond({"ok": False, "error": "bad json"})
            continue
        rid = req.get("id")
        op = req.get("op")
        try:
            if op == "ping":
                respond({"id": rid, "ok": True})
            elif op == "ocr":
                boxes = engine.ocr(req)
                extra = {}
                # 只有整图 OCR（无 clip、scale=1）才顺便算质量；局部裁剪复识不重复算
                if req.get("clip") is None and float(req.get("scale") or 1.0) == 1.0:
                    try:
                        extra["quality"] = image_quality_metrics(load_image(req["path"]))
                    except Exception:
                        pass
                respond({"id": rid, "ok": True, "boxes": boxes, **extra})
            elif op == "fast_find":
                paths = [str(p) for p in (req.get("paths") or [])]
                results = engine.fast_find_batch(
                    paths,
                    max_edge=int(req.get("max_edge") or 1680),
                    workers=int(req.get("workers") or 8),
                )
                respond({"id": rid, "ok": True, "results": results})
            elif op == "encode":
                enc = engine.encode(req)
                respond({"id": rid, "ok": True, **enc})
            elif op == "table_grid":
                angle = int(req.get("rotate") or 0)
                if angle not in (90, 270):
                    angle = engine.decide_rotation(req["path"])
                grid = detect_table_grid(
                    req["path"],
                    target_long_edge=int(req.get("target_long_edge") or 2200),
                    max_rows=int(req.get("max_rows") or 90),
                    max_cols=int(req.get("max_cols") or 26),
                    rotate=angle,
                )
                try:
                    grid["quality"] = image_quality_metrics(load_image(req["path"]))
                except Exception:
                    pass
                respond({"id": rid, "ok": True, **grid})
            elif op == "render_pdf":
                pages = render_pdf(req["path"], dpi=req.get("dpi", 300))
                respond({"id": rid, "ok": True, "pages": pages})
            elif op == "shutdown":
                respond({"id": rid, "ok": True})
                return 0
            else:
                respond({"id": rid, "ok": False, "error": f"unknown op: {op}"})
        except Exception as e:
            log(f"op={op} 出错: {e}")
            respond({"id": rid, "ok": False, "error": str(e)})
    return 0


if __name__ == "__main__":
    sys.exit(main())
