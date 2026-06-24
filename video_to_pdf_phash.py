import argparse
import sys
from pathlib import Path
from typing import Optional, Tuple, List

import cv2
import numpy as np
from PIL import Image
from PIL import JpegImagePlugin, PdfImagePlugin

_ = (JpegImagePlugin, PdfImagePlugin)

# Pillow 9/10 compatibility for LANCZOS
try:
    RESAMPLING_LANCZOS = Image.Resampling.LANCZOS  # Pillow >=10
except Exception:
    RESAMPLING_LANCZOS = Image.LANCZOS  # Pillow <10


def parse_crop(crop_str: Optional[str]) -> Optional[Tuple[int, int, int, int]]:
    if not crop_str:
        return None
    try:
        parts = [int(x.strip()) for x in crop_str.split(',')]
        if len(parts) != 4:
            raise ValueError
        x, y, w, h = parts
        if w <= 0 or h <= 0:
            raise ValueError
        return x, y, w, h
    except Exception as exc:
        raise argparse.ArgumentTypeError("--crop 需要格式 x,y,w,h，例如 100,80,1280,720") from exc


def compute_phash_bits(gray_image: np.ndarray) -> np.ndarray:
    resized = cv2.resize(gray_image, (32, 32), interpolation=cv2.INTER_AREA).astype(np.float32)
    dct = cv2.dct(resized)
    low = dct[:8, :8]
    flat = low.flatten()
    median = np.median(flat[1:])
    bits = (low > median).astype(np.uint8).flatten()[1:]  # 63 bits
    return bits


def hamming_distance(a: np.ndarray, b: np.ndarray) -> int:
    return int(np.count_nonzero(a != b))


def safe_imwrite_png(path: Path, bgr_image: np.ndarray, compression: int = 3) -> None:
    success, buf = cv2.imencode('.png', bgr_image, [cv2.IMWRITE_PNG_COMPRESSION, compression])
    if not success:
        raise RuntimeError(f"编码PNG失败: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, 'wb') as f:
        f.write(buf.tobytes())


def render_to_a4(image: Image.Image) -> Image.Image:
    # A4 @ 300 DPI: 2480 x 3508 pixels
    a4_w, a4_h = 2480, 3508
    canvas = Image.new('RGB', (a4_w, a4_h), color='white')
    img = image.convert('RGB')
    scale = min(a4_w / img.width, a4_h / img.height)
    new_w = max(1, int(img.width * scale))
    new_h = max(1, int(img.height * scale))
    resized = img.resize((new_w, new_h), RESAMPLING_LANCZOS)
    offset = ((a4_w - new_w) // 2, (a4_h - new_h) // 2)
    canvas.paste(resized, offset)
    return canvas


def auto_detect_crop_region(frame_bgr: np.ndarray, pad: int = 6, sat_thresh: int = 40, val_thresh: int = 210,
                            morph_kernel: int = 15, min_area_ratio: float = 0.05) -> Tuple[int, int, int, int]:
    """Detects the biggest white-ish rectangle area (typical PPT slide region) from the first frame.
    Returns (x, y, w, h). Falls back to full frame if detection fails.
    """
    height, width = frame_bgr.shape[:2]
    hsv = cv2.cvtColor(frame_bgr, cv2.COLOR_BGR2HSV)
    h_ch, s_ch, v_ch = cv2.split(hsv)
    white_mask = (s_ch < sat_thresh) & (v_ch > val_thresh)
    mask = white_mask.astype(np.uint8) * 255
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (morph_kernel, morph_kernel))
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel, iterations=2)
    mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, kernel, iterations=1)
    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not contours:
        return 0, 0, width, height
    areas = [cv2.contourArea(c) for c in contours]
    idx = int(np.argmax(areas))
    x, y, w, h = cv2.boundingRect(contours[idx])
    if w * h < min_area_ratio * width * height:
        return 0, 0, width, height
    # expand
    x0 = max(0, x - pad)
    y0 = max(0, y - pad)
    x1 = min(width, x + w + pad)
    y1 = min(height, y + h + pad)
    return x0, y0, x1 - x0, y1 - y0


def extract_frames_to_pdf(
    video_path: Path,
    output_pdf: Path,
    output_dir: Path,
    sample_seconds: float,
    threshold: int,
    crop_region: Optional[Tuple[int, int, int, int]],
    scale_width: Optional[int],
    a4: bool,
    max_pages: Optional[int],
    auto_trim: bool,
    auto_trim_ratio: float,
    auto_trim_pad: int,
    auto_trim_sides: str,
    auto_crop: bool,
    auto_crop_pad: int,
    auto_crop_min_area_ratio: float,
) -> None:
    cap = cv2.VideoCapture(str(video_path))
    if not cap.isOpened():
        raise RuntimeError(f"无法打开视频: {video_path}")

    fps = cap.get(cv2.CAP_PROP_FPS)
    if not fps or fps <= 1e-3:
        fps = 25.0
    step = max(1, int(round(fps * max(1e-3, sample_seconds))))

    print(f"视频: {video_path}")
    print(f"FPS: {fps:.3f}, 每 {sample_seconds}s 取一帧 -> 步长 {step}")
    print(f"阈值: {threshold}, 裁剪: {crop_region if crop_region else ('auto' if auto_crop else '无')}, 缩放宽度: {scale_width if scale_width else '不缩放'}")
    print(f"临时输出目录: {output_dir}")

    output_dir.mkdir(parents=True, exist_ok=True)

    frame_index = 0
    kept_index = 0
    last_bits: Optional[np.ndarray] = None
    saved_paths: List[Path] = []
    global_trim_bounds: Optional[Tuple[int, int, int, int]] = None  # (y0, y1, x0, x1) offsets within save_img

    while True:
        ret, frame = cap.read()
        if not ret:
            break

        if frame_index % step != 0:
            frame_index += 1
            continue

        work = frame

        # Auto-crop from the first sampled frame if requested
        if auto_crop and crop_region is None and kept_index == 0 and frame_index % step == 0:
            ax, ay, aw, ah = auto_detect_crop_region(
                frame,
                pad=int(auto_crop_pad),
                min_area_ratio=float(auto_crop_min_area_ratio),
            )
            crop_region = (ax, ay, aw, ah)
            print(f"Auto-crop detected: x={ax}, y={ay}, w={aw}, h={ah}")

        if crop_region:
            x, y, w, h = crop_region
            h_img, w_img = work.shape[:2]
            x0 = max(0, min(x, w_img - 1))
            y0 = max(0, min(y, h_img - 1))
            x1 = max(1, min(x + w, w_img))
            y1 = max(1, min(y + h, h_img))
            if x1 <= x0 or y1 <= y0:
                frame_index += 1
                continue
            work = work[y0:y1, x0:x1]

        gray = cv2.cvtColor(work, cv2.COLOR_BGR2GRAY)
        bits = compute_phash_bits(gray)

        is_new_slide = last_bits is None or hamming_distance(bits, last_bits) >= threshold
        if is_new_slide:
            kept_index += 1
            # 可选缩放统一宽度，保持纵横比
            save_img = work
            if scale_width and scale_width > 0 and work.shape[1] != scale_width:
                scale = scale_width / work.shape[1]
                new_w = scale_width
                new_h = max(1, int(round(work.shape[0] * scale)))
                save_img = cv2.resize(work, (new_w, new_h), interpolation=cv2.INTER_AREA)

            # 自动去除白色上下(或四周)边框，以统一裁掉 PPT 留白
            if auto_trim:
                y0, y1, x0, x1 = (0, save_img.shape[0], 0, save_img.shape[1])
                if global_trim_bounds is None:
                    hsv = cv2.cvtColor(save_img, cv2.COLOR_BGR2HSV)
                    s = hsv[:, :, 1]
                    v = hsv[:, :, 2]
                    white = (s < 40) & (v > 210)
                    # 计算每行/列白色占比
                    row_ratio = white.mean(axis=1)
                    col_ratio = white.mean(axis=0)

                    def find_bounds(ratio_vec: np.ndarray, ratio_thresh: float) -> Tuple[int, int]:
                        top = 0
                        while top < ratio_vec.shape[0] and ratio_vec[top] >= ratio_thresh:
                            top += 1
                        bottom = ratio_vec.shape[0] - 1
                        while bottom >= 0 and ratio_vec[bottom] >= ratio_thresh:
                            bottom -= 1
                        if bottom < top:
                            return 0, ratio_vec.shape[0]
                        return max(0, top - auto_trim_pad), min(ratio_vec.shape[0], bottom + 1 + auto_trim_pad)

                    if auto_trim_sides in ("tb", "all"):
                        y0, y1 = find_bounds(row_ratio, auto_trim_ratio)
                    if auto_trim_sides == "all":
                        x0, x1 = find_bounds(col_ratio, auto_trim_ratio)
                    global_trim_bounds = (y0, y1, x0, x1)
                    print(f"Auto-trim bounds set from first kept frame: y[{y0}:{y1}] x[{x0}:{x1}] (ratio>={auto_trim_ratio}, pad={auto_trim_pad}, sides={auto_trim_sides})")

                y0, y1, x0, x1 = global_trim_bounds
                # 安全保护
                y0 = max(0, min(y0, save_img.shape[0]))
                y1 = max(y0 + 1, min(y1, save_img.shape[0]))
                x0 = max(0, min(x0, save_img.shape[1]))
                x1 = max(x0 + 1, min(x1, save_img.shape[1]))
                save_img = save_img[y0:y1, x0:x1]

            out_path = output_dir / f"slide-{kept_index:03d}.png"
            safe_imwrite_png(out_path, save_img)
            saved_paths.append(out_path)
            last_bits = bits

            if max_pages and len(saved_paths) >= max_pages:
                print(f"达到最大页数限制: {max_pages}")
                break

        if kept_index and kept_index % 25 == 0:
            print(f"已保留 {kept_index} 帧…")

        frame_index += 1

    cap.release()

    if not saved_paths:
        raise RuntimeError("未提取到任何帧。尝试降低阈值(--threshold)或缩短取样间隔(--sample-seconds)")

    print(f"共选取 {len(saved_paths)} 帧，开始生成 PDF: {output_pdf}")

    images: List[Image.Image] = []
    for p in saved_paths:
        img = Image.open(p).convert('RGB')
        if a4:
            img = render_to_a4(img)
        images.append(img)

    first, rest = images[0], images[1:]
    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    first.save(str(output_pdf), format="PDF", save_all=True, append_images=rest)
    print(f"完成: {output_pdf} (共 {len(images)} 页)")


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="基于pHash的PPT关键帧提取并合成为PDF")
    parser.add_argument("--input", "-i", required=True, help="输入视频路径")
    parser.add_argument("--output", "-o", default=None, help="输出PDF路径(默认与视频同名)" )
    parser.add_argument("--out-dir", default=None, help="保留帧输出目录(默认: <视频目录>/slides_phash)")
    parser.add_argument("--sample-seconds", type=float, default=0.5, help="取样间隔秒数(默认0.5)")
    parser.add_argument("--threshold", type=int, default=10, help="pHash 汉明距离阈值，越大越少(默认10)")
    parser.add_argument("--crop", type=parse_crop, default=None, help="裁剪区域 x,y,w,h，例如 100,80,1280,720")
    parser.add_argument("--scale-width", type=int, default=None, help="将保留帧按宽度统一缩放(像素)。保持比例")
    parser.add_argument("--a4", action="store_true", help="将图片居中放置到A4(300DPI)页面上")
    parser.add_argument("--max-pages", type=int, default=None, help="限制最大页数(可选)")
    parser.add_argument("--auto-trim", action="store_true", help="自动去除白色留白(默认仅上下)")
    parser.add_argument("--auto-trim-ratio", type=float, default=0.98, help="将行/列中白色像素占比>=该值视为留白(0~1)")
    parser.add_argument("--auto-trim-pad", type=int, default=6, help="在裁切边界上各留出像素余量")
    parser.add_argument("--auto-trim-sides", choices=["tb", "all"], default="tb", help="仅上下(tb)或四周(all)去白")
    parser.add_argument("--auto-crop", action="store_true", help="从首帧自动检测PPT区域并裁剪")
    parser.add_argument("--auto-crop-pad", type=int, default=6, help="自动裁剪时在检测框外扩像素余量")
    parser.add_argument("--auto-crop-min-area-ratio", type=float, default=0.05, help="最小检测区域面积占整帧比例，过小则放弃自动裁剪")

    args = parser.parse_args(argv)

    # Robust handling of Unicode/odd paths: avoid resolve strict failures
    try:
        video_path = Path(args.input).expanduser().resolve()
    except OSError:
        video_path = Path(args.input)
    if not video_path.exists():
        # still try raw string
        video_path = Path(args.input)
        if not video_path.exists():
            print(f"找不到视频: {video_path}", file=sys.stderr)
            return 2

    output_pdf = Path(args.output).expanduser().resolve() if args.output else video_path.with_suffix('.pdf')
    out_dir = Path(args.out_dir).expanduser().resolve() if args.out_dir else video_path.parent / 'slides_phash'

    if args.threshold < 1 or args.threshold > 63:
        print("--threshold 建议在 1..63 之间", file=sys.stderr)
        return 2

    try:
        extract_frames_to_pdf(
            video_path=video_path,
            output_pdf=output_pdf,
            output_dir=out_dir,
            sample_seconds=float(args.sample_seconds),
            threshold=int(args.threshold),
            crop_region=args.crop,
            scale_width=int(args.scale_width) if args.scale_width else None,
            a4=bool(args.a4),
            max_pages=int(args.max_pages) if args.max_pages else None,
            auto_trim=bool(args.auto_trim),
            auto_trim_ratio=float(args.auto_trim_ratio),
            auto_trim_pad=int(args.auto_trim_pad),
            auto_trim_sides=str(args.auto_trim_sides),
            auto_crop=bool(args.auto_crop),
            auto_crop_pad=int(args.auto_crop_pad),
            auto_crop_min_area_ratio=float(args.auto_crop_min_area_ratio),
        )
        return 0
    except Exception as exc:
        print(f"错误: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())


