import threading
import queue
import io
import sys
from pathlib import Path
from typing import Optional
import tempfile
import subprocess
import shutil
import cv2

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from video_to_pdf_phash import extract_frames_to_pdf, parse_crop
from downloader import download_video


def _enable_windows_dpi_awareness() -> None:
    if sys.platform != 'win32':
        return
    try:
        import ctypes
        user32 = ctypes.windll.user32
        try:
            # Prefer Per-Monitor V2 when available (Win10+)
            user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))  # DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2
            return
        except Exception:
            pass
        try:
            # Fallback to Per-Monitor (Win8.1+)
            ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PROCESS_PER_MONITOR_DPI_AWARE
            return
        except Exception:
            pass
        try:
            # Legacy system-wide awareness
            user32.SetProcessDPIAware()
        except Exception:
            pass
    except Exception:
        pass


class StreamToQueue(io.TextIOBase):
    def __init__(self, line_queue: "queue.Queue[str]") -> None:
        self._queue = line_queue

    def write(self, s: str) -> int:
        if s:
            self._queue.put(s)
        return len(s)

    def flush(self) -> None:
        return None


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("video2pdf GUI")
        # Adjust Tk scaling to current DPI for crisp rendering
        try:
            dpi = float(self.winfo_fpixels('1i'))  # pixels per inch
            self.tk.call('tk', 'scaling', dpi / 72.0)
        except Exception:
            pass
        self.geometry("760x680")
        self._build_ui()

        self._log_queue: "queue.Queue[str]" = queue.Queue()
        self._worker: Optional[threading.Thread] = None
        self.after(100, self._drain_log)

    def _parse_time_to_seconds(self, text: str) -> Optional[float]:
        s = (text or "").strip()
        if not s:
            return None
        try:
            parts = s.split(":")
            if len(parts) == 1:
                return float(parts[0])
            if len(parts) == 2:
                m = int(parts[0])
                sec = float(parts[1])
                return m * 60 + sec
            if len(parts) == 3:
                h = int(parts[0])
                m = int(parts[1])
                sec = float(parts[2])
                return h * 3600 + m * 60 + sec
            raise ValueError("invalid format")
        except Exception:
            raise ValueError("时间格式应为 秒 或 mm:ss 或 hh:mm:ss，可带小数")

    def _ffmpeg_path(self) -> Optional[str]:
        return shutil.which("ffmpeg")

    def _trim_video_with_ffmpeg(self, src: Path, dst: Path, start_s: Optional[float], end_s: Optional[float]) -> None:
        ffmpeg = self._ffmpeg_path()
        if not ffmpeg:
            raise RuntimeError("未找到 ffmpeg，请先安装并加入 PATH")

        args = [ffmpeg, "-y"]
        duration = None
        if start_s is not None and end_s is not None and end_s > start_s:
            duration = end_s - start_s
        if start_s is not None:
            args += ["-ss", f"{start_s:.3f}"]
        args += ["-i", str(src)]
        if duration is not None:
            args += ["-t", f"{duration:.3f}"]

        fast_args = args + ["-c", "copy", str(dst)]
        res = subprocess.run(fast_args, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=False)
        if res.returncode == 0 and dst.exists() and dst.stat().st_size > 0:
            return

        slow_args = args + ["-c:v", "libx264", "-preset", "fast", "-crf", "23", "-c:a", "aac", "-b:a", "192k", str(dst)]
        res2 = subprocess.run(slow_args, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=False)
        if res2.returncode != 0 or not dst.exists() or dst.stat().st_size == 0:
            raise RuntimeError("ffmpeg 剪切失败")

    def _build_ui(self) -> None:
        pad = 8
        main = ttk.Frame(self)
        main.pack(fill=tk.BOTH, expand=True, padx=pad, pady=pad)

        # File inputs
        file_frame = ttk.LabelFrame(main, text="输入/输出")
        file_frame.pack(fill=tk.X, padx=0, pady=(0, pad))

        self.input_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.outdir_var = tk.StringVar()
        self.dldir_var = tk.StringVar()
        self.start_time_var = tk.StringVar()
        self.end_time_var = tk.StringVar()

        self._row(file_frame, "输入视频/网址", self.input_var, browse=lambda: self._browse_file(self.input_var, [
            ("Video files", "*.mp4;*.mov;*.mkv;*.avi;*.flv;*.webm"), ("All files", "*.*")
        ]))
        self._row(file_frame, "输出PDF(可空)", self.output_var, browse=lambda: self._save_file(self.output_var, (
            ("PDF", "*.pdf"), ("All files", "*.*")
        )))
        self._row(file_frame, "帧输出目录(可空)", self.outdir_var, browse=lambda: self._choose_dir(self.outdir_var))
        self._row(file_frame, "下载目录(可空)", self.dldir_var, browse=lambda: self._choose_dir(self.dldir_var))
        self._row(file_frame, "开始时间(可空)", self.start_time_var)
        self._row(file_frame, "结束时间(可空)", self.end_time_var)

        # Time format hint
        hint_row = ttk.Frame(file_frame)
        hint_row.pack(fill=tk.X, pady=(2, 0))
        ttk.Label(hint_row, text="时间格式示例： 90（秒）  或  01:30  或  1:02:03.5").pack(side=tk.LEFT)

        # Parameters
        params = ttk.LabelFrame(main, text="参数")
        params.pack(fill=tk.X, padx=0, pady=(0, pad))

        self.sample_var = tk.StringVar(value="0.5")
        self.threshold_var = tk.StringVar(value="10")
        self.crop_var = tk.StringVar(value="")
        self.scale_width_var = tk.StringVar(value="")
        self.max_pages_var = tk.StringVar(value="")
        self.a4_var = tk.BooleanVar(value=False)

        self._row(params, "采样秒数", self.sample_var)
        self._row(params, "阈值(1-63)", self.threshold_var)
        self._row(params, "裁剪 x,y,w,h(可空)", self.crop_var)
        self._row(params, "统一宽度(像素,可空)", self.scale_width_var)
        self._row(params, "最大页数(可空)", self.max_pages_var)

        a4_row = ttk.Frame(params)
        a4_row.pack(fill=tk.X, pady=(4, 0))
        a4_cb = ttk.Checkbutton(a4_row, text="A4 排版", variable=self.a4_var)
        a4_cb.pack(side=tk.LEFT)

        # Auto-trim
        trim = ttk.LabelFrame(main, text="自动去白边")
        trim.pack(fill=tk.X, padx=0, pady=(0, pad))
        self.auto_trim_var = tk.BooleanVar(value=True)
        self.auto_trim_ratio_var = tk.StringVar(value="0.98")
        self.auto_trim_pad_var = tk.StringVar(value="6")
        self.auto_trim_sides_var = tk.StringVar(value="tb")

        trim_top = ttk.Frame(trim)
        trim_top.pack(fill=tk.X)
        ttk.Checkbutton(trim_top, text="启用自动去白边", variable=self.auto_trim_var).pack(side=tk.LEFT)

        grid = ttk.Frame(trim)
        grid.pack(fill=tk.X, pady=(4, 0))
        self._grid_row(grid, "比例阈值(0-1)", self.auto_trim_ratio_var, 0)
        self._grid_row(grid, "边界余量(px)", self.auto_trim_pad_var, 1)

        side_row = ttk.Frame(trim)
        side_row.pack(fill=tk.X, pady=(4, 0))
        ttk.Label(side_row, text="方向").pack(side=tk.LEFT)
        side = ttk.Combobox(side_row, textvariable=self.auto_trim_sides_var, values=("tb", "all"), state="readonly", width=8)
        side.pack(side=tk.LEFT, padx=(8, 0))

        # Auto-crop
        ac = ttk.LabelFrame(main, text="自动裁剪(PPT区域)")
        ac.pack(fill=tk.X, padx=0, pady=(0, pad))
        self.auto_crop_var = tk.BooleanVar(value=True)
        self.auto_crop_pad_var = tk.StringVar(value="6")
        self.auto_crop_min_area_var = tk.StringVar(value="0.05")

        ac_top = ttk.Frame(ac)
        ac_top.pack(fill=tk.X)
        self.auto_crop_cb = ttk.Checkbutton(ac_top, text="启用自动裁剪", variable=self.auto_crop_var, command=self._on_toggle_auto_crop)
        self.auto_crop_cb.pack(side=tk.LEFT)
        self.manual_select_var = tk.BooleanVar(value=False)
        self.manual_select_cb = ttk.Checkbutton(ac_top, text="主动框选 PPT 区域", variable=self.manual_select_var, command=self._on_toggle_manual_select)
        self.manual_select_cb.pack(side=tk.LEFT, padx=(12, 0))

        ac_grid = ttk.Frame(ac)
        ac_grid.pack(fill=tk.X, pady=(4, 0))
        self._grid_row(ac_grid, "外扩(px)", self.auto_crop_pad_var, 0)
        self._grid_row(ac_grid, "最小面积比例", self.auto_crop_min_area_var, 1)

        # Action and log
        action = ttk.Frame(main)
        action.pack(fill=tk.X, pady=(0, pad))
        self.run_btn = ttk.Button(action, text="开始运行", command=self._on_run)
        self.run_btn.pack(side=tk.LEFT)

        self.log = tk.Text(main, height=18, wrap=tk.WORD)
        self.log.pack(fill=tk.BOTH, expand=True)
        ttk.Scrollbar(self.log, command=self.log.yview)
        self.log.configure(state=tk.NORMAL)

    def _row(self, parent: tk.Widget, label: str, var: tk.StringVar, browse=None) -> None:
        r = ttk.Frame(parent)
        r.pack(fill=tk.X, pady=(4, 0))
        ttk.Label(r, text=label, width=20).pack(side=tk.LEFT)
        e = ttk.Entry(r, textvariable=var)
        e.pack(side=tk.LEFT, fill=tk.X, expand=True)
        if browse is not None:
            ttk.Button(r, text="浏览...", command=browse).pack(side=tk.LEFT, padx=(8, 0))

    def _grid_row(self, parent: tk.Widget, label: str, var: tk.StringVar, row: int) -> None:
        ttk.Label(parent, text=label, width=20).grid(row=row, column=0, sticky=tk.W, pady=2)
        e = ttk.Entry(parent, textvariable=var, width=16)
        e.grid(row=row, column=1, sticky=tk.W, pady=2)

    def _browse_file(self, var: tk.StringVar, types) -> None:
        path = filedialog.askopenfilename(title="选择视频", filetypes=types)
        if path:
            var.set(path)

    def _save_file(self, var: tk.StringVar, types) -> None:
        path = filedialog.asksaveasfilename(title="选择输出PDF", defaultextension=".pdf", filetypes=types)
        if path:
            var.set(path)

    def _choose_dir(self, var: tk.StringVar) -> None:
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            var.set(path)

    def _on_toggle_auto_crop(self) -> None:
        # If enabling auto-crop, disable manual select
        try:
            if bool(self.auto_crop_var.get()):
                self.manual_select_var.set(False)
        except Exception:
            pass

    def _on_toggle_manual_select(self) -> None:
        # If enabling manual select, disable auto-crop
        try:
            if bool(self.manual_select_var.get()):
                self.auto_crop_var.set(False)
        except Exception:
            pass

    def _append_log(self, text: str) -> None:
        self.log.insert(tk.END, text)
        self.log.see(tk.END)

    def _drain_log(self) -> None:
        try:
            while True:
                chunk = self._log_queue.get_nowait()
                self._append_log(chunk)
        except queue.Empty:
            pass
        self.after(100, self._drain_log)

    def _on_run(self) -> None:
        if self._worker and self._worker.is_alive():
            messagebox.showinfo("运行中", "请等待当前任务完成")
            return

        input_path = self.input_var.get().strip()
        if not input_path:
            messagebox.showwarning("缺少输入", "请先输入本地视频路径或视频网址")
            return

        self.log.delete("1.0", tk.END)
        self.run_btn.configure(state=tk.DISABLED)

        def work() -> None:
            stdout = StreamToQueue(self._log_queue)
            stderr = StreamToQueue(self._log_queue)
            try:
                # Determine if input is URL
                is_url = input_path.lower().startswith(("http://", "https://"))

                # If URL, download first
                if is_url:
                    dl_dir_str = self.dldir_var.get().strip()
                    dl_dir = Path(dl_dir_str) if dl_dir_str else Path(tempfile.gettempdir()) / "video2pdf_downloads"
                    print(f"检测到网址输入，开始下载到: {dl_dir}")

                    def on_progress(d: dict) -> None:
                        try:
                            status = d.get("status")
                            if status == "downloading":
                                p = d.get("_percent_str", "").strip()
                                s = d.get("_speed_str", "").strip()
                                eta = d.get("_eta_str", "").strip()
                                line = f"[下载中] {p}  速度 {s}  剩余 {eta}\r"
                                self._log_queue.put(line)
                            elif status == "finished":
                                total = d.get("total_bytes") or d.get("total_bytes_estimate")
                                size_mb = f"{(int(total)/1048576):.1f}MB" if total else "?"
                                self._log_queue.put(f"\n[下载完成] 已保存临时文件，大小约 {size_mb}\n")
                        except Exception:
                            pass

                    downloaded_path = download_video(
                        url=input_path,
                        output_dir=dl_dir,
                        on_progress=on_progress,
                    )
                    input_p = downloaded_path
                    print(f"已下载: {input_p}")
                else:
                    input_p = Path(input_path)

                # Optional trimming via ffmpeg if time range provided
                start_s = None
                end_s = None
                try:
                    start_s = self._parse_time_to_seconds(self.start_time_var.get())
                except Exception as e:
                    raise RuntimeError(f"开始时间格式错误: {e}")
                try:
                    end_s = self._parse_time_to_seconds(self.end_time_var.get())
                except Exception as e:
                    raise RuntimeError(f"结束时间格式错误: {e}")

                if start_s is not None or end_s is not None:
                    if end_s is not None and start_s is not None and end_s <= start_s:
                        raise RuntimeError("结束时间必须大于开始时间")
                    clip_dir = (input_p.parent if input_p.parent.exists() else Path(tempfile.gettempdir())) / "video2pdf_segments"
                    clip_dir.mkdir(parents=True, exist_ok=True)
                    clip_p = clip_dir / (input_p.stem + ".clip.mp4")
                    print(f"检测到剪切参数，开始剪切: start={start_s if start_s is not None else '未指定'}, end={end_s if end_s is not None else '未指定'}")
                    self._trim_video_with_ffmpeg(input_p, clip_p, start_s, end_s)
                    input_p = clip_p
                    print(f"剪切完成: {input_p}")

                # Report video resolution to help manual crop configuration
                try:
                    cap_probe = cv2.VideoCapture(str(input_p))
                    if cap_probe.isOpened():
                        width = int(cap_probe.get(cv2.CAP_PROP_FRAME_WIDTH) or 0)
                        height = int(cap_probe.get(cv2.CAP_PROP_FRAME_HEIGHT) or 0)
                        if not width or not height:
                            ok, frame0 = cap_probe.read()
                            if ok and frame0 is not None:
                                height, width = frame0.shape[:2]
                        cap_probe.release()
                        if width and height:
                            self._log_queue.put(f"\n视频分辨率: {width} x {height}\n")
                except Exception:
                    pass

                output_pdf_str = self.output_var.get().strip()
                output_pdf = Path(output_pdf_str) if output_pdf_str else input_p.with_suffix('.pdf')
                out_dir_str = self.outdir_var.get().strip()
                out_dir = Path(out_dir_str) if out_dir_str else input_p.parent / 'slides_phash'

                # Manual ROI selection (takes precedence over auto-crop)
                manual_crop = None
                if bool(self.manual_select_var.get()):
                    try:
                        cap0 = cv2.VideoCapture(str(input_p))
                        ok, frame = cap0.read()
                        cap0.release()
                        if not ok or frame is None:
                            raise RuntimeError("无法读取视频首帧用于框选")
                        self._log_queue.put("请在弹出的窗口中框选 PPT 区域，按回车确认，Esc 取消。\n")
                        roi = cv2.selectROI("框选 PPT 区域", frame, showCrosshair=True, fromCenter=False)
                        cv2.destroyWindow("框选 PPT 区域")
                        x, y, w, h = map(int, roi)
                        if w > 0 and h > 0:
                            manual_crop = (x, y, w, h)
                            self._log_queue.put(f"已选择区域: x={x}, y={y}, w={w}, h={h}\n")
                        else:
                            self._log_queue.put("已取消框选，继续使用原设置。\n")
                    except Exception as e:
                        self._log_queue.put(f"框选失败: {e}\n")

                # Parse numerics with defaults
                sample_seconds = float(self.sample_var.get() or 0.5)
                threshold = int(self.threshold_var.get() or 10)
                crop = parse_crop(self.crop_var.get().strip() or None)
                if manual_crop is not None:
                    crop = manual_crop
                scale_width = int(self.scale_width_var.get()) if (self.scale_width_var.get().strip()) else None
                max_pages = int(self.max_pages_var.get()) if (self.max_pages_var.get().strip()) else None
                a4 = bool(self.a4_var.get())

                auto_trim = bool(self.auto_trim_var.get())
                auto_trim_ratio = float(self.auto_trim_ratio_var.get() or 0.98)
                auto_trim_pad = int(self.auto_trim_pad_var.get() or 6)
                auto_trim_sides = self.auto_trim_sides_var.get() or 'tb'

                auto_crop = False if manual_crop is not None else bool(self.auto_crop_var.get())
                auto_crop_pad = int(self.auto_crop_pad_var.get() or 6)
                auto_crop_min_area_ratio = float(self.auto_crop_min_area_var.get() or 0.05)

                from contextlib import redirect_stdout, redirect_stderr
                with redirect_stdout(stdout), redirect_stderr(stderr):
                    extract_frames_to_pdf(
                        video_path=input_p,
                        output_pdf=output_pdf,
                        output_dir=out_dir,
                        sample_seconds=sample_seconds,
                        threshold=threshold,
                        crop_region=crop,
                        scale_width=scale_width,
                        a4=a4,
                        max_pages=max_pages,
                        auto_trim=auto_trim,
                        auto_trim_ratio=auto_trim_ratio,
                        auto_trim_pad=auto_trim_pad,
                        auto_trim_sides=auto_trim_sides,
                        auto_crop=auto_crop,
                        auto_crop_pad=auto_crop_pad,
                        auto_crop_min_area_ratio=auto_crop_min_area_ratio,
                    )
                self._log_queue.put("\n完成.\n")
            except Exception as exc:
                self._log_queue.put(f"\n错误: {exc}\n")
                messagebox.showerror("运行失败", str(exc))
            finally:
                self.run_btn.configure(state=tk.NORMAL)

        self._worker = threading.Thread(target=work, daemon=True)
        self._worker.start()


if __name__ == "__main__":
    _enable_windows_dpi_awareness()
    App().mainloop()


