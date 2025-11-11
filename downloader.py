import os
from pathlib import Path
from typing import Callable, Optional

from yt_dlp import YoutubeDL


def download_video(
    url: str,
    output_dir: Path,
    *,
    cookies: Optional[Path] = None,
    quality: str = "best",  # e.g. "best", "1080p", "720p"
    proxy: Optional[str] = None,
    playlist: bool = False,
    subtitles: bool = False,
    ffmpeg_location: Optional[Path] = None,
    on_progress: Optional[Callable[[dict], None]] = None,
) -> Path:
    """
    Download a single video or a playlist entry using yt-dlp and return the output file path.
    """
    output_dir.mkdir(parents=True, exist_ok=True)

    # map quality to yt-dlp format selector
    max_height = None
    if quality and quality.endswith("p") and quality[:-1].isdigit():
        try:
            max_height = int(quality[:-1])
        except Exception:
            max_height = None
    if max_height:
        fmt = f"bv*[height<=?{max_height}]+ba/b[height<=?{max_height}]"
    else:
        fmt = "bv*+ba/b"

    def _hook(d: dict) -> None:
        if on_progress:
            on_progress(d)

    ydl_opts = {
        "outtmpl": os.path.join(str(output_dir), "%(title).200s [%(id)s].%(ext)s"),
        "format": fmt,
        "merge_output_format": "mp4",
        "concurrent_fragment_downloads": 4,
        "noplaylist": not playlist,
        "ignoreerrors": False,
        "retries": 10,
        "fragment_retries": 10,
        "continuedl": True,
        "nopart": False,
        "progress_hooks": [_hook],
        "ffmpeg_location": str(ffmpeg_location) if ffmpeg_location else None,
        "proxy": proxy,
        "cookiefile": str(cookies) if cookies else None,
        "writesubtitles": subtitles,
        "embedsubtitles": subtitles,
        "subtitleslangs": ["zh-Hans", "zh", "en"] if subtitles else [],
    }

    with YoutubeDL({k: v for k, v in ydl_opts.items() if v is not None}) as ydl:
        info = ydl.extract_info(url, download=True)
        # Determine the resulting file path
        if isinstance(info, dict):
            # playlist case
            if "entries" in info and info["entries"]:
                entry = info["entries"][0]
                if isinstance(entry, dict) and entry.get("requested_downloads"):
                    return Path(entry["requested_downloads"][0]["filepath"])
                return Path(ydl.prepare_filename(entry))
            # single video
            if info.get("requested_downloads"):
                return Path(info["requested_downloads"][0]["filepath"])
            return Path(ydl.prepare_filename(info))
        raise RuntimeError("Download did not return information for the file path")


