from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


APP_ROOT = Path(__file__).resolve().parent.parent
CACHE_ROOT = APP_ROOT / "cache"

# 第三方库缓存也收敛到项目根目录下的可见 cache/，方便统一清理。
os.environ.setdefault("XDG_CACHE_HOME", str(CACHE_ROOT))
os.environ.setdefault("MPLCONFIGDIR", str(CACHE_ROOT / "matplotlib"))
os.environ.setdefault("PIP_CACHE_DIR", str(CACHE_ROOT / "pip"))
os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", str(CACHE_ROOT / "tools" / "playwright"))

DEFAULT_HOME_URL = ""
DEFAULT_QUERY_URL = ""
DEFAULT_REFRESH_NAME = ""
DEFAULT_REFRESH_CLASS = ""
DEFAULT_THREAD_COUNT = 2
DEFAULT_OCR_RETRIES = 3


@dataclass(frozen=True)
class AppPaths:
    root: Path

    @property
    def data(self) -> Path:
        return self.root

    @property
    def cookies(self) -> Path:
        return self.cache / "cookies"

    @property
    def logs(self) -> Path:
        return self.root / "logs"

    @property
    def output(self) -> Path:
        return self.root / "output"

    @property
    def temp(self) -> Path:
        return self.cache / "temp"

    @property
    def tools(self) -> Path:
        return self.cache / "tools"

    @property
    def cache(self) -> Path:
        return self.root / "cache"

    @property
    def ocr(self) -> Path:
        return self.cache / "captcha"

    def ensure(self) -> None:
        for path in [self.logs, self.output, self.cache, self.cookies, self.temp, self.ocr, self.tools]:
            path.mkdir(parents=True, exist_ok=True)
            keep = path / ".gitkeep"
            if not keep.exists():
                keep.touch()


PATHS = AppPaths(APP_ROOT)
PATHS.ensure()
