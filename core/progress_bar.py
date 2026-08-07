"""
进度条工具类 - 为数据迁移/导出工具提供统一的终端进度显示
支持：
  - 普通计数进度条（总数已知的操作）
  - 已过时间 / 预计剩余时间
  - 平滑更新频率控制（避免刷屏）
  - Windows/Linux 兼容 ASCII 字符
"""

import sys
import time
from datetime import datetime
from typing import Optional


class ProgressBar:
    """统一的单行进度条
    - 使用 █ ░ 块字符（和 ExportProgress 相同风格）
    - 数字自动 K/M 简写（2700000 → 2.7M）
    - 支持 \r 回车覆盖刷新（单行刷新，避免换行刷屏）
    - 支持 write_log() 在不破坏进度条的前提下穿插打印日志
    - 线程安全（多线程共享同一进度条对象时安全）
    """

    def __init__(
        self,
        total: int,
        prefix: str = "进度",
        width: int = 30,
        unit: str = "条",
        show_time: bool = True,
        min_update_interval: float = 0.25,
    ):
        self.total = max(1, total)
        self.prefix = prefix
        self.width = width
        self.unit = unit
        self.show_time = show_time
        self.min_update_interval = min_update_interval
        self.start_time = time.time()
        self.current = 0
        self._last_update = 0.0
        self._last_line_len = 0
        self._finished = False
        # 多线程锁：防止 bar.add 与 write_log 并发输出混乱
        try:
            import threading
            self._lock = threading.Lock()
        except Exception:
            self._lock = None

    @staticmethod
    def _human(n: int) -> str:
        """数字简写：1200 -> 1.2K, 1500000 -> 1.5M"""
        if n >= 1000000:
            return f"{n / 1000000:.1f}M"
        if n >= 1000:
            return f"{n / 1000:.1f}K"
        return f"{n}"

    def write_log(self, msg: str) -> None:
        """在进度条进行中穿插打印一条日志，同时保持进度条的下一行刷新。
        这样 MigrationLogger 输出的 [INFO] 行不会破坏进度条显示。
        """
        if self._finished:
            print(msg)
            return
        if self._lock:
            self._lock.acquire()
        try:
            # 清除当前进度条行，换行，打印日志，再重新渲染进度条
            sys.stdout.write("\r" + " " * self._last_line_len + "\r")
            sys.stdout.write(msg + "\n")
            sys.stdout.flush()
            self._render(force=True)
            self._last_line_len = 0
        finally:
            if self._lock:
                self._lock.release()

    def _format_time(self, seconds: float) -> str:
        """格式化时间（无小数秒，如 5s、1m30s"""
        if seconds < 60:
            return f"{seconds:.0f}s"
        minutes = int(seconds // 60)
        sec = int(seconds % 60)
        if minutes < 60:
            return f"{minutes}m{sec:02d}s"
        hours = minutes // 60
        minutes = minutes % 60
        return f"{hours}h{minutes:02d}m"

    def update(self, current: Optional[int] = None, extra: str = "") -> None:
        """更新进度。current 为 None 时自动 +1"""
        if self._finished:
            return
        if current is not None:
            self.current = current
        else:
            self.current += 1

        # 节流：避免频繁更新
        now = time.time()
        if now - self._last_update < self.min_update_interval and self.current < self.total:
            return
        self._last_update = now

        if self._lock:
            self._lock.acquire()
        try:
            self._render()
        finally:
            if self._lock:
                self._lock.release()

    def add(self, delta: int, extra: str = "") -> None:
        """增加指定数量"""
        if delta <= 0:
            return
        self.update(self.current + delta, extra)

    def _render(self, force: bool = False) -> None:
        """渲染进度条（在终端中回车重写）
        force=True 时忽略节流，强制刷新（用于 write_log 后恢复进度条行）
        """
        progress = min(self.current, self.total) / self.total
        filled = int(self.width * progress)
        bar = "█" * filled + "░" * (self.width - filled)
        percent = int(progress * 100)

        elapsed = time.time() - self.start_time

        # ========== ETA 计算（生产级：单调递减，禁止时光倒流）==========
        eta = 0.0
        if self.current > 0 and progress > 0 and self.current < self.total:
            # 基于当前平均速度估算剩余时间
            avg_speed = self.current / elapsed
            remaining_rows = self.total - self.current
            calc_eta = remaining_rows / max(avg_speed, 0.001)

            # 单调递减保护：ETA 只降不升
            if not hasattr(self, '_min_eta_observed'):
                self._min_eta_observed = calc_eta
            elif calc_eta < self._min_eta_observed:
                self._min_eta_observed = calc_eta

            eta = max(self._min_eta_observed, 0.0)

        parts = [
            f"{self.prefix}",
            f"[{bar}]",
            f"{percent:3d}%",
            f"{self._human(self.current)}/{self._human(self.total)}",
        ]
        if self.show_time:
            if self.current < self.total and eta > 1.0:
                parts.append(f"剩 {self._format_time(eta)}")
            elif self.current < self.total and progress >= 0.98:
                parts.append("剩 即将完成")
            parts.append(f"用时 {self._format_time(elapsed)}")

        line = " ".join(parts)

        # 清尾并输出 —— 即使 isatty=False（IDE / 管道）也用 \r 刷新
        if len(line) < self._last_line_len:
            line = line + " " * (self._last_line_len - len(line))
        sys.stdout.write("\r" + line)
        sys.stdout.flush()
        self._last_line_len = len(line)

    def finish(self, extra: str = "") -> None:
        """完成进度条"""
        if self._finished:
            return
        self._finished = True
        self.current = self.total
        if self._lock:
            self._lock.acquire()
        try:
            self._render()
        finally:
            if self._lock:
                self._lock.release()
        sys.stdout.write("\n")
        sys.stdout.flush()

    def get_elapsed(self) -> float:
        """获取已耗时（秒）"""
        return time.time() - self.start_time

    def __enter__(self):
        return self

    def __exit__(self, *args):
        self.finish()


class Spinner:
    """简单旋转指示器（用于总数未知的操作）"""

    _CHARS = ["-", "\\", "|", "/"]

    def __init__(self, message: str = "处理中"):
        self.message = message
        self.start_time = time.time()
        self._idx = 0
        self._finished = False
        self._last_update = 0.0
        self._last_line_len = 0
        self._is_tty = hasattr(sys.stdout, "isatty") and sys.stdout.isatty()

    def tick(self, extra: str = "") -> None:
        """刷新一次"""
        if self._finished:
            return
        now = time.time()
        if now - self._last_update < 0.15:
            return
        self._last_update = now
        ch = self._CHARS[self._idx % len(self._CHARS)]
        self._idx += 1
        elapsed = now - self.start_time
        line = f"{self.message}... {ch} 已用时 {elapsed:.1f}s"
        if extra:
            line += " " + extra
        if self._is_tty:
            if len(line) < self._last_line_len:
                line = line + " " * (self._last_line_len - len(line))
            sys.stdout.write("\r" + line)
            sys.stdout.flush()
        else:
            sys.stdout.write(line + "\n")
        self._last_line_len = len(line)

    def finish(self, message: str = "完成") -> None:
        """结束指示器，显示最终状态"""
        if self._finished:
            return
        self._finished = True
        elapsed = time.time() - self.start_time
        line = f"{self.message}... {message} 总用时 {elapsed:.2f}s"
        if self._is_tty:
            sys.stdout.write("\r" + line + " " * max(0, self._last_line_len - len(line)) + "\n")
            sys.stdout.flush()
        else:
            sys.stdout.write(line + "\n")


def print_section(title: str, width: int = 60) -> None:
    """打印带边框的标题"""
    border = "=" * width
    print()
    print(border)
    print(f" {title}")
    print(border)
    print()
    print("-" * width)
    print(f"【{title}】")
    print("-" * width)


def print_summary(title: str, lines: list, total_time: float, width: int = 60) -> None:
    """打印结果汇总"""
    border = "-" * width
    print()
    print(border)
    print(f"【{title}】")
    print(border)
    for line in lines:
        print(f"  {line}")
    print(border)
    print(f"总用时: {total_time:.2f} 秒")
    print(border)


def timestamp_str() -> str:
    """返回当前时间字符串"""
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")