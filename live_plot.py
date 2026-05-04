"""Live plot example for a Coherent FieldMax II power meter.

Edit the values in `SETTINGS` to change startup behavior, especially
`history_seconds`, which controls how much recent data stays visible.
"""

import math
import threading
import time
from collections import deque
from dataclasses import dataclass
from typing import Any

import matplotlib.pyplot as plt
from fieldmax_power_meter import error_print, power_meter_handler
from matplotlib.animation import FuncAnimation
from matplotlib.artist import Artist


@dataclass(slots=True)
class LivePlotSettings:
    """Startup settings for the live plot example."""

    dll_path: str | None = None
    device_idx: int = 0
    wavelength_nm: int | None = 1980
    auto_range: bool | None = True
    zero_on_start: bool = False
    read_interval_s: float = 0.20
    history_seconds: float = 1000.0
    average_seconds: float = 30.0
    read_timeout_s: float = 2.0
    window_width: int = 1000
    window_height: int = 600
    redraw_interval_ms: int = 100
    window_title: str = "FieldMaxPM"
    always_on_top: bool = True


SETTINGS = LivePlotSettings()


def format_power(power_w: float | None) -> str:
    """Format a power value using a readable unit."""
    if power_w is None:
        return "--"

    abs_power = abs(power_w)
    if abs_power >= 1.0:
        return f"{power_w:.3f} W"
    if abs_power >= 1e-3:
        return f"{power_w * 1e3:.3f} mW"
    if abs_power >= 1e-6:
        return f"{power_w * 1e6:.3f} uW"
    if abs_power >= 1e-9:
        return f"{power_w * 1e9:.3f} nW"
    return f"{power_w:.3e} W"


def format_power_mw(power_w: float | None) -> str:
    """Format the power value in milliwatts with 0 decimals."""
    if power_w is None:
        return "--"
    return f"{power_w * 1e3:.0f} mW"


def configure_meter(settings: LivePlotSettings) -> power_meter_handler:
    """Connect to the meter and apply the requested startup settings."""
    pm = power_meter_handler(dll_path=settings.dll_path)

    if not pm.connect(settings.device_idx, sync=False, timeout_s=5):
        raise RuntimeError(f"Failed to connect to FieldMax device index {settings.device_idx}.")

    if settings.wavelength_nm is not None:
        success = pm.set_wavelength_nm(settings.wavelength_nm, sync=False)
        if not success:
            raise RuntimeError(f"Failed to set wavelength to {settings.wavelength_nm} nm.")

    if settings.auto_range is not None:
        success = pm.set_auto_range(settings.auto_range, sync=True)
        if not success:
            raise RuntimeError("Failed to set auto range.")

    if settings.zero_on_start:
        pm.set_current_power_to_0()

    return pm


def validate_settings(settings: LivePlotSettings) -> None:
    """Reject invalid startup values before talking to the meter."""
    if settings.history_seconds <= 0:
        raise ValueError("history_seconds must be greater than 0.")
    if settings.average_seconds < 0:
        raise ValueError("average_seconds must be 0 or greater.")
    if settings.read_interval_s <= 0:
        raise ValueError("read_interval_s must be greater than 0.")
    if settings.read_timeout_s <= 0:
        raise ValueError("read_timeout_s must be greater than 0.")
    if settings.redraw_interval_ms <= 0:
        raise ValueError("redraw_interval_ms must be greater than 0.")


class LivePlotApp:
    """Small matplotlib app that draws a scrolling power trace."""

    def __init__(self, meter: power_meter_handler, settings: LivePlotSettings):
        self.meter = meter
        self.settings = settings
        self.samples: deque[tuple[float, float]] = deque()
        self.samples_lock = threading.Lock()
        self.stop_event = threading.Event()
        self.reader_thread = threading.Thread(target=self._reader_loop, daemon=True)
        self.latest_power_w: float | None = None
        self.status_text = "Connecting..."
        self.topmost_timer: Any | None = None

        self.fig, self.ax = plt.subplots(figsize=(settings.window_width / 100, settings.window_height / 100))
        self.fig.canvas.mpl_connect("close_event", lambda event: self.close())  # noqa

        try:
            self.fig.canvas.manager.set_window_title(settings.window_title)  # type: ignore
        except Exception:
            pass

        (self.line_raw,) = self.ax.plot([], [], color="#0078d4", label="Raw", zorder=1)
        (self.line_avg,) = self.ax.plot(
            [],
            [],
            color="red",
            linestyle="-",
            label=f"Avg (~{settings.average_seconds:.1f}s)",
            zorder=2,
        )
        (self.avg_window_line,) = self.ax.plot(
            [],
            [],
            color="#808080",
            linestyle=":",
            linewidth=1,
            label="_nolegend_",
        )

        self.status_text_obj = self.ax.text(
            0.01,
            0.95,
            "",
            transform=self.ax.transAxes,
            ha="left",
            va="top",
            fontsize=10,
            bbox={"facecolor": "white", "alpha": 0.8, "edgecolor": "none"},
        )

        self.ax.set_xlabel("Seconds ago")
        self.ax.set_ylabel("Power [mW]")
        self.ax.set_title(settings.window_title)
        self.ax.legend(loc="lower left")
        self.ax.grid(True, linewidth=0.5, color="#cccccc", alpha=0.6)
        self.ax.set_xlim(settings.history_seconds, 0)
        self.fig.tight_layout()

        if settings.always_on_top:
            self.fig.canvas.mpl_connect("draw_event", lambda event: self._set_window_on_top())  # noqa
            self._set_window_on_top()

            try:
                timer = self.fig.canvas.new_timer(interval=1000)
                timer.add_callback(self._set_window_on_top)
                timer.start()
                self.topmost_timer = timer
            except Exception:
                pass

        self.animation = FuncAnimation(
            self.fig,
            self._draw_plot,
            interval=settings.redraw_interval_ms,
            blit=False,
            cache_frame_data=False,
        )

    def run(self) -> None:
        """Start background reads and launch the UI loop."""
        self.status_text = "Connected"
        self.reader_thread.start()
        plt.show()

    def close(self) -> None:
        """Stop the reader thread and release the meter connection."""
        self.stop_event.set()

        if self.topmost_timer is not None:
            try:
                self.topmost_timer.stop()
            except Exception:
                pass
            self.topmost_timer = None

        try:
            self.reader_thread.join(timeout=1.0)
        except RuntimeError:
            pass

        try:
            self.meter.final_shutdown()
        except Exception:
            pass

    def _set_window_on_top(self) -> None:
        """Try to keep the figure window above other windows."""
        try:
            manager = self.fig.canvas.manager
            window = getattr(manager, "window", None)

            if window is None and hasattr(self.fig.canvas, "get_tk_widget"):
                try:
                    tk_widget = self.fig.canvas.get_tk_widget()
                    window = tk_widget.winfo_toplevel()
                except Exception:
                    window = None

            if window is None:
                return

            if hasattr(window, "attributes"):
                window.attributes("-topmost", True)
            elif hasattr(window, "wm_attributes"):
                window.wm_attributes("-topmost", True)
            elif hasattr(window, "setWindowFlags"):
                from PyQt5 import QtCore  # noqa
                window.setWindowFlag(QtCore.Qt.WindowType.WindowStaysOnTopHint, True)
                window.show()
                window.raise_()
                window.activateWindow()
                if hasattr(window, "raise_"):
                    window.raise_()
                if hasattr(window, "activateWindow"):
                    window.activateWindow()

        except Exception:
            pass

    def _reader_loop(self) -> None:
        """Poll the meter continuously in the background."""
        next_read_time = time.monotonic()

        while not self.stop_event.is_set():
            now = time.monotonic()
            wait_time = next_read_time - now

            if wait_time > 0 and self.stop_event.wait(wait_time):
                break

            power_mean = self.meter.read_power_W(
                print_error=False,
                timeout_s=self.settings.read_timeout_s,
            )[1]

            sample_time = time.monotonic()

            if power_mean is not None:
                with self.samples_lock:
                    self.samples.append((sample_time, power_mean))
                    self._trim_samples_locked(sample_time)

                self.latest_power_w = power_mean
                self.status_text = ""
            else:
                self.status_text = "Waiting for valid reading..."

            next_read_time = sample_time + self.settings.read_interval_s

    def _trim_samples_locked(self, current_time: float) -> None:
        """Drop samples that are outside the visible history window."""
        cutoff = current_time - self.settings.history_seconds

        while self.samples and self.samples[0][0] < cutoff:
            self.samples.popleft()

    def _copy_samples(self) -> list[tuple[float, float]]:
        """Return a snapshot of the current sample buffer."""
        current_time = time.monotonic()

        with self.samples_lock:
            self._trim_samples_locked(current_time)
            return list(self.samples)

    def _effective_average_count(self) -> int:
        """Return the number of samples to average based on the configured seconds."""
        if self.settings.average_seconds <= 0:
            return 1

        count = round(self.settings.average_seconds / self.settings.read_interval_s)
        return max(1, count)

    def _compute_running_average(
        self,
        samples: list[tuple[float, float]],
    ) -> list[tuple[float, float]]:
        """Compute a simple trailing average over the most recent samples."""
        if not samples:
            return []

        average_count = self._effective_average_count()

        if average_count <= 1:
            return []

        window: deque[float] = deque()
        running_sum = 0.0
        avg_samples: list[tuple[float, float]] = []

        for timestamp, power in samples:
            window.append(power)
            running_sum += power

            if len(window) > average_count:
                running_sum -= window.popleft()

            avg_samples.append((timestamp, running_sum / len(window)))

        return avg_samples

    def _select_display_units(
        self,
        samples: list[tuple[float, float]],
        avg_samples: list[tuple[float, float]],
    ) -> tuple[float, str, int]:
        """Return scale, unit label, and decimals based on the current history."""
        max_power = 0.0

        if samples:
            max_power = max(abs(power) for _, power in samples)

        if avg_samples:
            max_power = max(max_power, max(abs(power) for _, power in avg_samples))

        if max_power == 0.0 and self.latest_power_w is not None:
            max_power = abs(self.latest_power_w)

        if max_power > 1.0:
            return 1.0, "W", 3

        return 1e3, "mW", 1

    def _format_significant(self, value: float, sig: int) -> str:
        """Format a number with a fixed number of significant digits."""
        if value == 0:
            return f"0.{'0' * (sig - 1)}"

        sign = "-" if value < 0 else ""
        abs_value = abs(value)
        magnitude = math.floor(math.log10(abs_value))
        decimals = max(0, sig - magnitude - 1)

        return f"{sign}{abs_value:.{decimals}f}"

    def _format_display_power(
        self,
        power_w: float | None,
        scale: float,
        unit: str,
        decimals: int,
    ) -> str:
        """Format a power value using the selected display unit."""
        del decimals

        if power_w is None:
            return "--"

        value = power_w * scale
        return f"{self._format_significant(value, 3)} {unit}"

    def _draw_plot(self, frame: int) -> tuple[Artist, ...]:
        """Render the live plot frame."""
        del frame

        samples = self._copy_samples()
        avg_samples = self._compute_running_average(samples)
        scale, unit, decimals = self._select_display_units(samples, avg_samples)

        self.ax.set_ylabel(f"Power [{unit}]")

        artists: tuple[Artist, ...] = (
            self.line_raw,
            self.line_avg,
            self.avg_window_line,
            self.status_text_obj,
        )

        if not samples:
            self.line_raw.set_data([], [])
            self.line_avg.set_data([], [])
            self.avg_window_line.set_data([], [])
            self.ax.set_xlim(self.settings.history_seconds, 0)
            self.status_text_obj.set_text(
                f"Latest: {self._format_display_power(self.latest_power_w, scale, unit, decimals)}"
            )
            return artists

        now = time.monotonic()

        x = [now - sample_time for sample_time, _ in samples]
        y = [power * scale for _, power in samples]

        self.line_raw.set_data(x, y)

        if avg_samples:
            x_avg = [now - sample_time for sample_time, _ in avg_samples]
            y_avg = [avg_power * scale for _, avg_power in avg_samples]
            self.line_avg.set_data(x_avg, y_avg)
        else:
            y_avg = []
            self.line_avg.set_data([], [])

        if y_avg:
            min_power = min(min(y), min(y_avg))
            max_power = max(max(y), max(y_avg))
        else:
            min_power = min(y)
            max_power = max(y)

        if min_power == max_power:
            padding = max(abs(min_power) * 0.05, 1e-9)
        else:
            padding = (max_power - min_power) * 0.1

        min_power -= padding
        max_power += padding

        self.ax.set_xlim(self.settings.history_seconds, 0)
        self.ax.set_ylim(min_power, max_power)

        avg_window_x = max(0.0, self.settings.average_seconds)
        self.avg_window_line.set_data(
            [avg_window_x, avg_window_x],
            [min_power, max_power],
        )

        latest_text = self._format_display_power(self.latest_power_w, scale, unit, decimals)
        avg_text = self._format_display_power(
            avg_samples[-1][1] if avg_samples else None,
            scale,
            unit,
            decimals,
        )

        self.status_text_obj.set_text(
            f"Latest: {latest_text} | "
            f"Avg (~{self.settings.average_seconds:.1f}s): {avg_text} | "
            f"Status: {self.status_text}"
        )

        return artists


def main() -> None:
    """Run the live plot example."""
    meter = None

    try:
        validate_settings(SETTINGS)
        meter = configure_meter(SETTINGS)

        app = LivePlotApp(meter, SETTINGS)
        app.run()

    except Exception as exc:
        error_print(f"Live plot failed: {exc}")

        if meter is not None:
            try:
                meter.final_shutdown()
            except Exception:
                pass


if __name__ == "__main__":
    main()
