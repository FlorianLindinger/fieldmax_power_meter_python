"""Process-isolated Python wrapper for Coherent FieldMax II power meters.

This code started as an adaptation of `pyFieldMaxII`:
https://github.com/jscman/pyFieldMaxII

Example:
    pm = power_meter_handler(dll_path=None)
    pm.connect(device_idx=0)
    pm.set_current_power_to_0()  # optional
    power_min_W, power_mean_W, power_max_W = pm.read_power_W()
    pm.disconnect()
"""

import atexit
import ctypes
import math
import multiprocessing
import os
import time
import traceback

MODULE_DIR = os.path.dirname(os.path.abspath(__file__))
LOCAL_DLL_PATH = os.path.join(MODULE_DIR, "FieldMax2Lib.dll")
GLOBAL_DLL_PATH = (
    r"C:\Program Files (x86)\Coherent\FieldMaxII PC\Drivers\Win10"
    r"\FieldMax2Lib\x64\FieldMax2Lib.dll"
)
FIELDMAX_INSTALLER_URL = "https://repo.coherent.com/software/FieldMaxII_v3.3.2.9_rc1_setup.exe"


def error_print(message, max_wrapper_len=20, wrapper_symbol="=", middle_symbol="-"):
    """Print an error banner and the active traceback, if one exists."""
    msg_len = len(message)
    if msg_len > max_wrapper_len:
        msg_len = max_wrapper_len
    print(wrapper_symbol * msg_len)
    print(message)
    error = traceback.format_exc()
    if error.strip() != "NoneType: None":
        print(middle_symbol * msg_len)
        print(error, end="")
    print(wrapper_symbol * msg_len)


def _resolve_dll_path(dll_path: str | None) -> str:
    """Resolve an explicit DLL path or the standard local-then-global lookup."""
    if dll_path is None:
        candidates = [LOCAL_DLL_PATH, GLOBAL_DLL_PATH]
    else:
        candidates = [os.path.abspath(dll_path)]

    for candidate in candidates:
        if os.path.exists(candidate):
            return os.path.abspath(candidate)

    if dll_path is None:
        raise FileNotFoundError(
            "[Error]: DLL file not found. Checked local path "
            f'"{LOCAL_DLL_PATH}" and expected global path "{GLOBAL_DLL_PATH}". '
            f'Install the Coherent software from "{FIELDMAX_INSTALLER_URL}" '
            "or place FieldMax2Lib.dll next to this module."
        )

    resolved_path = os.path.abspath(dll_path)
    raise FileNotFoundError(
        f'[Error]: DLL file not found at "{resolved_path}". '
        f'Install the Coherent software from "{FIELDMAX_INSTALLER_URL}" or use '
        "power_meter_handler(dll_path=None) to search the local repo copy first "
        "and then the global install path."
    )


def _driver_worker(conn, dll_path: str):
    """Execute all DLL calls in a child process to isolate DLL hangs."""
    conn.send(("init_ok", None))
    dll = None

    def _load_dll():
        """Load the vendor DLL lazily inside the worker process."""
        nonlocal dll
        if dll is None:
            dll = ctypes.windll.LoadLibrary(dll_path)
        return dll

    while True:
        try:
            cmd, payload = conn.recv()
        except EOFError:
            break
        except Exception:
            break

        if cmd == "stop":
            conn.send(("ok", True))
            break

        try:
            dll = _load_dll()

            if cmd == "open":
                f = dll.fm2LibOpenDriver
                f.restype = ctypes.c_int32
                f.argtypes = [ctypes.c_int16]
                meter_id = f(ctypes.c_int16(int(payload["device_idx"])))
                conn.send(("ok", int(meter_id)))

            elif cmd == "close":
                f = dll.fm2LibCloseDriver
                f.restype = ctypes.c_int16
                f.argtypes = [ctypes.c_int32]
                rc = f(ctypes.c_int32(int(payload["meter_id"])))
                conn.send(("ok", int(rc)))

            elif cmd == "sync":
                f = dll.fm2LibSync
                f.restype = ctypes.c_int16
                f.argtypes = [ctypes.c_int32]
                rc = f(ctypes.c_int32(int(payload["meter_id"])))
                conn.send(("ok", int(rc)))

            elif cmd == "send_command":
                f = dll.fm2LibPackagedSendReply
                f.restype = ctypes.c_int16
                f.argtypes = [
                    ctypes.c_int32,  # Meter ID
                    ctypes.c_char_p,  # Send Command
                    ctypes.POINTER(ctypes.c_char),  # Meter Reply
                    ctypes.POINTER(ctypes.c_int16),  # Reply size
                ]

                buffer_len = int(payload["buffer_len"])
                reply_buffer = ctypes.create_string_buffer(buffer_len)
                size = ctypes.c_int16(buffer_len)

                rc = f(
                    ctypes.c_int32(int(payload["meter_id"])),
                    str(payload["command"]).encode("ascii"),
                    reply_buffer,
                    ctypes.byref(size),
                )

                raw = bytes(reply_buffer)
                reply = raw.split(b"\x00", 1)[0].decode("ascii", errors="replace")
                conn.send(("ok", {"rc": int(rc), "reply": reply, "size": int(size.value)}))

            elif cmd == "get_serial_number":
                f = dll.fm2LibGetSerialNumber
                f.restype = ctypes.c_int16
                f.argtypes = [
                    ctypes.c_int32,
                    ctypes.POINTER(ctypes.c_char * 16),
                    ctypes.POINTER(ctypes.c_int16),
                ]
                returnBuffer = (ctypes.c_char * 16)()
                size = ctypes.c_int16(16)
                rc = f(ctypes.c_int32(int(payload["meter_id"])), returnBuffer, ctypes.pointer(size))
                conn.send(
                    (
                        "ok",
                        {
                            "rc": int(rc),
                            "serial": returnBuffer.value.decode(errors="replace"),
                        },
                    )
                )

            elif cmd == "get_data":
                f = dll.fm2LibGetData
                f.restype = ctypes.c_int16
                out_array = (ctypes.c_uint8 * 64)()
                addr = ctypes.c_int16(int(payload["addr"]))
                rc = f(
                    int(payload["meter_id"]),
                    ctypes.pointer(out_array),
                    ctypes.pointer(addr),
                )
                conn.send(("ok", {"rc": int(rc), "raw": bytes(out_array)}))

            elif cmd == "zero_start":
                f = dll.fm2LibZeroStart
                f.restype = ctypes.c_int16
                f.argtypes = [ctypes.c_int32]
                rc = f(ctypes.c_int32(int(payload["meter_id"])))
                conn.send(("ok", int(rc)))

            elif cmd == "zero_reply":
                f = dll.fm2LibGetZeroReply
                f.restype = ctypes.c_int16
                f.argtypes = [ctypes.c_int32]
                rc = f(ctypes.c_int32(int(payload["meter_id"])))
                conn.send(("ok", int(rc)))

            else:
                conn.send(("err", f"Unknown worker command: {cmd}"))

        except Exception as e:
            conn.send(("err", repr(e)))


class _DriverProcess:
    """Manage the worker process that owns the DLL handle."""

    def __init__(self, dll_path: str):
        """Start a fresh worker process for the supplied DLL path."""
        self.dll_path = os.path.abspath(dll_path)
        self.parent_conn = None
        self.proc = None
        self.start()

    def start(self):
        """Spawn the worker process and wait for its initialization signal."""
        ctx = multiprocessing.get_context("spawn")
        parent_conn, child_conn = ctx.Pipe()

        proc = ctx.Process(
            target=_driver_worker,
            args=(child_conn, self.dll_path),
            daemon=False,
        )
        proc.start()

        self.parent_conn = parent_conn
        self.proc = proc

        if self.parent_conn.poll(10.0):
            status, value = self.parent_conn.recv()
            if status != "init_ok":
                raise RuntimeError(f"Failed to start DLL worker: {value}")
        else:
            exitcode = proc.exitcode
            self.terminate()
            raise TimeoutError(f"Timed out starting DLL worker process (exitcode={exitcode})")

    def terminate(self):
        """Terminate the worker process immediately and drop local handles."""
        try:
            if self.proc is not None and self.proc.is_alive():
                self.proc.terminate()
                self.proc.join(timeout=1.0)
        finally:
            self.proc = None
            self.parent_conn = None

    def restart(self):
        """Restart the worker process after a timeout or crash."""
        self.terminate()
        self.start()

    def request(self, cmd: str, payload: dict, timeout_s: float | None = None):
        """Send a command to the worker and optionally enforce a timeout."""
        if self.proc is None or self.parent_conn is None or not self.proc.is_alive():
            self.restart()

        self.parent_conn.send((cmd, payload))  # type:ignore

        if timeout_s is None:
            return self.parent_conn.recv()  # type:ignore

        if self.parent_conn.poll(timeout_s):  # type:ignore
            return self.parent_conn.recv()  # type:ignore

        self.restart()
        return "timeout", None  ############
        # raise TimeoutError(f"Worker timed out on command '{cmd}' after {timeout_s} s")

    def stop(self):
        """Ask the worker to stop cleanly, then terminate if needed."""
        try:
            if self.proc is not None and self.parent_conn is not None and self.proc.is_alive():
                try:
                    self.parent_conn.send(("stop", {}))
                    if self.parent_conn.poll(0.5):
                        self.parent_conn.recv()
                except Exception:
                    pass
        finally:
            self.terminate()


class power_meter_handler:
    """High-level interface for talking to a Coherent FieldMax II meter."""

    def __init__(self, dll_path: str | None = None):
        """Create a meter handler for a local or globally installed DLL.

        Args:
            dll_path: Path to `FieldMax2Lib.dll`. If `None`, the code first
                checks for a repo-local `FieldMax2Lib.dll` next to this module,
                then falls back to the expected global install path.
        """
        resolved_dll_path = _resolve_dll_path(dll_path)

        self._driver_proc = _DriverProcess(resolved_dll_path)
        self._connected_meter_id = None
        atexit.register(self.final_shutdown)

    def send_command(
        self, command: str, print_error=True, buffer_len: int = 100, sync=True, timeout_s: float | None = None
    ):
        """Send a low-level packaged command to the meter and return its reply.
        
        See "fieldmaxii-labview-examples.zip/Getting Started FMII LV.pdf" for commands: 
        https://www.coherent.com/content/dam/coherent/site/en/resources/laser-measurement-and-control-help-center/software-drivers-and-manuals/fieldmax-ii/fieldmaxii-labview-examples.zip
        """

        if self._connected_meter_id is None:
            print("Connect to FieldMax power meter before sending commands.")
            return None
        else:
            try:
                result = self._request(
                    "send_command",
                    {
                        "meter_id": self._connected_meter_id,
                        "command": command,
                        "buffer_len": buffer_len,
                    },
                    timeout_s=timeout_s,
                )
                if result is None:
                    return None

                rc = result["rc"]  # rc seems to always return -1
                reply = result["reply"]

                # print("send_command reply:", reply)
                if rc != -1:
                    print(
                        f"[Info] rc is not -1 for meter_id {self._connected_meter_id}, command {command}, buffer len {buffer_len}, timeout {timeout_s}"
                    )

                if sync == True:
                    self._sync()  # seems to be needed

                return reply

            except Exception as e:
                if print_error == True:
                    error_print(f"[Error] Failed to send command to FieldMax power meter: {e}")
                return None

    def connect(
        self,
        device_idx: int = 0,
        print_error: bool = True,
        sync: bool = True,
        timeout_s: float | None = 5.0,
    ) -> bool:
        """Open the device at `device_idx` and verify that it responds."""
        if self._connected_meter_id is not None:
            self.disconnect(print_error=print_error, timeout_s=2)
        try:
            meter_id = self._request(
                "open",
                {"device_idx": int(device_idx)},
                timeout_s=timeout_s,
            )

            if (
                meter_id != -1
            ):  # -1 means fail but not -1 does not guarante working connection -> test via get_serial_number()
                self._connected_meter_id = meter_id

                connected = self.is_confirmed_connected(timeout_s=2)
                if connected == False:
                    if print_error:
                        error_print(f"[Error] Failed to connect to FieldMax power meter on device index {device_idx}")
                    self._connected_meter_id = None
                    return False
                else:
                    if sync:
                        self._sync(timeout_s=timeout_s)
                    return True
            else:
                if print_error:
                    error_print(f"[Error] Failed to connect to FieldMax power meter on device index {device_idx}")
                self._connected_meter_id = None
                return False

        except Exception as e:
            self._connected_meter_id = None
            if print_error:
                error_print(f"[Error] Failed to connect to FieldMax power meter: {e}")
            return False

    def disconnect(
        self,
        print_error: bool = True,
        timeout_s: float | None = 3.0,
    ) -> bool:
        """Close the current meter connection, if one is active."""
        if self._connected_meter_id is None:
            return True

        try:
            self._request(
                "close",
                {"meter_id": int(self._connected_meter_id)},
                timeout_s=timeout_s,
            )
            return True

        except Exception as e:
            if print_error:
                error_print(f"[Error] Failed to disconnect FieldMax power meter: {e}")
            return False

        finally:
            self._connected_meter_id = None

    def get_meter_id(self) -> int | None:
        """Return the driver-assigned ID for the active meter connection."""
        return self._connected_meter_id

    def get_serial_number(self, print_error=True, timeout_s: float | None = 5):
        """Return the connected meter serial number, or `None` on failure."""
        if self._connected_meter_id is not None:
            try:
                result = self._request("get_serial_number", {"meter_id": self._connected_meter_id}, timeout_s=timeout_s)
                if result is None:
                    return None

                serial = result["serial"]

                if serial is not None and serial != "":
                    return serial
                else:
                    return None
            except Exception as e:
                if print_error == True:
                    error_print(f"[Error] Failed to get FieldMax power meter serial number: {e}")
                return None
        else:
            if print_error == True:
                error_print("[Error] Connect to FieldMax power meter first.")
            return None

    def set_current_power_to_0(self, print_error=True) -> None:
        """Run the meter's zeroing routine until it reports completion."""
        if self._connected_meter_id is None:
            print("[Error] Open connection to FieldMax power meter before zeroing.")
        else:
            try:
                self._zeroing_start()
                ans = self._zeroing_reply()
                while ans == 1:
                    ans = self._zeroing_reply()
            except Exception as e:
                if print_error == True:
                    error_print(f"[Error] Failed to set zero for FieldMax power meter: {e}")

    def is_connected(self) -> bool:
        """Return `True` when a meter ID is currently stored locally."""
        return self._connected_meter_id is not None

    def is_confirmed_connected(self, timeout_s=2) -> bool:
        """Verify the connection by fetching the serial number."""
        try:
            sn = self.get_serial_number(timeout_s=timeout_s, print_error=False)
            if sn is None:
                return False
            else:
                return True
        except Exception:
            return False

    def read_power_W(
        self, print_error: bool = True, retries: int = 5, retry_delay_s: float = 0.05, timeout_s: float | None = 5
    ) -> tuple[float, float, float] | tuple[None, None, None]:
        """Return min/mean/max in watts from the latest valid power samples."""
        if self._connected_meter_id is None:
            print("[Error] Open connection to FieldMax power meter before reading data.")
            return None, None, None
        else:
            try:
                for _ in range(retries + 1):
                    data = self._read_power_array_W(timeout_s=timeout_s)  # returns 8 floats

                    if data is None:
                        if print_error:
                            error_print("[Error] Failed to get FieldMax power meter power.")
                        return None, None, None

                    real_data = [value for value in data if value != 0]
                    if real_data:
                        # 0 means no data
                        return (
                            min(real_data),
                            math.fsum(real_data) / len(real_data),
                            max(real_data),
                        )
                    else:
                        time.sleep(retry_delay_s)
                else:
                    if print_error:
                        error_print("[Error] Failed to get real FieldMax power meter power.")
                    return None, None, None
            except Exception as e:
                if print_error:
                    error_print(f"[Error] Failed to read FieldMax power meter power: {e}")
                return None, None, None

    def set_wavelength_nm(self, wavelength_nm, sync=True, timeout_s: float | None = 5, print_error=True):
        """Set the wavelength in nanometers. Returns False if not sucessfully set, including if wavelength was clamped by allowed range."""
        if wavelength_nm is not None:
            response = self.send_command(f"WOO{wavelength_nm}", sync=sync, timeout_s=timeout_s)
            if response != "" and response is not None:
                if float(response.split(",")[0]) != wavelength_nm:
                    if print_error == True:
                        error_print(
                            f"[Warning] Wavelength set to edge of allowd range ({response.split(',')[0]} nm) because requested wavelength ({wavelength_nm} nm) outside the allowed range ({response.split(',')[1]}-{response.split(',')[2]} nm)."
                        )
                    return False
                else:
                    return True
            else:
                if print_error == True:
                    error_print("[Error] failed to set wavelength.")
                return False
        else:
            return True

    def get_wavelength_nm(self, sync=True, timeout_s: float | None = 5) -> None | float:
        """Return the configured wavelength in nanometers."""
        response = self.send_command("WOO", sync=sync, timeout_s=timeout_s)
        if response is not None and response != "":
            return float(response.split(",")[0])
        else:
            return None

    def set_auto_range(self, on=True, sync=True, timeout_s: float | None = 5, print_error=True):
        """Enable or disable the meter's auto-ranging mode."""
        if on is not None:
            response = self.send_command(f"AUT{int(on)}", sync=sync, timeout_s=timeout_s)

            if response is not None and response != "":
                if bool(response) == on:
                    return True
                else:
                    if print_error == True:
                        error_print("[Error] failed to set auto range.")
                    return False
            else:
                if print_error == True:
                    error_print("[Error] failed to set auto range.")
                return False
        return True

    def get_auto_range(self, sync=True, timeout_s: float | None = 5) -> bool | None:
        """Return the current auto-ranging state, or `None` on failure."""
        response = self.send_command("AUT", sync=sync, timeout_s=timeout_s)
        if response is not None and response != "":
            return bool(response)
        else:
            return None

    ###
    # backend methods

    def final_shutdown(self, timeout_s: float | None = 2.0):
        """Disconnect the meter and stop the background DLL worker."""
        try:
            self.disconnect(print_error=False, timeout_s=timeout_s)
        except Exception:
            pass
        try:
            self._driver_proc.stop()
        except Exception:
            pass

    def _request(self, cmd: str, payload: dict, timeout_s: float | None = 5.0):
        """Send a worker request and normalize worker errors and timeouts."""
        status, value = self._driver_proc.request(cmd, payload, timeout_s=timeout_s)
        if status == "err":
            raise RuntimeError(value)
        if status == "timeout":
            self.disconnect()
            return None
        else:
            return value

    def _read_power_array_W(self, timeout_s: float | None = 5):
        """Read the raw sample block and decode the power values in watts."""
        result = self._request("get_data", {"meter_id": self._connected_meter_id, "addr": 8}, timeout_s=timeout_s)
        if result is None:
            return None
        out_array = (ctypes.c_uint8 * 64).from_buffer_copy(result["raw"])
        return self._data_bytes2float(out_array)[0]

    def _sync(self, timeout_s: float | None = 5.0) -> None:
        """Flush pending meter state via the vendor DLL sync call."""
        self._request(
            "sync",
            {"meter_id": self._connected_meter_id},
            timeout_s=timeout_s,
        )

    def _zeroing_start(self, timeout_s: float | None = 5.0):
        """Start the meter zeroing procedure."""
        return self._request("zero_start", {"meter_id": self._connected_meter_id}, timeout_s=timeout_s)

    def _zeroing_reply(self, timeout_s: float | None = 5.0):
        """Poll the meter for the current zeroing status."""
        return self._request("zero_reply", {"meter_id": self._connected_meter_id}, timeout_s=timeout_s)

    def _data_bytes2float(self, l):
        """Decode the DLL data block into interleaved power and period arrays."""
        float_p = ctypes.cast(l, ctypes.POINTER(ctypes.c_float))
        power = [float_p[0], float_p[2], float_p[4], float_p[6], float_p[8], float_p[10], float_p[12], float_p[14]]
        period = [float_p[1], float_p[3], float_p[5], float_p[7], float_p[9], float_p[11], float_p[13], float_p[15]]
        return (power, period)
