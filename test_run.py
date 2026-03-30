"""Example script for connecting to a FieldMax meter and logging readings."""

import os
import sys
import time
from datetime import datetime

os.chdir(os.path.dirname(os.path.abspath(__file__)))
from fieldmax_power_meter import error_print, power_meter_handler

def main():
    """Run a simple measurement loop and append results to `test.csv`."""
    dll_path = None  # search local repo copy first, then global install path
    device_idx = 0
    wavelength_nm = 1980  # set None to not change
    auto_range = True  # set None to not change

    time_between_reads_s = 0.5
    log_path = "test.csv"
    num_reads = 3  # None for infinite

    try:
        pm = power_meter_handler(dll_path=dll_path)
        success = pm.connect(device_idx, sync=False, timeout_s=5)
        if success == False:
            print("Failed to connect to PM. Aborting")
            sys.exit()

        success = pm.set_wavelength_nm(wavelength_nm, sync=False)  # set wavelength
        if success == False:
            print("Failed to set wl. Aborting")
            sys.exit()

        success = pm.set_auto_range(auto_range, sync=True)  # set auto range
        if success == False:
            print("Failed to set auto range. Aborting")
            sys.exit()

        if not os.path.exists(log_path):
            with open(log_path, "w", encoding="utf-8") as f:
                f.write("Date (Year.Month.Day)\tTime (Hour:Minute:Second)\tPower (Watt)\n")

        if num_reads is None:  # type:ignore
            infinite = True
            num_reads = 1
        else:
            infinite = False
        ran_once = False
        while (infinite == True) or (ran_once == False):
            for _ in range(num_reads):
                power_mean = pm.read_power_W()[1]
                if power_mean is not None:
                    print(power_mean)
                    ts = datetime.now().astimezone().strftime("%Y.%m.%d\t%H:%M:%S")
                    line = f"{ts}\t{power_mean:.3f}\n"

                    with open(log_path, "a", encoding="utf-8") as f:
                        f.write(line)
                time.sleep(time_between_reads_s)
            ran_once = True

    except Exception as e:
        error_print(f"Except: {e}")
    finally:  # works also for keyborad interrupt
        try:
            pm.final_shutdown()  # type:ignore
        except Exception:
            pass


if __name__ == "__main__":
    main()
