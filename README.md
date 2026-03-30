# fieldmax_power_meter_python

Python wrapper for Coherent FieldMax II power meters on Windows using
`FieldMax2Lib.dll`.

This code started as an adaptation of
[`pyFieldMaxII`](https://github.com/jscman/pyFieldMaxII) by [`jacman`](https://github.com/jscman).

This project provides:

- A small Python API for connecting to a FieldMax II meter.
- A worker-process wrapper around the vendor DLL so DLL hangs do not take down
  the main Python process.
- An example script that logs readings to `test.csv`.

## What This Repo Does

The main module, `fieldmax_power_meter.py`, exposes the
`power_meter_handler` class. It can:

- Connect to a FieldMax II meter by device index.
- Read power values in watts.
- Zero the current power reading.
- Set and read wavelength.
- Enable or disable auto ranging.
- Send low-level packaged commands when needed.

## Requirements

- Windows
- A Coherent FieldMax II meter
- Coherent's `FieldMax2Lib.dll`

## Where To Get The DLL

`FieldMax2Lib.dll` is installed by Coherent's FieldMaxII PC software, which can be downloaded under https://repo.coherent.com/software/FieldMaxII_v3.3.2.9_rc1_setup.exe
and comes from the website https://www.coherent.com/de/laser-power-energy-measurement/meters/fieldmax.

After installation, this the DLL should be at the path ```C:\Program Files (x86)\Coherent\FieldMaxII PC\Drivers\Win10\FieldMax2Lib\x64\FieldMax2Lib.dll```.

You have three supported options after the dll installation:

1. Use `power_meter_handler()` 
   which finds the dll at the default global install path if it was installed there.
2. Copy `FieldMax2Lib.dll` into this repo and use `power_meter_handler()` which finds the dll locally.
3. Copy `FieldMax2Lib.dll` into any folder and provide the full dll path via `power_meter_handler(dll_path="example\\FieldMax2Lib.dll")`.

`FieldMax2Lib.dll` is a third-party vendor file with its own license. Do not
commit it to a public repository or redistribute it unless its own license
allows that. This repo's `.gitignore` already excludes it.

## Basic Usage

```python
from fieldmax_power_meter import power_meter_handler

pm = power_meter_handler(dll_path=None)

if pm.connect(device_idx=0):
    try:
        # pm.set_current_power_to_0() # optional
        power_min_W, power_mean_W, power_max_W = pm.read_power_W()
        print(power_min_W, power_mean_W, power_max_W)
    finally:
        pm.disconnect()
```

## Example Script

`test_run.py` shows a simple logging workflow. It:

- Connects to the first detected device.
- Optionally sets wavelength and auto range.
- Reads power repeatedly.
- Appends timestamped measurements to `test.csv`.

Run it from the repository directory:

```powershell
python .\test_run.py
```
