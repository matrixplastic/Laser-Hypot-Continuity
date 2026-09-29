# Laser Hypot Continuity

## Table of Contents

1. [Overview](#overview)
2. [Features](#features)
3. [Usage](#usage)
4. [Technical](#technical)
5. [Syntax](#syntax)
6. [Configuration](#configuration)
7. [Contributions](#contributions)
8. [Installation](#installation)
9. [Help](#help)
10. [License](#license)

## Overview

This project is designed to run a 10-cavity fixture through a variety of tests and functions, including:

- Hypot Test
- Continuity Test
- Laser Marking

## Features

- Automated testing for Hypot and Continuity
- Laser marking functionality
- Detailed logging and reporting
- Configurable settings for different test parameters

## Usage

- Double click .exe to start the program
- Left column: the Devices box shows each tester, switch, the laser marker and the machine-data
  link with a green or red dot; Messages shows why a run did not start; START and Emergency STOP
- Change devices... opens a panel that lists the USB serial devices attached right now. Pick one
  for a role and press Connect: the driver is swapped in place, no restart, and the serial number
  is saved to settings.ini. The laser marker's IP can be changed and reconnected the same way.
  Not available while a test is running
- Middle: progress bar and the state of each cavity
- Right column: the test settings the next run will use, read only. A value in amber differs from
  the program's default, which means settings.ini has drifted from the code
- Admin settings (password default 6789, changeable in settings.ini) enable or disable the tests
  or lasering per cavity and edit the test parameters; saving them updates the read-only view and
  the next run at once
- Emergency STOP disables as much as possible and force closes the program
- tests/ui_smoke.py runs the whole window without hardware (`xvfb-run python tests/ui_smoke.py`,
  scenarios connected, missing, fail) and writes screenshots; use it before a build

## Technical

- Python 3.12.3
- **Hypot Model:** Hypot3805 <https://www.arisafety.com/model-3805.html>
- **Hypot Switch Model:** SC6540 <https://www.arisafety.com/model-sc6540.html>
- **Laser Model:** Keyence 3 Axis MD-X1000 Laser Marker <https://www.keyence.com/products/marker/laser-marker/md-x1000_1500/models/md-x1000/>
- Developed and tested on Windows 11 Pro

## Syntax

- **Variables**: camelCase
- **Functions**: snake_case

## Configuration

- settings.ini sits next to the exe (not in the working directory) and is created on first run
- [Hypot] holds the test parameters. They are controlled process values: keep them equal to the
  LHC work instruction, and note the WI revision in the comment above `defaultHypotSettings`
  in Main.py when they change
- Every run logs the settings it used and reads them back from both testers before the first
  cavity; a mismatch stops the run with a message in the status box
- [Hardware IDs] holds the USB hardware IDs of the two testers and two switches
- [MachineData] turns on publishing to the plant machine-data broker (see below)
- A value that will not parse stops the START button with a message; fix the file and restart

## Machine data

With `[MachineData] enabled = 1` the program publishes to the plant MQTT broker on the
machinedata VM (`host`, `port`, `user`, `password`, `machine_id`): its state, one cycle per
fixture run with good and reject counts and every test parameter the run used, alarms
(settings mismatch, tester timeout, laser not marking, startup errors) and a heartbeat every
minute. That is what the ERP floor view and the OEE rollup read. If the broker is unreachable
the program still runs tests and logs the outage. Details, topics and the server-side steps
are in the ERP repo, `docs/work-instructions/machine-monitor-lhc.md`.

## Contributions

- Made, created, and designed by Tony Martin at Matrix Plastic Products

## Installation

- It is recommended to disable "Allow the computer to turn off this device to save power" in control panel on the USB Hubs, 
or else they can lose connection and need to be unplugged and replugged<br>

- To install this project, clone the repository and install the required dependencies:

  - Install IVI Drivers <https://www.ivifoundation.org/Shared-Components/default.html>
  - Install Serial Drivers <https://www.ni.com/en/support/downloads/drivers/download.ni-visa.html#565016>
  - Install Hardware Drivers according to your model. Instructions are also included <https://www.arisafety.com/support/instrument-drivers>

```sh
git clone https://github.com/matrixplastic/Laser-Hypot-Continuity.git
cd /Path/To/Your/Cloned/Project
pip install -r requirements.txt
```

## Building the exe

The machine runs a PyInstaller build. Build from a clean checkout of the commit being released,
keep the previous exe next to the new one named by date, and copy the machine's settings.ini
and logs folder off before swapping:

```sh
pip install pyinstaller
pyinstaller --onedir --name LaserHypotCont --add-data "SC6540.dll;." --add-data "ARI38XX_64.dll;." Main.py
```

Adjust the `--add-data` entries to wherever the two driver DLLs live on the build PC. Commit the
generated `LaserHypotCont.spec` so the next build is the same build.

## Releasing a test-parameter change

Any change to [Hypot] values or to `defaultHypotSettings` is a controlled process change: it
needs the work instruction revised first, a golden-sample run on all three fixtures on the new
build with the tester display checked against the WI, and sign-off recorded before the exe
goes on the machine.

## Help
The error codes are listed in the *.chm file which should be at C:\Program Files\IVI Foundation\IVI\Drivers\ARI38XX. Please refer to the following pages if you would like to know more about the errors,

ARI38X IVI-COM Driver -> Reference -> Errors and Warnings

ARI38X IVI-COM Driver -> Reference -> Driver Hierarchy -> IARI38XX -> Utility -> ErrorQuery


## License

MIT License
