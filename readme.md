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
- Left side START button starts the tests
- Emergency STOP button disables as much as possible, and force closes the program
- Right half shows progress bar of tests, as well as current settings on each cavity below it
- Admin panel to modify those settings to enable/disable the tests or lasering on each cavity
- Password default is 6789, can be changed in settings.ini

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

- settings.ini to modify settings of hypot, continuity, and laser
- settings.ini also for modifying hardware IDs before program launch

## Contributions

- Made, created, and designed by Tony Martin at Matrix Plastic Products

## Installation

It is recommended to disable "Allow the computer to turn off this device to save power" in control panel on the USB Hubs, 
or else they can lose connection and need to be unplugged and replugged<br>
To install this project, clone the repository and install the required dependencies:

Install IVI Drivers <https://www.ivifoundation.org/Shared-Components/default.html>
Install Serial Drivers <https://www.ni.com/en/support/downloads/drivers/download.ni-visa.html#565016>
Install Hardware Drivers according to your model. Instructions are also included <https://www.arisafety.com/support/instrument-drivers>

```sh
git clone https://github.com/matrixplastic/Laser-Hypot-Continuity.git
cd /Path/To/Your/Cloned/Project
pip install -r requirements.txt
```

## Help
The error codes are listed in the *.chm file which should be at C:\Program Files\IVI Foundation\IVI\Drivers\ARI38XX. Please refer to the following pages if you would like to know more about the errors,

ARI38X IVI-COM Driver -> Reference -> Errors and Warnings

ARI38X IVI-COM Driver -> Reference -> Driver Hierarchy -> IARI38XX -> Utility -> ErrorQuery


## License

MIT License
