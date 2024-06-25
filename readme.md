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
9. [License](#license)

# Project Overview and Syntax
This was originally created using Python 3.12.3 in 2024

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
- **Hypot Model:** Hypot3805 https://www.arisafety.com/model-3805.html
- **Hypot Switch Model:** SC6540 https://www.arisafety.com/model-sc6540.html
- **Laser Model:** Keyence 3 Axis MD-X1000 Laser Marker https://www.keyence.com/products/marker/laser-marker/md-x1000_1500/models/md-x1000/

## Syntax
- **Variables**: camelCase
- **Functions**: snake_case

## Configuration
- config.yaml to modify settings of hypot, continuity, and laser

## Contributions
- Made, created, and designed by Tony Martin at Matrix Plastic Products

## Installation
To install this project, clone the repository and install the required dependencies:

```sh
git clone https://github.com/yourusername/yourproject.git
cd yourproject
pip install -r requirements.txt
```
## License
None
