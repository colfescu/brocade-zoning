# Brocade Zoning tool 

This software creates aliases, zones and zone configs for Brocade switches using an associated Excel spreadsheet.
Single-initiator-multiple-target is the rule used for zoning

## Table of Contents

- [Introduction](#introduction)
- [Features](#features)
- [Installation](#installation)
- [Usage](#usage)
- [Contributing](#contributing)
- [Licence](#licence)
- [Contact](#contact)

## Introduction

This software creates aliases, zones and zone configs for Brocade switches using an associated Excel spreadsheet
The Excel spreadsheet name must be: WWNs Zone Processing.xlsx
The software is written in Python and has one requirement only: openpyxl.
The purpose of the software is to help automate the zoning in large scale environments. 
Up to four SANs can be zoned using this software, but it is also easy to modify and increase that number.

## Features

Main software features:

- the excel spreadsheet doesn't require a specific name - it just needs to be the active sheet in the file
- the active sheet is the sheet Excel opens to when the file is opened
- to make it the active sheet, just open the file in Excel, move to the sheet you want and close Excel
The software:
- creates aliases based on Excel spreadsheet entries
- creates zones
- creates zones configs
- enables zones configs


## Installation

1. Clone the repository:
    ```sh
    git clone https://github.com/colfescu/yourprojectname.git
    ```
2. Navigate to the project directory:
    ```sh
    cd yourprojectname
    ```
3. Install the required dependencies:
    ```sh
    pip install openpyxl
    ```

## Usage

Add the python script and the Excel spreadsheet in the same directory. 
Run the script with:

```sh
python3 brocade-python-v7.py
```

## Licence

This software is licensed under the Apache v2 licence.

## Contact

If you find any bugs and would like to report them, my GitHub username is "colfescu"

