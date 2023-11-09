# Accounting

Accounting is a simple accounting program for Windows. It allows users to create reports for the current shift on 
trading in financial instruments.
## Screenshots

![Screenshot 1](/screenshots/Screenshot_1.png)

## Installation

To install Accounting, follow these steps:

1. Clone the repository to your local machine.
2. Create a virtual environment using the `requirements.txt` file:
```
python3 -m venv venv
.\venv_accounting\Scripts\activate
pip install -r requirements.txt
```
3. Build an executable file for Windows using PyInstaller and the `Accounting.spec` file:
```
pyinstaller Accounting.spec
```
4. The executable file will be located in the `dist` directory.

## Usage

To run Accounting, use the following command:
```
python main.py
```

Note: This program is designed to run on Windows. It has not been tested on other operating systems.

## License

This project is licensed under the MIT License - see the LICENSE.md file for details.