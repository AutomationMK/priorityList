# Generate Priority List Job Tracker

## Description

This is a simple project to generate the excel file for daily open jobs.
An excel file from previously made jobs needs to be available if not all old
ship dates given will no longer be included in the list and a new blank file will be generated.

## Getting Started

### Dependencies

* Python 3
* pandas version(2.2.3)
* numpy version(1.26.4)
* openpyxl version(3.1.3)
* WSL (Windows Subsystem For Linux) or any linux distribution

### Installing

* All actions need to be executed in the main project folder of running from command line
* The program also needs it's virtual environment activated
* pip will need to be ran for installing any dependencies in the requirements.txt file

Example running on Linux...

```bash
# in the main project folder as the working directory
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
playwright install chromium
deactivate
```

### Executing The Program

* Finally run the program by main.py from the main project folder
* You need to activate the python virtual environment

Example running on Linux...
```bash
# in the main project folder as the working directory
source venv/bin/activate
python main.py
deactivate
```

### Authors

Max Kranker (<max.kranker@colecarbide.com>)

## Version History

* 0.1
  * Initial Release

## License

This project is licensed under the MIT License - see the License.md file for details
