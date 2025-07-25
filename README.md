# Generate Priority List Job Tracker

## Description

This is a simple project to generate the excel file for daily open jobs.
Users need to add the old priorityList.xlsx edited file to the main folder.
Make sure that the date info is removed and only priorityList.xlsx is the name.
If not the program will not work correctly and will generate based on the old list in the folder.

## Getting Started

### Dependencies

* Python 3
* pandas version(2.2.3)
* numpy version(1.26.4)
* openpyxl version(3.1.3)
* WSL (Windows Subsystem For Linux), Windows Powershell or any Linux Distrobution

### Installing

* How/where to download your program
* Any modifications needed to be made to files/folders

### Executing program
* The program needs to be executed in the main project folder
* The program also needs it's virtual environment activated
* pip will need to be ran for installing any dependencies in the
requirements.txt file

Example running on Linux...

```bash
# in the main project folder as the working directory
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
playwright install chromium
deactivate
```

Example running in windows powershell...
```powercode
python -m venv venv
.\venv\Scripts\activate.ps1
pip install -r requirements.txt
playwright install chromium
deactivate
```
A powershell script is included for both install and run
and you just need to right click and run the program to install
The IT departmen might need to give you the ability to
run the program at first but after you can run as many times as
you want without permission.

* Finally run the program by main.py from the main project folder
Example running on Linux...

```bash
# in the main project folder as the working directory
source venv/bin/activate
python main.py
deactivate
```

Example running in windows powershell...
```powercode
.\venv\Scripts\activate.ps1
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
