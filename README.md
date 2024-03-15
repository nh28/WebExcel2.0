# WebExcel2.0
WebExcel in Python

## Requirements

- [ ] In requirements.txt
    - [ ] 
- [ ] Other
    - [ ] PortableGit or Git installed
    - [ ] MUST be on VPN

## Clone and Configure (using Git)

```
>>> git clone https://github.com/nh28/WebExcel2.0
```

## Set Up

create a virtual environment and download the necessary libraries
Note: if you are using PortableGit, in order to be able to pip install you need to open the command prompt git-cmd.exe

```
>>> cd WebExcel2.0
>>> py -3 -m venv .venv
>>> cd .venv
>>> cd Scripts
>>> activate.bat
>>> # install all the libraries from requirements.txt
>>> pip install -r requirements.txt

```

## Confirming Installation
You can make use of pip freeze to see what libraries you have installed in your virtual environment
If using PortableGit, in order to be able to pip you need to open the command prompt git-cmd.exe

```
>>> cd WebExcel2.0
>>> cd .venv
>>> cd Scripts
>>> activate.bat
>>> pip freeze > requirements.txt
>>> type requirements.txt
```



