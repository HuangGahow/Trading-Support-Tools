# Trading-Support-Tools

Some basic automational tools for generating confirmation letter and unwind details in FICC Derivatives Trading Internship.

You could use Pyinstaller to transform .py to exe.

Remainder: 

Use a virtual enviroment will avoid creating too large size of file.
Due to Pyinstaller will package all modules in your py environment, creating a new virtual enviroments without installing unnessesary modules will helps a lot.

Example:

#create virtual environment
conda create -n aotu python=3.6

#activate virtual environment
conda activate aotu

#delete virtual environment
conda remove -n aotu--all

#Pyinstaller packaginhg
Pyinstaller -F -w -i apple.ico(image) py_word.py
