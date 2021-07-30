# Excel-To-Latex
## Converts an excel file to an latex table and copies the code to the clipboard.

# Installation as executable for windows
1. Download the `.zip` file [here]()
2. Extract it

We recommend adding the folder to your PATH to have easy access to it

# Installation from source
1. Clone the repository
2. Run the command: 
```
pip install -r requirements.txt
```

# Usage
Run the `excelLatex.exe` or `main.py` with the argument `-f` and specify an
excel file.

It should now have successfully converted the excel file to latex code.

! Important: The latex code uses the package `longtable`. Make sure you add this line to your latex file:
```Latex
\usepackage{longtable}
```
