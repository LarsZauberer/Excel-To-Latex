# Excel-To-Latex
## Converts an excel file to an latex table and copies the code to the clipboard.

# Installation from source
1. Clone the repository or download the `.zip` file [here]()
2. Run the command in side the folder: 
```
pip install -r requirements.txt
```

# Usage
Run the `main.py` with the argument `-f` and specify an
excel file.

It should now have successfully converted the excel file to latex code.

! Important: The latex code uses the package `longtable`. Make sure you add this line to your latex file:
```Latex
\usepackage{longtable}
```
