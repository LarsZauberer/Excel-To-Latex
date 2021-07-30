import logging
import sys
from pathlib import Path
try:
    from rich.logging import RichHandler
    from rich.progress import track
except ModuleNotFoundError:
    logging.warning("Rich is not installed. Please install rich by typing: pip install rich or pip -r requirements.txt")
    sys.exit(1)
import argparse


parser = argparse.ArgumentParser()
parser.add_argument("-f", help="Specify your excel spreadsheed", action="store")
parser.add_argument("-v", help="Verbose", action="store_true", default=False)

args = parser.parse_args()

if args.v:
    FORMAT = "%(message)s"
    logging.basicConfig(
        level="DEBUG", format=FORMAT, datefmt="[%X]", handlers=[RichHandler()]
    )
else:
    FORMAT = "%(message)s"
    logging.basicConfig(
        level="INFO", format=FORMAT, datefmt="[%X]", handlers=[RichHandler()]
    )

log = logging.getLogger()
log.debug(f"Logging initialized!")

if args.f is None:
    log.warning(f"No excel file specified. Please use --help to get more information")
    sys.exit(1)
    
try:
    Path(args.f)
except Exception:
    log.exception(f"Not a valid file path")

try:
    import pandas as pd
except ModuleNotFoundError:
    log.warning(f"pandas module not installed")
    sys.exit(1)

try:
    from pylatex import Document, LongTable, MultiColumn
except ModuleNotFoundError:
    log.warning(f"pylatex module not installed")
    sys.exit(1)

try:
    data = pd.read_excel(Path(args.f))
except ImportError:
    log.warning(f"Please install openpyxl")


doc = Document()

keys = data.keys()
form = ""
for i in keys:
    form += "l "

form.strip()
log.debug(f"Form: {form}")

""" for _, row in data.iterrows():
    for i in row:
        print(i) """

with doc.create(LongTable(form)) as data_table:
    # Code from https://jeltef.github.io/PyLaTeX/current/examples/longtable.html
    data_table.add_hline()
    data_table.add_row([i for i in keys])
    data_table.add_hline()
    data_table.end_table_header()

    data_table.add_hline()
    data_table.add_row((MultiColumn(len(keys), align='r',
                        data='Continued on Next Page'),))
    data_table.add_hline()
    data_table.end_table_footer()
    data_table.add_hline()
    data_table.add_row((MultiColumn(len(keys), align='r',
                        data='Not Continued on Next Page'),))
    data_table.add_hline()
    data_table.end_table_last_footer()
    
    log.debug(f"Finished creating structure of the table")
    
    for _, row in track(data.iterrows(), total=data.shape[0], description="Adding Data"):
        d = []
        for i in row:
            d.append(i)
        data_table.add_row(d)
    
    log.info(f"Successfully added data")

log.debug(f"Creating File")
doc.generate_pdf("file", clean_tex=False)
