import os
from datetime import datetime
import pandas as pd
import re
from pypdf import PdfReader
import datefinder
from rapidfuzz import process
from dateutil import parser as date_parser

excel_path = input(r"Enter the path to the Excel file: ")
owner_df = pd.read_excel(excel_path, sheet_name= input("Please enter the name of the spreadsheet: "), header=1)