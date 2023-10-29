"""
File: make_report.py
Author: Anthony Schrapffer
Description: Main file with example to launch the powerpoint automatic filling.
    How to improve the process ? 
    1) Automatic generation of figure + input data in the portfolio folders
    2) Agilize the process using a combination of bash and run.def file
        to just change name of folder in the run.def to launch the full process.
"""

from src import PresentationCustomized
import os

dir_path = os.path.dirname(os.path.realpath(__file__)) + os.sep
dire_portfolio = dir_path+"Portfolio/Example/"

# If we always used these name, we would just have to specify the dire_portfolio variable
in_data = dire_portfolio + "inputs_description.xlsx"   
out_file = dire_portfolio + "output_presentation.pptx"

##########

pptx = PresentationCustomized(in_data = in_data, dire_portfolio = dire_portfolio, out_file = out_file)

