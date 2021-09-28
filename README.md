# Get and Save

A python 'lib' especially to be used for a myDRE Workspace with functions to:
- fetch zipped uploads from z:\inbox and move them to a predefined location
- unzip all files in the predefined location
- select the folder to be used (user interaction)
- select the file to be used (user interaction)
- save a dataframe in a predefined location as CSV or Excel

Leaving the middle place to turn the CSV file into a dataframe and do with
the data that must be done.

An example of how to use it, is added; both Jupyter Notebook and just python script.

The following packeges need to be available:
os
openpyxl
shutil
datetime
time
glob
zipfile

Instructions for myDRE:
- deploy VM with OSDS template
- in the windows VM
- open cmd miniconda
- conda install -c conda-forge oopenpyxl, shutil, glob, zipfile
