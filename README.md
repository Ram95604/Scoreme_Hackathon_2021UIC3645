# Scoreme_Hackathon_2021UIC3645
PDF Table Extractor
This repository contains Python scripts to extract tables from PDF files and save them into Excel format. Two different approaches are provided to handle distinct PDF structures:

test3.pdf - Uses PyMuPDF to extract tables based on text block coordinates.
test6.pdf - Uses pdfplumber to extract text and processes it into structured transaction data.
Both scripts have been tested and run successfully on Google Colab.

Features
Extracts tables from PDFs and saves them as Excel files.
Handles different PDF formats (e.g., simple tables in test3.pdf and bank statements in test6.pdf).
Cleans and structures data for bank statements with transaction details.
Prerequisites
To run these scripts on Google Colab, ensure you have the following:

A Google account to access Google Colab.
PDF files (test3.pdf and test6.pdf) uploaded to your Google Drive or Colab environment.
Dependencies
The scripts rely on the following Python libraries:

PyMuPDF==1.23.5 (for test3.pdf)
pdfplumber (for test6.pdf)
pandas (for both)
openpyxl (for Excel output)
You can install these dependencies directly in Google Colab using the commands below.

Setup and Installation
Follow these steps to run the scripts in Google Colab:

Open Google Colab:
Go to Google Colab.
Create a new notebook (File > New Notebook).
Install Dependencies:
Run the following commands in a code cell to install the required libraries:
bash

Collapse

Wrap

Copy
!pip install PyMuPDF==1.23.5
!pip install pdfplumber
!pip install pandas
!pip install openpyxl
Upload PDF Files:
Upload your PDF files (test3.pdf and/or test6.pdf) to the Colab environment:
