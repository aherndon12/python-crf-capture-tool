# Python CRF Capture Tool

A tool that automates the extraction and formatting of Case Report Forms (CRFs) from Medidata Rave EDC, producing clear PDF outputs for clinical trial database builds.

## Features

* Automated login and session handling for Medidata Rave EDC
* Full-page screenshot capture and PDF generation of CRFs
* Cleaned UI capture with dropdown options displayed explicitly
* Consolidation of individual CRFs into a single merged PDF casebook

## Requirements

* Python 3.x
* Chrome browser and ChromeDriver

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/aherndon12/python-crf-capture-tool.git
   ```

2. Install required Python packages:
   ```bash
   pip install -r requirements.txt
   ```

3. Create a `.env` file in the project root and add your credentials:
   ```
   IMEDIDATA_USERNAME=your_username
   IMEDIDATA_PASSWORD=your_password
   ```

## Usage

1. Populate `URLs.xlsx` with your CRF names and URLs
2. Execute the script:
   ```bash
   python app.py
   ```
3. Output PDFs will be saved in the generated `output_<date>` directory

## Project Structure

```
.
├── app.py
├── .env
├── .gitignore
├── requirements.txt
├── URLs.xlsx
├── chromedriver-win64/
└── output_<date>/ (auto-generated upon execution)
```

## Security Notes

* The `.env` file contains sensitive credentials and is excluded from version control
* Never commit your `.env` file to version control
* Keep your credentials secure and don't share them

## Contact

* Andrew Herndon
* [LinkedIn](https://www.linkedin.com/in/andrew-herndon)

*Presented at PharmaSUG 2025 – Paper SD-339.*
