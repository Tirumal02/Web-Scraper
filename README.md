# Web-Scraper
# TGSRTC Service Monitor

This Python script automates the process of scraping bus service data from the [TGSRTC website](https://www.tgsrtcbus.in), compares it with a reference Excel file, and sends an email notification with a summary of missing and extra services.

## Features

- Uses **Selenium WebDriver** to scrape bus service information from the TGSRTC portal.
- Compares scraped service numbers with a reference dataset (`Service_data.xlsx`).
- Identifies **missing** and **extra** services.
- Sends email alerts with HTML-formatted tables using the **SMTP** protocol.
- Automatically retries scraping in case of errors.

## Requirements

- Python 3.7+
- Google Chrome installed
- ChromeDriver matching your Chrome version

### Python Libraries

You can install the dependencies using:

```bash
pip install pandas selenium tabulate openpyxl
