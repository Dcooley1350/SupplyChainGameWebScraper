# SupplyChainGameWebScraper

A (hacked together) web scraper script written in python to scrape the supply chain simulation for data and output it to excel sheet.

## Usage

Before running, make sure the following env vars are defined:

USERNAME    = "name used to login to the supply chain game
PASSWORD    = "password used to login to the supply chain game
OUTPUT_FILE = "file to dump excel data to(defaults to output.xlsx)"

Make sure to create a virtual env, activate, and install modules before you try to run

```sh
python3 -m venv ./venv
source ./venv/bin/activate
pip install -r requirements.txt
```

Run with `python scrape.py`

Cheers!
