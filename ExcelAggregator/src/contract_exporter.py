import logging
import os
import re
import warnings

from openpyxl import Workbook

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

base_dir = os.path.dirname(os.path.abspath(__file__))

CSV_DIRECTORY = os.path.join(base_dir, 'outputs', 'tsv')
OUTPUT_DIRECTORY = os.path.join(base_dir, 'outputs', 'xlsx')
OUTPUT_FILE = os.path.join(OUTPUT_DIRECTORY, 'centralized_data.xlsx')

keys_patterns = {
    "Contract Period": r"CONTRACT PERIOD\s+(.+)",
    "Supplier Name": r"SUPPLIER NAME\s+(.+)",
    "Vendor VAT number": r"VENDOR VAT NUMBER\s+(.+)",
    "SAP number vendor": r"SAP NUMBER VENDOR\s+(.+)",
    "Vendor street": r"STREET\s+(.+)",
    "Vendor number": r"NUMBER\s+(.+)",
    "Vendor postal code": r"POSTAL CODE\s+(.+)",
    "Vendor city": r"CITY\s+(.+)",
    "Vendor country": r"COUNTRY\s+(.+)",
    "Payment terms": r"PAYMENT TERMS\s+(.+)",
    "Delivery Period from": r"FROM\s+([\d-]+ [\d:]+)",
    "Delivery Period to": r"TO\s+([\d-]+ [\d:]+)",
    "Renewal": r"RENEWAL\s+(.+)",
    "Invoicing": r"INVOICING\s+(.+)",
    "Begin/end period": r"BEGIN/END PERIOD\s+(.+)",
    "Index": r"!!! Index\s+(.+)",
    "Monthly fee per user": r"YEAR \d\s+(.+?/year)",
    "Additional fee": r"ADDITIONAL FEE\s+CALCULATION\s+\(describe pls\)\s+(.+)",
    "Number of subscribers": r"NUMBER OF SUBSCRIBERS\s+OTHER =\s+\(explain briefly\)\s+(.+)",
    "Index calculation": r"INDEX\s+CALCULATION OF INDEX\s+(.+)"
}


def extract_data(pattern, text, multiple=False):
    if multiple:
        return re.findall(pattern, text)
    else:
        match = re.search(pattern, text)
        return match.group(1).strip() if match else ""


def main():
    logging.info("Starting the script")

    wb = Workbook()
    ws = wb.active
    ws.title = "Centralized Data"

    os.makedirs(OUTPUT_DIRECTORY, exist_ok=True)

    max_channels = 0
    for filename in os.listdir(CSV_DIRECTORY):
        if filename.endswith(".tsv"):
            file_path = os.path.join(CSV_DIRECTORY, filename)
            try:
                with open(file_path, 'r', encoding='utf-8') as file:
                    content = file.read()
            except UnicodeDecodeError:
                with open(file_path, 'r', encoding='latin-1') as file:
                    content = file.read()

            channels = extract_data(r"(.+?)\s+\((.+?)\)", content, multiple=True)
            filtered_channels = [channel for channel in channels if channel[0] not in ["NEW", "old"]]
            if len(filtered_channels) > max_channels:
                max_channels = len(filtered_channels)

    logging.info(f"Maximum number of channels: {max_channels}")

    basic_headers = ["Filename"] + list(keys_patterns.keys())
    channel_headers = []
    for i in range(max_channels):
        channel_headers.append(f"Channel {i + 1}")
        channel_headers.append(f"Packs Channel {i + 1}")

    year_headers = [f"YEAR {i + 1} FIXED FEE IN €" for i in range(4)]
    headers = basic_headers + year_headers + channel_headers

    ws.append(headers)

    for filename in os.listdir(CSV_DIRECTORY):
        if filename.endswith(".tsv"):
            file_path = os.path.join(CSV_DIRECTORY, filename)
            try:
                with open(file_path, 'r', encoding='utf-8') as file:
                    content = file.read()
            except UnicodeDecodeError:
                with open(file_path, 'r', encoding='latin-1') as file:
                    content = file.read()

            row = [filename]
            yearly_fees = {}
            for key in keys_patterns.keys():
                if key != "YEARLY FIXED FEE IN €":
                    value = extract_data(keys_patterns[key], content)
                    row.append(value)

            yearly_fee_pattern = re.compile(r"YEAR (\d+)\s*(\d+)?")
            matches = yearly_fee_pattern.findall(content)
            yearly_fees = {f"YEAR {year} FIXED FEE IN €": fee for year, fee in matches if fee}

            for header in year_headers:
                fee = yearly_fees.get(header, "")
                if fee and len(fee) > 1:
                    row.append(fee)
                else:
                    row.append("")

            channels = extract_data(r"(.+?)\s+\((.+?)\)", content, multiple=True)
            filtered_channels = [channel for channel in channels if channel[0] not in ["NEW", "old"]]
            channels_row = []
            for channel in filtered_channels:
                channels_row.append(channel[0])
                channels_row.append(', '.join([pack.strip() for pack in channel[1].split(',')]))

            while len(channels_row) < len(channel_headers):
                channels_row.append("")

            final_row = row + channels_row
            ws.append(final_row)

    if os.path.exists(OUTPUT_FILE):
        os.remove(OUTPUT_FILE)

    wb.save(OUTPUT_FILE)

    logging.info(f"Centralized Excel file created at: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
