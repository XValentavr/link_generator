import base64
import hashlib
from urllib.parse import quote

import openpyxl


def generate_unique_link(domain, reference_number, email, name, secret_key):
    base64_email = base64.b64encode(email.encode()).decode()
    encoded_name = quote(name)
    data_to_hash = f"{secret_key}{email}{reference_number}".encode()
    hashed_data = hashlib.sha1(data_to_hash).hexdigest()

    return f"https://se.trustpilot.com/evaluate/{domain}?a={reference_number}&b={base64_email}&c={encoded_name}&e={hashed_data}"


while True:
    print('Enter path to file')
    path = str(input())

    # read a file from entered path
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active

    new_column_letter = chr(ord('A') + 4)
    # need to write in correct column
    column_count = 0
    for i, row in enumerate(sheet.iter_rows(min_row=1, values_only=True), start=1):
        email = row[0]
        if email is not None:
            unique_link = generate_unique_link(domain="prodiga.se", reference_number=row[2], email=row[0], name=row[1],
                                               secret_key=row[3])
            sheet[new_column_letter + str(column_count + 1)] = unique_link
            column_count += 1
        else:
            print("Email value is empty. Skipping this row.")

    workbook.save(path)

    print("Done! Unique links written to the new column in 'test.xlsx'.")
    break
