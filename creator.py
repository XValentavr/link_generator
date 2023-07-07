import openpyxl
import base64
import urllib.parse
import hashlib
import uuid

while True:
    print('Enter path to file')
    path = str(input())

    # read a file from entered path
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active

    new_column_letter = chr(ord('A') + sheet.max_column)

    secret_key = str(uuid.uuid4())

    # need to write in correct column
    column_count = 0
    for i, row in enumerate(sheet.iter_rows(min_row=1, values_only=True), start=1):
        email = row[0]
        if email is not None:

            email_byte = base64.b64encode(email.encode('utf-8'))

            email_encrypted = email_byte.decode("utf-8")
            name_encrypted = urllib.parse.quote(row[0])
            # create encryption
            SHA1_string = (secret_key + email + str(row[2]))

            SHA1_encrypted = hashlib.sha1(str.encode(SHA1_string)).hexdigest()

            unique_link = ("https://se.trustpilot.com/evaluate/prodiga.se?a=" + str(row[2]) + "&b="
                           + email_encrypted + "&c=" + name_encrypted + "&e=" + SHA1_encrypted)

            sheet[new_column_letter + str(column_count + 1)] = unique_link
            column_count += 1
        else:
            print("Email value is empty. Skipping this row.")

    workbook.save(path)

    print("Done! Unique links written to the new column in 'test.xlsx'.")
    break
