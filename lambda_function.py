import json
import os
import xlwings as xw


def lambda_handler(event, context):
    if event["headers"]["authorization"] != os.environ["XLWINGS_API_KEY"]:
        return {
            "statusCode": 401,
            "headers": {
                "Content-Type": "text/html",
            },
            "body": "Unauthorized",
        }
    with xw.Book(json=json.loads(event["body"])) as book:

        sheet1 = book.sheets[0]
        df = sheet1["A1"].expand().options("df", index=False).value
        sheet1["G1"].value = df.T

        return book.json()
