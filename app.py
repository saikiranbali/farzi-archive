import os
import json
import requests
import pandas as pd
from flask import Flask, render_template, request, jsonify
from io import BytesIO 
import xlrd
from flask_cors import CORS

app = Flask(__name__)

def lines_to_string(lines):
    line_strs = []
    for line in lines:
        line_str = "{"
        line_str += f"quantity: {line['quantity']}, "
        line_str += f"variantName: \"{line['variantName']}\", "
        line_str += f"productName: \"{line['productName']}\", "
        line_str += f"productSku: \"{line['productSku']}\", "
        line_str += f"productSku: \"{line['productSku']}\", "
        line_str += f"unitPriceGrossAmount: \"{line['unitPriceGrossAmount']}\", "
        line_str += f"unitPriceNetAmount: \"{line['unitPriceNetAmount']}\", "
        line_str += f"imageUrl: \"{line['imageUrl']}\""
        line_str += "}"
        line_strs.append(line_str)
    return "[" + ", ".join(line_strs) + "]"

def process_excel(df):
    # df = pd.read_excel(file)
    grouped_data = df.groupby('foreignOrderId')
    results = []

    for foreignOrderId, group in grouped_data:
        # calculate the sum of totalGrossAmount and totalNetAmount for this foreignOrderId
        gross_sum = group['unitPriceNetAmount'].sum()
        net_sum = group['totalNetAmount'].sum()

    for foreignOrderId, group in grouped_data:
        lines = []
        for _, row in group.iterrows():
            lines.append({
        'quantity': int(row['quantity']),
        'variantName': row['variantName'],
        'productName': row['productName'],
        'productSku': row['productSku'],
        'unitPriceGrossAmount': row['unitPriceGrossAmount'],
        'unitPriceNetAmount': row['unitPriceNetAmount'],
        'imageUrl': ""
    })

        # ... (same as before)


        row = group.iloc[0]
        payload = {
            'query': f'''mutation {{
                archiveOrderCreate(input: {{
                    foreignOrderId: "{foreignOrderId}",
                    email: "{row['email']}",
                    lines: {lines_to_string(lines)},
                    placedOn: "{row['placed_on']}",
                    totalGrossAmount: "{gross_sum}",
                    totalNetAmount: "{net_sum}",
                    discountAmount: "{row['discountAmount']}",
                    metadata: "{{ shopify_order_id: 4525580779564 }}",
                    privateMetadata: "{row['privateMetadata']}",
                    shippingAddress: {{
                        firstName: "{row['sa_firstName']}",
                        lastName: "{row['sa_lastName']}",
                        address1: "{row['sa_address1']}",
                        address2: "{row['sa_address2']}",
                        city: "{row['sa_city']}",
                        zip: "{row['sa_zip']}",
                        province: "{row['sa_province']}",
                        phone: "{row['sa_phone']}",
                        countryCode: IN
                    }},
                    billingAddress: {{
                        firstName: "{row['ba_firstName']}",
                        lastName: "{row['ba_lastName']}",
                        address1: "{row['ba_address1']}",
                        address2: "{row['ba_address2']}",
                        city: "{row['ba_city']}",
                        zip: "{row['ba_zip']}",
                        province: "{row['ba_province']}",
                        phone: "{row['ba_phone']}",
                        countryCode: IN
                    }}
                }}) {{
                    archiveOrder {{
                        id,
                        created,
                        userEmail
                    }},
                    archiveOrderErrors {{
                        message,
                        field,
                        code
                    }}
                }}
            }}'''
        }

        url = 'https://gourmetgardenhapi.farziengineer.co/graphql/?source=website'
        headers = {'Authorization': 'Bearer BX7iC0evrcwbp1RaFSLn5lnNTYWmSS', 'Content-Type': 'application/json'}
        response = requests.post(url, headers=headers, data=json.dumps(payload))
        results.append(response.text)

    print(payload)

    return results

# @app.route('/', methods=['GET', 'POST'])
# def index():

#     if request.method == 'POST':
#         try:
#             file = request.files['file']
#             file_buffer = BytesIO(file.read())
#             df = pd.read_excel(file_buffer)
#             results = process_excel(df)
#             return jsonify(results)
#         except Exception as e:
#             return jsonify({'error': str(e)}), 400
#     # if request.method == 'POST':
#     #     file = request.files['file']
#     #     results = process_excel(file)
#     #     return jsonify(results)

#     return render_template('index.html')

@app.route('/', methods=['GET', 'POST'])
def index():

    if request.method == 'POST':
        try:
            file = request.files['file']
            file_buffer = BytesIO(file.read())
            workbook = xlrd.open_workbook(file_contents=file_buffer.read())
            sheet = workbook.sheet_by_index(0)
            headers = [sheet.cell_value(0, col_index) for col_index in range(sheet.ncols)]
            data = []
            for row_index in range(1, sheet.nrows):
                row = {}
                for col_index in range(sheet.ncols):
                    row[headers[col_index]] = sheet.cell_value(row_index, col_index)
                data.append(row)
            df = pd.DataFrame(data)
            results = process_excel(df)
            return jsonify(results)
        except Exception as e:
            return jsonify({'error': str(e)}), 400

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)