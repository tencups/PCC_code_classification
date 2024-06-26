from flask import Flask, render_template
import pandas as pd
import re

app = Flask(__name__)

def read_excel_and_print_column(file_path):
    # Read the Excel file
    df = pd.read_excel(file_path)

    # Check if the column "Abstract Text (only)" exists in the dataframe
    if 'Abstract Text (only)' in df.columns:
        for index, row in df.iterrows():
            #print(row['Abstract Text (only)'])
            print(f"row {index}")
    else:
        print("The column 'Abstract Text (only)' does not exist in the provided Excel file.")

@app.route('/')
def home():
    # Call the function and provide the path to your Excel file
    read_excel_and_print_column('ATSB all grants-app FY2021-2024.xlsx')
    return render_template('index.html')

def assign_pcc_code(text):
    keyword_to_code = {
        'data science': 'DS',
        'genomic': 'GP',
        'imaging technology': 'IT',
        'nanotechnology': 'NT'
    }
    for keyword, code in keyword_to_code.items():
        if re.search(r'\b{}\b'.format(re.escape(keyword)), text, re.IGNORECASE):
            return code
    return None  # return none if no matching keyword found


if __name__ == "__main__":
    app.run(debug=True)
