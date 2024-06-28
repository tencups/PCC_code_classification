from flask import Flask, render_template
import pandas as pd
import re

app = Flask(__name__)

def read_excel_and_print_column(file_path):
    # Read the Excel file
    df = pd.read_excel(file_path)

    abstract_found = []

    # Check if the column "Abstract Text (only)" exists in the dataframe
    if 'Abstract Text (only)' in df.columns:
        for index, row in df.iterrows():
            abstract_found.append(row['Abstract Text (only)'])
            # print(row['Abstract Text (only)'])
    else:
        print("The column 'Abstract Text (only)' does not exist in the provided Excel file.")

def specific_aim(df):
    # check if  "Abstract Text (only)" column exists
    if 'Abstract Text (only)' in df.columns:
        found_aims = []

        # iterate through each row in the df
        for index, row in df.iterrows():
            abstract_text = row['Abstract Text (only)']
            if isinstance(abstract_text, str):
                # split abstract_text into sentences regex 
                sentences = re.split(r'(?<!\w\.\w.)(?<![A-Z][a-z]\.)(?<=\.|\?)\s', abstract_text)
                
                # iterate through each sentence to find occurrences of "aim" 
                for sentence in sentences:
                    if re.search(r'\baim\b', sentence, flags=re.IGNORECASE):
                        found_aims.append(sentence.strip())  # add sentence containing "aim"
            else:
                found_aims.append(None)  # add None for non-string abstract_text

        # edge case
        while len(found_aims) < len(df):
            found_aims.append(None)

        # add found aims as a new column 'Specific Aims'
        df['Specific Aims'] = found_aims

        return df
    else:
        print("The column 'Abstract Text (only)' does not exist in the provided DataFrame.")
        return None

@app.route('/')
def home():
    # Call the function and provide the path to your Excel file
    file_path = 'ATSB all grants-app FY2021-2024.xlsx'
    
    # read abstract texts from Excel
    abstract_texts = read_excel_and_print_column(file_path)

    # read Excel file into data frame
    df = pd.read_excel(file_path)

    # process dataframe to extract specific aims
    df_processed = specific_aim(df)

    if df_processed is not None:
        found_aims = []
        for index, row in df_processed.iterrows():
            found_aims.append(row['Specific Aims'])
            # print(f"Found aims for row {index}: {row['Specific Aims']}")

        return render_template('index.html', data=df_processed.to_html(), found_aims=found_aims)
    else:
        return "Error processing the Excel file."
    
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
