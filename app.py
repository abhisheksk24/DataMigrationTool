from flask import Flask, render_template, request, send_file
from simple_salesforce import Salesforce
import pandas as pd
import os

app = Flask(__name__)

sf_connection = None

@app.route('/')
def home():
    return render_template('home.html')

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        security_token = request.form['token']
        
        try:
            global sf_connection
            sf_connection = Salesforce(username=username, password=password, security_token=security_token)
            all_objects = sf_connection.describe()['sobjects']
            return render_template('select_objects.html', all_objects=all_objects)
        except Exception as e:
            return f"Failed to authenticate: {str(e)}"
    else:
        return render_template('index.html')

@app.route('/export', methods=['POST'])
def export_objects():
    selected_objects = request.form.getlist('objects')
    text_to_remove = request.form.get('text_remove')
    column_name = request.form.get('column_name', '2GP Fields') 

    uploaded_file = request.files.get('file')
    if uploaded_file and uploaded_file.filename != '':
        uploaded_excel = pd.ExcelFile(uploaded_file)
    else:
        uploaded_excel = None

    objects_df = pd.DataFrame(selected_objects, columns=['Object Name'])
    excel_file = 'salesforce_objects.xlsx'
    with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
        objects_df.to_excel(writer, sheet_name='Objects List', index=False)

        for obj_name in selected_objects:
            fields = getattr(sf_connection, obj_name).describe()['fields']
            fields_df = pd.DataFrame(fields)

            headers_df = pd.DataFrame(columns=['S#', '1GP Fields', '2GP Fields' , f'Object Name ({obj_name})'])

            fields_data = pd.DataFrame({
                'S#': range(1, len(fields_df) + 1),
                '1GP Fields': fields_df['name'],
                '2GP Fields': [None] * len(fields_df),
                f'Object Name ({obj_name})': [None] * len(fields_df)
            })

            if text_to_remove:
                fields_data['1GP Fields'] = fields_data['1GP Fields'].str.replace(text_to_remove, '')

            final_df = pd.concat([headers_df, fields_data], ignore_index=True)

            if uploaded_excel and obj_name in uploaded_excel.sheet_names:
                uploaded_df = pd.read_excel(uploaded_excel, sheet_name=obj_name)
                df22gp = set(uploaded_df.loc[:, column_name])
                df11gp = set(final_df.loc[:, "1GP Fields"])
                df2e = df22gp.difference(df11gp)
                df1e = df11gp.difference(df22gp)

                if df2e == df1e:
                    print("Same")
                else:
                    for i in range(len(final_df)):
                        if final_df.at[i, '1GP Fields'] in df1e:
                            pass
                        else:
                            final_df.loc[i, ['2GP Fields']] = [final_df.at[i, '1GP Fields']]
                    for i in df2e:
                        new_row = {'S#': None, '1GP Fields': None, '2GP Fields': i}
                        final_df = pd.concat([final_df, pd.DataFrame([new_row])], ignore_index=True)

            sheet_name = obj_name[:31] if len(obj_name) > 31 else obj_name
            final_df.to_excel(writer, sheet_name=sheet_name, index=False)

    return send_file(excel_file, as_attachment=True, download_name='salesforce_objects.xlsx')

@app.route('/map_fields', methods=['GET', 'POST'])
def map_fields():
    if request.method == 'POST':
        file1 = request.files.get('file1')
        file2 = request.files.get('file2')
        row1_name = request.form.get('row1')
        row2_name = request.form.get('row2')

        if file1 and file2:
            excel_file1 = pd.ExcelFile(file1)
            excel_file2 = pd.ExcelFile(file2)

            output_file = 'updated_file.xlsx'
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                for sheet_name in excel_file1.sheet_names:
                    df1 = pd.read_excel(excel_file1, sheet_name=sheet_name)

                    if sheet_name in excel_file2.sheet_names:
                        df2 = pd.read_excel(excel_file2, sheet_name=sheet_name)
                        
                        try:
                            df22gp = set(df2.loc[:, row2_name].dropna())
                        except KeyError:
                            print(f"Column name '{row2_name}' not found in sheet '{sheet_name}' of file 2. Skipping to the next sheet.")
                            df1.to_excel(writer, sheet_name=sheet_name, index=False)
                            continue

                        df11gp = set(df1.loc[:, row1_name].dropna())
                        df2e = df22gp.difference(df11gp)
                        df1e = df11gp.difference(df22gp)

                        if df2e != df1e:
                            for i in df1.index:
                                if df1.at[i, row1_name] in df1e:
                                    continue
                                else:
                                    df1.at[i, row2_name] = df1.at[i, row1_name]

                            for item in df2e:
                                new_row = {row1_name: None, row2_name: item}
                                df1 = df1._append(new_row, ignore_index=True)

                    df1.to_excel(writer, sheet_name=sheet_name, index=False)

            return send_file(output_file, as_attachment=True, download_name='updated_file.xlsx')
        else:
            return "Please upload both files."
    else:
        return render_template('map_fields.html')
    
if __name__ == '__main__':
    app.run(debug=True)
