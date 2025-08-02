from flask import Flask, render_template, request, redirect, send_file
import os
import pandas as pd
from process_excel import process_file

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files['file']
        file_type = request.form['type']
        if file:
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(filepath)

            revenue, cost, profit = process_file(filepath, file_type)

            return render_template('report.html',
                                   revenue=revenue.to_dict(orient='records'),
                                   cost=cost.to_dict(orient='records'),
                                   profit=profit.to_dict(orient='records'),
                                   columns={
                                       'revenue': revenue.columns,
                                       'cost': cost.columns,
                                       'profit': profit.columns
                                   },
                                   file_path=filepath,
                                   file_type=file_type)
    return render_template('index.html')


@app.route('/download', methods=['POST'])
def download():
    file_type = request.form['file_type']
    file_path = request.form['file_path']

    revenue, cost, profit = process_file(file_path, file_type)

    # Save to a new Excel file
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'report_output.xlsx')
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        revenue.to_excel(writer, sheet_name='Revenue', index=False)
        cost.to_excel(writer, sheet_name='Cost', index=False)
        profit.to_excel(writer, sheet_name='Profit', index=False)

    return send_file(output_path, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True)
