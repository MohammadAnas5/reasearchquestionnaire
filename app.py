from flask import Flask, render_template, request, redirect
from openpyxl import Workbook, load_workbook

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():
    if request.method == 'POST':
        data = request.form.to_dict()
        add_data_to_excel(data)
        return "Thanks"

def add_data_to_excel(data):
    wb = load_workbook('data.xlsx')
    ws = wb.active

    new_row = [data['gender'], data['course'], data['q3'], data['q4'], data['q5'], data['q6'], data['q7'], data['q8'], data['q9'], data['q10'], data['q11'], data['q12'], data['q13'], data['q14'], data['q15'], data['q16'], data['q17']]  # Adjust this based on your form fields
    ws.append(new_row)
    wb.save('data.xlsx')

if __name__ == '__main__':
    app.run(host="0.0.0.0")
