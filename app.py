from flask import Flask, request, render_template, send_file
import pandas as pd
import io
import os

app = Flask(__name__, template_folder='templates')


def classify_layover(row):
    try:
        inbound = str(row['INBOUND']).strip()
        outbound = str(row['OUTBOUND']).strip()
        in_dep, in_arr = inbound.split('-')
        out_dep, out_arr = outbound.split('-')
    except Exception:
        return 'Unknown'

    if out_dep == out_arr:
        return 'Ramp Return'
    elif in_arr != out_dep:
        return 'Diversion'
    else:
        return 'Night Stop'


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files['file']
        if not file:
            return "No file uploaded.", 400

        df = pd.read_excel(file, sheet_name=0)

        # Clean column names
        df.columns = df.columns.str.replace('\n', ' ').str.strip()
        df.rename(columns={
            'CITY PAIR IN BOUND': 'INBOUND',
            'CITY PAIR OUT BOUND': 'OUTBOUND'
        }, inplace=True)

        df['LAYOVER_TYPE'] = df.apply(classify_layover, axis=1)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Classified')

        output.seek(0)
        return send_file(output, download_name="classified_layover.xlsx", as_attachment=True)

    return render_template('upload.html')


if __name__ == '__main__':
app.run(host="0.0.0.0", port=10000)
