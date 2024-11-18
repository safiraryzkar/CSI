from flask import Flask, render_template, request, redirect, url_for, send_file
import pandas as pd
from io import BytesIO

app = Flask(__name__)

def calculate_csi(df, name=None):
    # Ambil kolom yang berisi nilai CSI (asumsikan kolom 4-15)
    csi_columns = df.columns[4:15]
    
    # Hitung total nilai dan CSI per Nama yang dinilai
    df['Total Nilai'] = df[csi_columns].sum(axis=1)
    df['CSI (%)'] = (df['Total Nilai'] / (len(csi_columns) * 5)) * 100  # Skala 5
    
    if name:
        # Filter data berdasarkan nama yang dinilai
        df = df[df['Nama yang dinilai'] == name]
        if df.empty:
            return None, f"Tidak ditemukan data untuk nama: {name}"
    
    # Kelompokkan hasil per Nama yang dinilai
    result = df.groupby('Nama yang dinilai').agg({
        'Total Nilai': 'sum',
        'CSI (%)': 'mean'
    }).reset_index()
    return result, None

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return redirect('/')
    
    file = request.files['file']
    if file.filename == '':
        return redirect('/')
    
    # Baca file Excel
    try:
        df = pd.read_excel(file)
    except Exception as e:
        return f"Error membaca file: {e}"
    
    # Ambil nama dari input pengguna
    name = request.form.get('name')
    result, error = calculate_csi(df, name)
    
    if error:
        return render_template('report.html', error=error)
    
    # Simpan hasil ke dalam buffer
    output = BytesIO()
    result.to_excel(output, index=False)
    output.seek(0)
    
    return render_template('report.html', tables=result.to_html(classes='table table-striped'), name=name)

if __name__ == '__main__':
    app.run(debug=True)
