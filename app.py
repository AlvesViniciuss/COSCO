from flask import Flask, request, render_template, send_file
import pandas as pd
import os
from io import BytesIO

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads/'

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files['file']
        if file and file.filename.endswith('.xlsx'):
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(file_path)
            result = process_file(file_path)
            return render_template('resultado.html', tables=[result.to_html(index=False)], titles=['Quantidades Diárias'])
    return render_template('index.html')

@app.route('/export', methods=['POST'])
def export():
    # Receber o DataFrame do resultado (como um exemplo simples, você pode querer usar um método mais robusto)
    df = pd.read_excel('uploads/result.xlsx')

    # Criar um buffer em memória
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Quantidades Diárias')

    # Voltar o buffer para o início
    output.seek(0)

    # Enviar o arquivo para download
    return send_file(output, as_attachment=True, download_name='quantidades_diarias.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

def process_file(file_path):
    df = pd.read_excel(file_path)

    # Verificar e imprimir os nomes das colunas para depuração
    print("Colunas disponíveis:", df.columns)

    # Atualizar os nomes das colunas
    df['Unit Time In'] = pd.to_datetime(df.get('Unit Time In', pd.NaT))
    df['Unit Time Out'] = pd.to_datetime(df.get('Unit Time Out', pd.NaT))

    # Adicionar a regra para 'Unit Type Length'
    def get_quantity(length):
        if length == "20'":
            return 1
        else:
            return 2

    df['Quantity'] = df['Unit Type Length'].apply(get_quantity)

    start_date = df['Unit Time In'].min()
    end_date = df['Unit Time Out'].max()
    all_dates = pd.date_range(start=start_date, end=end_date)

    date_counts = pd.DataFrame({'DATA': all_dates})
    date_counts['QUANTIDADE'] = 0

    for index, row in df.iterrows():
        mask = (date_counts['DATA'] >= row['Unit Time In']) & (date_counts['DATA'] <= row['Unit Time Out'])
        date_counts.loc[mask, 'QUANTIDADE'] += row['Quantity']

    date_counts.reset_index(drop=True, inplace=True)
    date_counts['DATA'] = date_counts['DATA'].dt.strftime('%d/%m/%Y')

    # Salvar o resultado em um arquivo Excel
    result_path = 'uploads/result.xlsx'
    date_counts.to_excel(result_path, index=False)

    return date_counts

if __name__ == '__main__':
    app.run(debug=True)
