"""
Excel + Python Automatisierung
Projekt: Interaktives Dashboard mit Flask (Demo 4)
Erstellt: August 2026
"""

from flask import Flask, render_template, request
import pandas as pd
import plotly.express as px
import plotly.utils
import json
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            return render_template('index.html', error='Keine Datei ausgewählt.')
        file = request.files['file']
        if file.filename == '':
            return render_template('index.html', error='Keine Datei ausgewählt.')
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            try:
                df = pd.read_excel(filepath)
                table_html = df.to_html(classes='table table-striped', index=False)
                graphs = []
                numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
                if len(numeric_cols) > 0:
                    fig = px.bar(df, x=df.columns[0], y=numeric_cols[0], 
                                 title=f'{numeric_cols[0]} nach {df.columns[0]}')
                    graphs.append(json.dumps(fig, cls=plotly.utils.PlotlyJSONEncoder))
                if len(numeric_cols) > 1:
                    fig = px.line(df, x=df.columns[0], y=numeric_cols,
                                  title='Entwicklung der Kennzahlen')
                    graphs.append(json.dumps(fig, cls=plotly.utils.PlotlyJSONEncoder))
                if len(numeric_cols) > 0:
                    fig = px.pie(df, values=numeric_cols[0], names=df.columns[0],
                                 title=f'Verteilung {numeric_cols[0]}')
                    graphs.append(json.dumps(fig, cls=plotly.utils.PlotlyJSONEncoder))
                return render_template('index.html', table=table_html, graphs=graphs, filename=filename)
            except Exception as e:
                return render_template('index.html', error=f'Fehler: {str(e)}')
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
