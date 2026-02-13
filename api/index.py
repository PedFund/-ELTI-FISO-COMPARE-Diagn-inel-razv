from flask import Flask, request, send_file, jsonify, render_template_string
import pandas as pd
import io
import re
import xlsxwriter
from api.utils import process_dataframe, to_float # Импорт логики

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

HTML_TEMPLATE = '''<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Kidski Аналитика</title>
    <style>
        body { font-family: -apple-system, sans-serif; background: #f0f2f5; display: flex; align-items: center; justify-content: center; min-height: 100vh; margin: 0; }
        .card { background: white; padding: 2rem; border-radius: 16px; box-shadow: 0 10px 25px rgba(0,0,0,0.1); width: 100%; max-width: 600px; }
        
        .tabs { display: flex; margin-bottom: 2rem; border-bottom: 2px solid #eee; }
        .tab { flex: 1; text-align: center; padding: 1rem; cursor: pointer; color: #718096; font-weight: 600; transition: 0.3s; }
        .tab.active { color: #4299e1; border-bottom: 2px solid #4299e1; margin-bottom: -2px; }
        .tab:hover { background: #f7fafc; }
        
        .form-content { display: none; }
        .form-content.active { display: block; }
        
        h1 { font-size: 1.5rem; text-align: center; color: #2d3748; margin-bottom: 0.5rem; }
        p { text-align: center; color: #718096; margin-bottom: 2rem; }
        
        .upload-area { border: 2px dashed #cbd5e0; border-radius: 12px; padding: 2rem; text-align: center; cursor: pointer; transition: 0.2s; background: #f7fafc; margin-bottom: 1rem; }
        .upload-area:hover { border-color: #4299e1; background: #ebf8ff; }
        .icon { font-size: 2.5rem; display: block; margin-bottom: 0.5rem; }
        
        .btn { width: 100%; padding: 1rem; border: none; border-radius: 8px; font-weight: bold; font-size: 1rem; cursor: pointer; margin-top: 1rem; transition: 0.2s; }
        .btn-blue { background: #4299e1; color: white; }
        .btn-blue:hover { background: #3182ce; }
        .btn-green { background: #48bb78; color: white; }
        .btn-green:hover { background: #38a169; }
        
        #status { margin-top: 1rem; padding: 1rem; border-radius: 8px; display: none; text-align: center; }
        .error { background: #fff5f5; color: #c53030; }
        .success { background: #f0fff4; color: #2f855a; }
        
        .file-name { font-size: 0.9rem; color: #4a5568; margin-top: 0.5rem; font-weight: 500; }
    </style>
</head>
<body>
    <div class="card">
        <h1>📊 Kidski Аналитика</h1>
        
        <div class="tabs">
            <div class="tab active" onclick="switchTab('single')">Один файл (Excel)</div>
            <div class="tab" onclick="switchTab('compare')">Сравнение (Word)</div>
        </div>

        <div id="single" class="form-content active">
            <p>Генерация подробного отчета и графиков по одной диагностике</p>
            <form onsubmit="handleSingle(event)">
                <div class="upload-area" onclick="document.getElementById('fileSingle').click()">
                    <span class="icon">📄</span>
                    <span>Выберите Excel или CSV файл</span>
                    <div id="nameSingle" class="file-name"></div>
                    <input type="file" id="fileSingle" hidden>
                </div>
                <button type="submit" class="btn btn-blue">Обработать и Скачать</button>
            </form>
        </div>

        <div id="compare" class="form-content">
            <p>Сравнение результатов (Начало vs Конец) и генерация Word-отчета</p>
            <form onsubmit="handleCompare(event)">
                <div class="upload-area" onclick="document.getElementById('fileStart').click()">
                    <span class="icon">1️⃣</span>
                    <span>Загрузите НАЧАЛО года</span>
                    <div id="nameStart" class="file-name"></div>
                    <input type="file" id="fileStart" hidden>
                </div>
                <div class="upload-area" onclick="document.getElementById('fileEnd').click()">
                    <span class="icon">2️⃣</span>
                    <span>Загрузите КОНЕЦ года</span>
                    <div id="nameEnd" class="file-name"></div>
                    <input type="file" id="fileEnd" hidden>
                </div>
                <button type="submit" class="btn btn-green">Сравнить и Скачать Word</button>
            </form>
        </div>

        <div id="status"></div>
    </div>

    <script>
        function switchTab(id) {
            document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
            document.querySelectorAll('.form-content').forEach(c => c.classList.remove('active'));
            document.querySelector(`.tab[onclick="switchTab('${id}')"]`).classList.add('active');
            document.getElementById(id).classList.add('active');
            document.getElementById('status').style.display = 'none';
        }

        function showStatus(msg, type) {
            const el = document.getElementById('status');
            el.innerHTML = msg;
            el.className = type;
            el.style.display = 'block';
        }

        // Обновление имен файлов
        document.getElementById('fileSingle').onchange = e => document.getElementById('nameSingle').innerText = e.target.files[0]?.name || '';
        document.getElementById('fileStart').onchange = e => document.getElementById('nameStart').innerText = e.target.files[0]?.name || '';
        document.getElementById('fileEnd').onchange = e => document.getElementById('nameEnd').innerText = e.target.files[0]?.name || '';

        // Обработка ОДНОГО файла
        async function handleSingle(e) {
            e.preventDefault();
            const file = document.getElementById('fileSingle').files[0];
            if (!file) return showStatus("Выберите файл!", "error");
            
            showStatus("⏳ Обработка...", "");
            const fd = new FormData();
            fd.append('file', file);
            
            try {
                const res = await fetch('/api/process', { method: 'POST', body: fd });
                handleResponse(res);
            } catch(err) { showStatus(err.message, "error"); }
        }

        // Обработка СРАВНЕНИЯ
        async function handleCompare(e) {
            e.preventDefault();
            const f1 = document.getElementById('fileStart').files[0];
            const f2 = document.getElementById('fileEnd').files[0];
            if (!f1 || !f2) return showStatus("Выберите оба файла!", "error");

            showStatus("⏳ Анализ данных и генерация Word...", "");
            const fd = new FormData();
            fd.append('file_start', f1);
            fd.append('file_end', f2);

            try {
                const res = await fetch('/api/compare', { method: 'POST', body: fd });
                handleResponse(res);
            } catch(err) { showStatus(err.message, "error"); }
        }

        async function handleResponse(res) {
            if (res.ok) {
                const blob = await res.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = res.headers.get('X-Filename') || 'report.xlsx'; // Дефолтное имя
                a.click();
                showStatus("✅ Готово! Файл скачан.", "success");
            } else {
                const txt = await res.text();
                try {
                    const json = JSON.parse(txt);
                    showStatus("❌ " + json.error, "error");
                } catch { showStatus("❌ Ошибка сервера", "error"); }
            }
        }
    </script>
</body>
</html>'''

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/api/process', methods=['POST'])
def process():
    try:
        f = request.files.get('file')
        if not f: return jsonify({'error': 'Нет файла'}), 400
        
        # 1. Считаем метрики (через utils.py)
        df = process_dataframe(f)
        
        # 2. Формируем имя
        match = re.match(r'(\d+)-(\d+)', f.filename)
        name_part = f"{match.group(1)}-{match.group(2)}" if match else "Results"
        filename = f"{name_part}_results.xlsx"

        # 3. Генерируем Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            wb = writer.book
            fmt_head = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1, "align": "center"})
            fmt_pct = wb.add_format({"num_format": "0.0%", "align": "center"})
            
            # --- ЛИСТ 1: Возрасты ---
            age = df["Возраст"].value_counts().reset_index()
            age.columns = ["Группа", "Кол-во"]
            if not age.empty:
                age.to_excel(writer, sheet_name="Возрасты", index=False)
                ch = wb.add_chart({"type": "pie"})
                ch.add_series({
                    "categories": ["Возрасты", 1, 0, len(age), 0],
                    "values": ["Возрасты", 1, 1, len(age), 1],
                    "data_labels": {"percentage": True}
                })
                writer.sheets["Возрасты"].insert_chart("D2", ch)

            # --- ЛИСТЫ ПОКАЗАТЕЛЕЙ (с % на графиках) ---
            for metric in ["Когнитивное развитие", "Воображение_итог", "ЭмСоцИнтеллект"]:
                sheet_name = metric.replace(" ", "_")[:30]
                counts = df[f"{metric}_уровень"].value_counts().reindex(
                    ["ниже нормативного", "нормативный", "выше нормативного"], fill_value=0
                ).reset_index()
                counts.columns = ["Уровень", "Кол-во"]
                counts["Доля"] = counts["Кол-во"] / len(df)

                counts.to_excel(writer, sheet_name=sheet_name, index=False)
                ws = writer.sheets[sheet_name]
                ws.set_column(0, 0, 20)
                ws.set_column(2, 2, 10, fmt_pct)
                
                ch = wb.add_chart({"type": "column"})
                ch.add_series({
                    "name": metric,
                    "categories": [sheet_name, 1, 0, 3, 0],
                    "values": [sheet_name, 1, 2, 3, 2], # Колонка Доля
                    "data_labels": {"value": True, "num_format": "0.0%"}
                })
                ch.set_y_axis({"num_format": "0%"})
                ws.insert_chart("E2", ch)

            # --- ЛИСТ ИТОГ ---
            df.to_excel(writer, sheet_name="Полные_данные", index=False)

        output.seek(0)
        res = send_file(output, as_attachment=True, download_name=filename)
        res.headers['X-Filename'] = filename
        res.headers['Access-Control-Expose-Headers'] = 'X-Filename'
        return res

    except Exception as e:
        return jsonify({'error': str(e)}), 500

app = app
