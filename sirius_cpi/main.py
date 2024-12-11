from fastapi import FastAPI, File, UploadFile, Request
import pandas as pd
import numpy as np
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
import os

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

score_mapping = {
    "Не соответствует ожиданиям": 0,
    "Значительно ниже ожиданий": 50,
    "Ниже ожиданий": 75,
    "Частично ниже ожиданий": 90,
    "Соответствует ожиданиям": 100,
    "Выше ожиданий": 125,
    "Превосходит все ожидания": 150,
    "Не взаимодействовал(-а)": np.nan 
}

question_weights = {
    "Как вы оцениваете качество/достоверность предоставляемых результатов?": 25,
    "Как вы оцениваете время выполнения сервисных услуг?": 25,
    "Как вы оцениваете консультации по профильным (профессиональным) вопросам?": 10,
    "Как вы оцениваете качество взаимодействия (телефон, e-mail, лично)?": 10,
    "Как вы оцениваете понятность/полезность выдаваемой интерпретации или отчетов?": 10,
    "Как вы оцениваете объем оказываемых услуг (оснащенность)?": 5,
    "Какая в целом ваша оценка работы подразделения?": 15
}

@app.get("/", response_class=HTMLResponse)
async def upload_page(request: Request):
    return templates.TemplateResponse("upload.html", {"request": request})

@app.post("/process-file/")
async def process_file(file: UploadFile = File(...)):
    try:
        input_df = pd.read_excel(file.file)
    except Exception as e:
        return {"error": "Invalid file format or content", "details": str(e)}
    
    input_df.columns = input_df.columns.str.strip()
    available_columns = input_df.columns.tolist()
    print(f"Available columns: {available_columns}")

    question_columns = [col for col in input_df.columns if col in question_weights]
    print(f"Question columns: {question_columns}")

    if not question_columns:
        return {
            "error": "No valid question columns found in input data",
            "available_columns": available_columns
        }
    processed_rows = []
    for col in question_columns:
        temp_df = input_df[[col]].copy()
        temp_df.rename(columns={col: "оценка"}, inplace=True)
        temp_df["Вопрос"] = col
        temp_df["Вес"] = question_weights.get(col, 0)
        print(f"Processed data for {col}:")
        print(temp_df.head())
        temp_df["оценка"] = temp_df["оценка"].replace(score_mapping)
        temp_df["оценка"].fillna(0, inplace=True)
        processed_rows.append(temp_df)
    if not processed_rows:
        return {"error": "No valid data to process"}

    processed_df = pd.concat(processed_rows, ignore_index=True)
    print("Processed DataFrame before grouping:")
    print(processed_df.head())
    try:
        average_df = processed_df.groupby("Вопрос", as_index=False).agg({
            "оценка": "mean", 
            "Вес": "first"
        })
    except KeyError as e:
        return {"error": f"Missing required column in input data: {e}"}
    print("Grouped data:")
    print(average_df)

    average_df['оценка с учетом веса, %'] = (
        average_df['оценка'] * average_df['Вес'] / 100
    ).round(1)

    overall_score = average_df['оценка с учетом веса, %'].sum().round(1)
    summary_row = {
        "Вопрос": "Итоговая оценка",
        "оценка с учетом веса, %": overall_score,
        "оценка": np.nan,
        "Вес": np.nan
    }
    average_df = pd.concat([average_df, pd.DataFrame([summary_row])], ignore_index=True)
    print("Final DataFrame with overall score:")
    print(average_df)
    output_file = "processed_data.xlsx"
    average_df.to_excel(output_file, index=False)

    return FileResponse(output_file, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=output_file)
