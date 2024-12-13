from fastapi import FastAPI, File, UploadFile, Request, Form
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
import pandas as pd
import numpy as np
import os
from pprint import pprint
from datetime import datetime, timedelta
from jose import JWTError, jwt
from passlib.context import CryptContext


app = FastAPI()

app.mount("/static", StaticFiles(directory="static"), name="static")

templates = Jinja2Templates(directory="templates")

department_question_weights = {
    "Контрактные услуги": {
        "Соответствует ли результат проведения работ/услуг по договору вашим ожиданиям?    ": 20,
        "Соответствует ли срок проведения работ/услуг по договору вашим ожиданиям?    ": 15,
        "Соответствует ли стоимость работ/услуг по договору вашим ожиданиям?    ": 15,
        "Как вы оцениваете качество административной поддержки по договору?    ": 15,
        "Соответствует ли срок согласования договора ожидаемому?    ": 15,
        "Соответствует ли уровень компетенций специалистов-исследователей ожидаемому?    ": 20,
    },
    "ДПО": {
        "Соответствует ли программа образовательного модуля вашим ожиданиям?": 25,
        "Соответствует ли уровень компетенций преподавателей программы ожидаемому?": 25,
        "Оцените программу курса по критериям     [Полезность информации]": 5,
        "Оцените программу курса по критериям     [Сложность программы]": 5,
        "Оцените программу курса по критериям     [Доступность изложения материала]": 5,
        "Оцените программу курса по критериям     [Соотношение теории и практики]": 5,
        "Оцените программу курса по критериям     [Полезность практических занятий]": 5,
        "Оцените программу курса по критериям     [Полнота и доступность ответов на вопросы аудитории]": 5,
        "Оцените организационное сопровождение программы     [Информационное сопровождение до начала и в течение программы организаторами]": 5,
        "Оцените организационное сопровождение программы     [Материально-техническое обеспечение]": 5,
        "Оцените организационное сопровождение программы     [Проживание]": 5,
        "Оцените организационное сопровождение программы     [Питание]": 5,
    },
    "ЕН": {
        "Как вы оцениваете время выполнения запросов?": 30,
        "Как вы оцениваете оснащенность предоставленной зоны?": 20,
        "Как вы оцениваете качество взаимодействия (телефон, e-mail, лично) и консультаций по работе?": 30,
        "Какая в целом ваша оценка работы подразделения?": 20,
    },
    "ИТО": {
        "Как вы оцениваете проведение пуско-наладочных работ?     [Качество]": 7.5,
        "Как вы оцениваете проведение пуско-наладочных работ?     [Сроки]": 7.5,
        "Как вы оцениваете проведение планового технического обслуживания оборудования?     [Качество]": 7.5,
        "Как вы оцениваете проведение планового технического обслуживания оборудования?     [Сроки]": 7.5,
        "Как вы оцениваете проведение ремонта оборудования?     [Качество]": 7.5,
        "Как вы оцениваете проведение ремонта оборудования?     [Сроки]": 7.5,
        "Как вы оцениваете метрологическое обеспечение оборудования (поверка, аттестация)?     [Качество]": 7.5,
        "Как вы оцениваете метрологическое обеспечение оборудования (поверка, аттестация)?     [Сроки]": 7.5,
        "Как вы оцениваете снабжение медицинскими газами и криожидкостями?  [Качество]": 5,
        "Как вы оцениваете снабжение медицинскими газами и криожидкостями?  [Сроки]": 5,
        'Как вы оцениваете информационную систему подачи заявок ServiceDesk "Сервис лабораторного оборудования" для обращений в инженерно-технический отдел? ': 10,
        "Какая в целом ваша оценка работы подразделения?": 20,
    },
    "Ресурсные центры": {
        "Как вы оцениваете качество/достоверность предоставляемых результатов?": 25,
        "Как вы оцениваете время выполнения сервисных услуг?": 25,
        "Как вы оцениваете консультации по профильным (профессиональным) вопросам?": 10,
        "Как вы оцениваете качество взаимодействия (телефон, e-mail, лично)?": 10,
        "Как вы оцениваете понятность/полезность выдаваемой интерпретации или отчетов?": 10,
        "Как вы оцениваете объем оказываемых услуг (оснащенность)?": 5,
        "Какая в целом ваша оценка работы подразделения?": 15,
    },
}

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


@app.get("/", response_class=HTMLResponse)
async def department_selection(request: Request):
    departments = list(department_question_weights.keys())
    return templates.TemplateResponse("department_selection.html", {"request": request, "departments": departments})

@app.get("/department/{department_name}/", response_class=HTMLResponse)
async def department_page(request: Request, department_name: str):
    return templates.TemplateResponse("upload.html", {"request": request, "department": department_name})

@app.post("/process-file/")
async def process_file(file: UploadFile = File(...), department: str = Form(...)):
    try:
        input_df = pd.read_excel(file.file)
    except Exception as e:
        return {"error": "Invalid file format or content", "details": str(e)}

    input_df.columns = input_df.columns.str.strip()

    if department not in department_question_weights:
        return {"error": f"Unknown department: {department}"}

    question_weights = department_question_weights[department]
    question_columns = [col for col in input_df.columns if col in question_weights]

        # New functionality: extract and return all questions for copying
    # questions_to_copy = input_df.columns.tolist()
    # print("Questions in the uploaded file:")
    # for question in questions_to_copy:
    #     print(question)

    if not question_columns:
        return {
            "error": "No valid question columns found in input data",
            "available_columns": input_df.columns.tolist()
        }

    processed_rows = []
    for col in question_columns:
        temp_df = input_df[[col]].copy()
        temp_df.rename(columns={col: "оценка"}, inplace=True)
        temp_df["Вопрос"] = col
        temp_df["Вес"] = question_weights.get(col, 0)
        temp_df["оценка"] = temp_df["оценка"].replace(score_mapping)
        temp_df["оценка"].fillna(0, inplace=True)
        processed_rows.append(temp_df)

    if not processed_rows:
        return {"error": "No valid data to process"}

    processed_df = pd.concat(processed_rows, ignore_index=True)
    average_df = processed_df.groupby("Вопрос", as_index=False).agg({
        "оценка": "mean",
        "Вес": "first"
    })
    average_df['оценка с учетом веса, %'] = (
        average_df['оценка'] * average_df['Вес'] / 100
    ).round(1)

    overall_score = average_df['оценка с учетом веса, %'].sum().round(1)
    summary_row = {
        "Вопрос": f"Итоговая оценка для отдела: {department}" if department else "Итоговая оценка",
        "оценка с учетом веса, %": overall_score,
        "оценка": np.nan,
        "Вес": np.nan
    }
    average_df = pd.concat([average_df, pd.DataFrame([summary_row])], ignore_index=True)
    output_file = f"{department}_processed_data.xlsx" if department else "processed_data.xlsx"
    average_df.to_excel(output_file, index=False)

    return FileResponse(output_file, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=output_file)
