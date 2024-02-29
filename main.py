from fastapi import FastAPI, Body, Response, Cookie
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from starlette.responses import RedirectResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles

from openpyxl import load_workbook
from openpyxl import Workbook

app = FastAPI()

#app.mount("/static", StaticFiles(directory="public", html=True))
#app.mount("/static", StaticFiles(directory="/", html=True))

link = "https://66b7-176-15-251-130.ngrok-free.app/"

origins = [
    link#,
    #"https://127.0.0.1:8000",
    #"https://localhost:8000"
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["GET", "POST"],#, "OPTIONS", "DELETE", "PATH", "PUT"],
    allow_headers=["Content-Type", "Accept", "Location", "Allow", "Content-Disposition", "Sec-Fetch-Dest"],
)



app.mount("/static", StaticFiles(directory="public", html=True))
app.mount("/static", StaticFiles(directory="/", html=True))



@app.get("/", response_class=HTMLResponse)
def root():
    return RedirectResponse(f"{link}static/menu.html")



@app.get("/menu")
def menu():
    wb = load_workbook('list.xlsx')
    sheet = wb['all']

    j=0
    HEADERS = {j : "ОГЛАВЛЕНИЕ"}
    for i in range(2, sheet.max_row+1):
        header = sheet[f"A{i}"].value
        if header not in HEADERS.values():
            j+=1
            HEADERS[j] = header
    
    return HEADERS



@app.get("/menu_plus")
def menu_plus():
    wb = load_workbook('list.xlsx')
    sheet = wb['all']

    j=0
    HEADERS = []
    for i in range(2, sheet.max_row+1):
        header = sheet[f"A{i}"].value
        if header not in HEADERS:
            j+=1
            HEADERS.append(header)

    MENU = dict()
    for header in HEADERS:
        MENU[header] = []
        for i in range(2, sheet.max_row+1):
            hdr = sheet[f"A{i}"].value
            if (hdr != None) and (hdr == header):# or (hdr.lower() in header.lower()):
                card = {
                    "id" : i,
                    "name" : sheet[f"A{i}"].value,
                    "text" : sheet[f"B{i}"].value,
                    "img" : sheet[f"C{i}"].value
                }
                MENU[header].append(card)
    
    return MENU

    

@app.get("/search_by_name/{name}")
def byname(name):
    name = str(name)
    wb = load_workbook('list.xlsx')
    sheet = wb['all']

    CARDS = []
    for i in range(2, sheet.max_row+1):
        nm = sheet[f"A{i}"].value
        if (nm != None) and ((nm == name) or (name.lower() in nm.lower())):
            card = {
                "id" : i,
                "name" : sheet[f"A{i}"].value,
                "text" : sheet[f"B{i}"].value,
                "img" : sheet[f"C{i}"].value
            }
            CARDS.append(card)
    
    return CARDS

@app.get("/search_by_text/{text}")
def bytext(text):
    text = str(text)
    wb = load_workbook('list.xlsx')
    sheet = wb['all']

    CARDS = []
    for i in range(2, sheet.max_row+1):
        nm = sheet[f"A{i}"].value
        txt = sheet[f"B{i}"].value
        if (txt != None) and ((sheet[f"A{i}"].value == text) or (text.lower() in nm.lower()) or (text.lower() in txt.lower())):
            card = {
                "id" : i,
                "name" : sheet[f"A{i}"].value,
                "text" : sheet[f"B{i}"].value,
                "img" : sheet[f"C{i}"].value
            }
            CARDS.append(card)
    
    return CARDS

@app.get("/application")
def application():
    text = "ПРИЛОЖЕНИЕ"
    wb = load_workbook('list.xlsx')
    sheet = wb['all']

    CARDS = []
    for i in range(2, sheet.max_row+1):
        nm = sheet[f"A{i}"].value
        if (nm != None) and ((sheet[f"A{i}"].value == text) or (text.lower() in nm.lower())):
            card = {
                "id" : i,
                "name" : sheet[f"A{i}"].value,
                "text" : sheet[f"B{i}"].value,
            }
            CARDS.append(card)
    
    return CARDS

#в конце
@app.get("img/{file}")
def files(file: str):
    #/imgs/1_1.jpg
    file = "/" + file
    return FileResponse(file) 