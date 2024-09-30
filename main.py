import os
import requests
import xlwings as xw

from urllib.parse import urlparse
from fastapi import Body, FastAPI, status, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import PlainTextResponse, FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pathlib import Path
from dotenv import load_dotenv
from config import BASE_URL
from pydantic import BaseModel
import openpyxl
from openpyxl.utils import get_column_letter
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
import openpyxl
from openpyxl.utils import get_column_letter

from utils import call_api


# Load environment variables from .env file
load_dotenv()

app = FastAPI()

# Mount the static files directory
app.mount("/static", StaticFiles(directory="static"), name="static")

@app.get("/")
async def read_index():
    xlwings_license = os.getenv('XLWINGS_LICENSE_KEY')
    return FileResponse("static/index.html")

@app.post("/forecastingapicall")
def api_call(token: str, 
             apiroute: str,  
             filename: str,          
             data: dict = Body,
             method: str = 'RidgeCV',
             n_hidden_features: int = 5,
             lags: int = 20,
             type_pi: str = 'gaussian',
             replications: int = 10,
             h: int = 10):        

    mean, lower, upper = call_api(token, apiroute, filename, 
                                  method, n_hidden_features, 
                                  lags, type_pi, 
                                  replications, 
                                  h)

    # Instantiate a Book object with the deserialized request body
    with xw.Book(json=data) as book:
        # Use xlwings as usual
        sheet = book.sheets[0]
        # Write the lists to Excel columns
        sheet.range("A1").value = "Mean"
        sheet.range("B1").value = "Lower"
        sheet.range("C1").value = "Upper"

        # Write the lists to Excel columns, one value per row
        for i in range(len(mean)):
            idx = i + 2            
            sheet.range(f"A{idx}").value = mean[i]   # Write the 'mean' list in column A
            sheet.range(f"B{idx}").value = lower[i]  # Write the 'lower' list in column B
            sheet.range(f"C{idx}").value = upper[i]  # Write the 'upper' list in column C
        # Pass the following back as the response
        return book.json()

@app.exception_handler(Exception)
async def exception_handler(request, exception):
    # This handles all exceptions, so you may want to make this more restrictive
    return PlainTextResponse(
        str(exception), status_code=status.HTTP_500_INTERNAL_SERVER_ERROR
    )

@app.post("/forecastingapicall2")
def api_call2(token: str, 
              apiroute: str,  
              filename: str,                       
              method: str = 'RidgeCV',
              n_hidden_features: int = 5,
              lags: int = 20,
              type_pi: str = 'gaussian',
              replications: int = 50,              
              h: int = 10,):        
    
    try: 

        mean, lower, upper, sims = call_api(token, apiroute, filename, 
                                    method, n_hidden_features, 
                                    lags, type_pi, 
                                    replications, 
                                    h)
            
        return JSONResponse(content={"mean": mean, "lower": lower, 
                                     "upper": upper, "sims": sims})
    
    except Exception: 

        mean, lower, upper = call_api(token, apiroute, filename, 
                                    method, n_hidden_features, 
                                    lags, type_pi, 
                                    replications, 
                                    h)
            
        return JSONResponse(content={"mean": mean, 
                                     "lower": lower, 
                                     "upper": upper})

# Office Scripts and custom functions in Excel on the web require CORS
cors_app = CORSMiddleware(
    app=app,
    allow_origins="*",
    allow_methods=["POST"],
    allow_headers=["*"],
    allow_credentials=True,
)

if __name__ == "__main__":
    import uvicorn

    # Check if running on Heroku by checking for the existence of the 'PORT' environment variable
    port = int(os.environ.get("PORT", 8001))

    is_heroku = "PORT" in os.environ

    if is_heroku:
        # Running on Heroku
        uvicorn.run("main:cors_app", host="0.0.0.0", port=port, reload=True)
    else:
        # Running locally
        uvicorn.run("main:cors_app", host="0.0.0.0", port=8001, reload=True)












