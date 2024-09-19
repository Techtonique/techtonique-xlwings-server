import json 
import mimetypes
import os
import requests
import xlwings as xw

from urllib.parse import urlparse
from fastapi import Body, FastAPI, status, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import PlainTextResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from pathlib import Path

from config import BASE_URL

app = FastAPI()

# Mount the static files directory
app.mount("/static", StaticFiles(directory="static"), name="static")

@app.get("/")
async def read_index():
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

    headers = {
    'Authorization': 'Bearer ' + token,
    }

    params = {'method': method, 
    'n_hidden_features': str(n_hidden_features),
    'lags': str(lags),
    'type_pi': str(type_pi),
    'replications': str(replications),
    'h': str(h)}

    # Download the file from the URL
    response = requests.get(filename)
    
    if response.status_code != 200:
        return {"error": "Failed to download file"}

    # Extract the filename from the URL
    parsed_url = urlparse(filename)
    local_filename = os.path.basename(parsed_url.path)
    
    # Get the content type
    content_type, _ = mimetypes.guess_type(local_filename)

    # If content_type couldn't be guessed, default to 'application/octet-stream'
    if content_type is None:
        content_type = 'application/octet-stream'

    # Prepare the file for uploading (file content, filename, and content_type)
    file_content = response.content
    files = {
        'file': (local_filename, file_content, content_type)
    }

    response = requests.post(BASE_URL + apiroute, 
        params=params, headers=headers, 
        files=files)

    res = response.json()

    # transform strings to lists
    mean = json.loads(res["mean"])
    lower = json.loads(res["lower"])
    upper = json.loads(res["upper"])

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

