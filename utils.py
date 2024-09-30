import json 
import mimetypes
import os
import requests
from urllib.parse import urlparse
from config import BASE_URL

# used by "/forecastingapicall2"
def call_api(token: str, 
             apiroute: str,  
             filename: str,          
             method: str = 'RidgeCV',
             n_hidden_features: int = 5,
             lags: int = 20,
             type_pi: str = 'gaussian',
             replications: int = 50,
             h: int = 10):        

    headers = {
    'Authorization': 'Bearer ' + token,
    }

    params = {'base_model': str(method), 
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

    try: 
        response = requests.post(BASE_URL + apiroute, 
            params=params, headers=headers, 
            files=files)

        res = response.json()

        # transform strings to lists
        mean = json.loads(res["mean"])
        lower = json.loads(res["lower"])
        upper = json.loads(res["upper"])
        try:
            sims = json.loads(res["sims"])
        except Exception:
            pass 
    except Exception as e:
        print("check validity of token")
        return {"error": str(e)}
    
    try:
        return mean, lower, upper, sims     
    except Exception as e:
        return mean, lower, upper    