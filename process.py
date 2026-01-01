import requests
import os
from dotenv import load_dotenv

load_dotenv()

def get_access_token():
    url = f"{os.getenv('ZOHO_ACCOUNTS_URL')}/oauth/v2/token"

    params = {
        "refresh_token": os.getenv("ZOHO_REFRESH_TOKEN"),
        "client_id": os.getenv("ZOHO_CLIENT_ID"),
        "client_secret": os.getenv("ZOHO_CLIENT_SECRET"),
        "grant_type": "refresh_token"
    }

    r = requests.post(url, params=params)
    r.raise_for_status()

    return r.json()["access_token"]
