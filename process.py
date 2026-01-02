import pandas as pd
import requests
from auth import get_access_token

DESK_BASE = "https://desk.zoho.in/api/v1"

ACCESS_TOKEN = get_access_token()

HEADERS = {
    "Authorization": f"Zoho-oauthtoken {ACCESS_TOKEN}",
    "Content-Type": "application/json"
}

def search_account(name):
    url = f"{DESK_BASE}/accounts/search"
    r = requests.get(url, headers=HEADERS, params={"accountName": name})

    if r.status_code == 200 and r.text:
        data = r.json()
        if "data" in data and data["data"]:
            return data["data"][0]["id"]

    return None


def create_account(row):
    url = f"{DESK_BASE}/accounts"
    payload = {
        "accountName": row["TestAccountName"],
        "state": row["TestState"],
        "city": row["TestCity"]
    }

    r = requests.post(url, headers=HEADERS, json=payload)
    if r.status_code in (200, 201):
        return r.json()["id"]

    raise Exception(f"Account create failed: {r.text}")


def update_account(account_id, row):
    url = f"{DESK_BASE}/accounts/{account_id}"
    payload = {
        "state": row["TestState"],
        "city": row["TestCity"]
    }
    r = requests.put(url, headers=HEADERS, json=payload)
    if r.status_code not in (200, 204):
        raise Exception(f"Account update failed: {r.text}")


def search_contact(email):
    url = f"{DESK_BASE}/contacts/search"
    r = requests.get(url, headers=HEADERS, params={"email": email})

    if r.status_code == 200 and r.text:
        data = r.json()
        if "data" in data and data["data"]:
            return data["data"][0]["id"]

    return None


def create_contact(row, account_id):
    url = f"{DESK_BASE}/contacts"
    payload = {
        "firstName": row["TestFirstName"],
        "lastName": row["TestLastName"],
        "email": row["TestEmail"],
        "phone": row["TestPhone"],
        "cf_lab_id": row["LabId"],
        "accountId": account_id,
        "ownerId": row["ContactOwnerId"]
    }

    r = requests.post(url, headers=HEADERS, json=payload)
    if r.status_code in (200, 201):
        return r.json()["id"]

    raise Exception(f"Contact create failed: {r.text}")


def update_contact(contact_id, row, account_id):
    url = f"{DESK_BASE}/contacts/{contact_id}"
    payload = {
        "phone": row["TestPhone"],
        "cf_lab_id": row["LabId"],
        "accountId": account_id,
        "ownerId": row["ContactOwnerId"]
    }
    r = requests.put(url, headers=HEADERS, json=payload)
    if r.status_code not in (200, 204):
        raise Exception(f"Contact update failed: {r.text}")


input_df = pd.read_excel("input.xlsx", sheet_name="Input_Data", engine="openpyxl")

results = []

for _, row in input_df.iterrows():
    try:
        account_id = search_account(row["TestAccountName"])

        if account_id:
            update_account(account_id, row)
        else:
            account_id = create_account(row)

        contact_id = search_contact(row["TestEmail"])

        if contact_id:
            update_contact(contact_id, row, account_id)
            status = "CONTACT_UPDATED"
        else:
            contact_id = create_contact(row, account_id)
            status = "CONTACT_CREATED"

        results.append([row["RowId"], account_id, contact_id, status, "SUCCESS"])

    except Exception as e:
        results.append([row["RowId"], "", "", "FAILED", str(e)])

result_df = pd.DataFrame(
    results,
    columns=["RowId", "AccountId", "ContactId", "Status", "Message"]
)

with pd.ExcelWriter("input.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
    result_df.to_excel(w, sheet_name="Processing_Result", index=False)
