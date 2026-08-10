from msal import ConfidentialClientApplication
import requests
import os
from dotenv import load_dotenv
from datetime import datetime

load_dotenv()


# ==========================================
# Azure Configuration
# ==========================================

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
OBJECT_ID = os.getenv("OBJECT_ID")


# ==========================================
# Generate Access Token
# ==========================================

def get_access_token():

    authority = f"https://login.microsoftonline.com/{TENANT_ID}"

    app = ConfidentialClientApplication(
        CLIENT_ID,
        authority=authority,
        client_credential=CLIENT_SECRET
    )

    token_result = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )

    if "access_token" not in token_result:
        return None

    return token_result["access_token"]


# ==========================================
# Create Teams Meeting
# ==========================================

def create_teams_meeting(
    candidateName,
    email,
    startDateTime,
    endDateTime,
    subject,
    interviewers,
    interviewersEmail
):

    # ==========================================
    # Access Token
    # ==========================================

    access_token = get_access_token()

    if access_token is None:

        return {
            "status": "Failed",
            "message": "Unable to generate access token."
        }


    # ==========================================
    # Headers
    # ==========================================

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }


    # ==========================================
    # Teams Meeting URL
    # ==========================================

    teams_meeting_url = (
        f"https://graph.microsoft.com/v1.0/"
        f"users/{OBJECT_ID}/onlineMeetings"
    )


    # ==========================================
    # Teams Meeting Payload
    # ==========================================

    teams_meeting_url_payload = {

        "startDateTime": startDateTime,

        "endDateTime": endDateTime,

        "subject": f"{subject} - {candidateName}"
    }


    # ==========================================
    # Create Meeting
    # ==========================================

    response = requests.post(
        teams_meeting_url,
        headers=headers,
        json=teams_meeting_url_payload
    )


    # ==========================================
    # Parse Response
    # ==========================================

    try:

        response_data = response.json()

    except Exception:

        return {
            "status": "Failed",
            "message": response.text
        }


    # ==========================================
    # Success
    # ==========================================

    if response.status_code == 201:

        start = datetime.fromisoformat(
            startDateTime.replace("Z", "+00:00")
        )

        end = datetime.fromisoformat(
            endDateTime.replace("Z", "+00:00")
        )

        duration = end - start


        return {

            "status": "Success",

            "candidateName": candidateName,

            "candidateEmail": email,

            "interviewers": interviewers,

            "interviewersEmail": interviewersEmail,

            "joinWebUrl": response_data.get(
                "joinWebUrl"
            ),

            "meetingCode": response_data.get(
                "meetingCode"
            ),

            "startDateTime": response_data.get(
                "startDateTime"
            ),

            "endDateTime": response_data.get(
                "endDateTime"
            ),

            "subject": response_data.get(
                "subject"
            ),

            "duration": str(duration)
        }


    # ==========================================
    # Failure
    # ==========================================

    return {

        "status": "Failed",

        "statusCode": response.status_code,

        "response": response_data
    }
