from msal import ConfidentialClientApplication
import requests
import os

from dotenv import load_dotenv
from datetime import datetime


load_dotenv()




TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
OBJECT_ID = os.getenv("OBJECT_ID")


# ==========================================
# Generate Microsoft Graph Access Token
# ==========================================

def get_access_token():

    authority = (
        f"https://login.microsoftonline.com/{TENANT_ID}"
    )

    app = ConfidentialClientApplication(

        CLIENT_ID,

        authority=authority,

        client_credential=CLIENT_SECRET
    )

    token_result = app.acquire_token_for_client(

        scopes=[
            "https://graph.microsoft.com/.default"
        ]
    )

    if "access_token" not in token_result:

        return None

    return token_result["access_token"]




def create_teams_meeting(

    candidateName,

    email,

    startDateTime,

    endDateTime,

    subject,

    interviewers

):

    # ==========================================
    # Generate Access Token
    # ==========================================

    access_token = get_access_token()

    if access_token is None:

        return {

            "status": "Failed",

            "message": "Unable to generate access token."

        }




    headers = {

        "Authorization":
            f"Bearer {access_token}",

        "Content-Type":
            "application/json"

    }


    # ==========================================
    # Microsoft Graph Online Meeting API
    # ==========================================

    teams_meeting_url = (

        "https://graph.microsoft.com/v1.0/"

        f"users/{OBJECT_ID}/onlineMeetings"

    )


    # ==========================================
    # Teams Meeting Payload
    # ==========================================

    teams_meeting_payload = {

        "startDateTime":
            startDateTime,

        "endDateTime":
            endDateTime,

        "subject":
            f"{subject} - {candidateName}",

        "joinMeetingIdSettings": {

            "isPasscodeRequired":
                True

        }

    }


    # ==========================================
    # Create Meeting
    # ==========================================

    response = requests.post(

        teams_meeting_url,

        headers=headers,

        json=teams_meeting_payload

    )


    # ==========================================
    # Read Response
    # ==========================================

    try:

        response_data = response.json()

    except Exception:

        return {

            "status": "Failed",

            "message": response.text

        }


    # ==========================================
    # Meeting Created Successfully
    # ==========================================

    if response.status_code == 201:


        # --------------------------------------
        # Calculate Duration
        # --------------------------------------

        start = datetime.fromisoformat(

            startDateTime.replace(

                "Z",

                "+00:00"

            )

        )

        end = datetime.fromisoformat(

            endDateTime.replace(

                "Z",

                "+00:00"

            )

        )

        duration = end - start


        # --------------------------------------
        # Meeting URL
        # --------------------------------------

        join_web_url = (

            response_data.get(

                "joinWebUrl"

            )

        )


        # --------------------------------------
        # Meeting ID + Passcode
        # --------------------------------------

        join_meeting_settings = (

            response_data.get(

                "joinMeetingIdSettings",

                {}

            )

        )


        meeting_id = (

            join_meeting_settings.get(

                "joinMeetingId"

            )

        )


        passcode = (

            join_meeting_settings.get(

                "passcode"

            )

        )


        # --------------------------------------
        # Success Response
        # --------------------------------------

        return {

            "status":
                "Success",

            "candidateName":
                candidateName,

            "candidateEmail":
                email,

            "interviewers":
                interviewers,

            "joinWebUrl":
                join_web_url,

            "meetingCode":
                meeting_id,

            "passcode":
                passcode,

            "startDateTime":
                response_data.get(

                    "startDateTime"

                ),

            "endDateTime":
                response_data.get(

                    "endDateTime"

                ),

            "subject":
                response_data.get(

                    "subject"

                ),

            "duration":
                str(duration)

        }


    # ==========================================
    # Microsoft Graph Error
    # ==========================================

    return {

        "status":
            "Failed",

        "statusCode":
            response.status_code,

        "response":
            response_data

    }
