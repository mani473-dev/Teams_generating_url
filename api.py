from fastapi import FastAPI
from pydantic import BaseModel
from typing import List


from main import create_teams_meeting


app = FastAPI()


# ==========================================
# Interviewer Model
# ==========================================

class Interviewer(BaseModel):

    name: str
    email: str


# ==========================================
# Request Model
# ==========================================

class TeamsMeetingRequest(BaseModel):

    candidateName: str

    email: str

    startDateTime: str

    endDateTime: str

    subject: str

    interviewers: List[Interviewer]


# ==========================================
# Home
# ==========================================

@app.get("/")
def home():

    return {
        "message": "Teams Meeting API is Running"
    }


# ==========================================
# Create Teams Meeting
# ==========================================

@app.post("/create-teams-meeting")
def execute(request: TeamsMeetingRequest):

    # Convert Pydantic objects into normal dictionaries

    interviewers = [
        interviewer.model_dump()
        for interviewer in request.interviewers
    ]

    return create_teams_meeting(

        request.candidateName,

        request.email,

        request.startDateTime,

        request.endDateTime,

        request.subject,

        interviewers
    )
