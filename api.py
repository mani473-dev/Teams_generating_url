from fastapi import FastAPI
from pydantic import BaseModel
from typing import List

from main import create_teams_meeting

app = FastAPI()


class Interviewer(BaseModel):
    name: str
    email: str


class TeamsMeetingRequest(BaseModel):
    candidateName: str
    email: str
    startDateTime: str
    endDateTime: str
    subject: str
    interviewers: List[Interviewer]


@app.get("/")
def home():
    return {
        "message": "Teams Meeting API is Running"
    }


@app.post("/create-teams-meeting")
def execute(request: TeamsMeetingRequest):

    return create_teams_meeting(
        request.candidateName,
        request.email,
        request.startDateTime,
        request.endDateTime,
        request.subject,
        request.interviewers
    )
