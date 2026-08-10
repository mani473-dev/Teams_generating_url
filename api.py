from fastapi import FastAPI
from pydantic import BaseModel
from typing import List

from main import create_teams_meeting

app = FastAPI()


class TeamsMeetingRequest(BaseModel):
    candidateName: str
    email: str
    startDateTime: str
    endDateTime: str
    subject: str
    interviewers: List[str]


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
