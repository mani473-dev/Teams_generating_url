from fastapi import FastAPI
from pydantic import BaseModel

from main import create_teams_meeting

app = FastAPI()


class TeamsMeetingRequest(BaseModel):
    candidateName: str
    startDateTime: str
    endDateTime: str
    subject: str


@app.post("/create-teams-meeting")
def execute(request: TeamsMeetingRequest):

    return create_teams_meeting(
        request.candidateName,
        request.startDateTime,
        request.endDateTime,
        request.subject
    )