from pydantic import BaseModel, Field
from typing import List, Optional

class ExtractedInfo(BaseModel):
    title: str = Field(description="Title or subject of the document")
    summary: str = Field(description="2-4 sentence summary of the document")
    key_topics: List[str] = Field(description="Main topics or themes discussed")
    entities: List[str] = Field(description="People, organizations, or places mentioned")
    dates: List[str] = Field(description="Important dates or time references found")
    action_items: Optional[List[str]] = Field(default=None, description="Any action items, tasks, or recommendations")
    sentiment: Optional[str] = Field(default=None, description="Overall tone: positive, neutral, or negative")