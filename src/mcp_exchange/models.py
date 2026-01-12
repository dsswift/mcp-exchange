"""Pydantic models for Microsoft Graph API responses."""

from __future__ import annotations

from datetime import datetime
from typing import Any

from pydantic import BaseModel, Field, field_validator


class EmailAddress(BaseModel):
    """Email address with optional name."""

    address: str
    name: str | None = None


class Recipient(BaseModel):
    """Email recipient."""

    email_address: EmailAddress = Field(alias="emailAddress")


class ItemBody(BaseModel):
    """Body content for messages and events."""

    content_type: str = Field(alias="contentType")  # "text" or "html"
    content: str


class MailFolder(BaseModel):
    """Mail folder in Exchange."""

    id: str
    display_name: str = Field(alias="displayName")
    parent_folder_id: str | None = Field(None, alias="parentFolderId")
    child_folder_count: int = Field(0, alias="childFolderCount")
    unread_item_count: int = Field(0, alias="unreadItemCount")
    total_item_count: int = Field(0, alias="totalItemCount")


class Message(BaseModel):
    """Email message."""

    id: str
    subject: str | None = None
    body: ItemBody | None = None
    body_preview: str | None = Field(None, alias="bodyPreview")
    sender: Recipient | None = None
    from_: Recipient | None = Field(None, alias="from")
    to_recipients: list[Recipient] = Field(default_factory=list, alias="toRecipients")
    cc_recipients: list[Recipient] = Field(default_factory=list, alias="ccRecipients")
    bcc_recipients: list[Recipient] = Field(default_factory=list, alias="bccRecipients")
    received_date_time: datetime | None = Field(None, alias="receivedDateTime")
    sent_date_time: datetime | None = Field(None, alias="sentDateTime")
    has_attachments: bool = Field(False, alias="hasAttachments")
    is_read: bool = Field(False, alias="isRead")
    is_draft: bool = Field(False, alias="isDraft")
    importance: str = "normal"  # "low", "normal", "high"
    parent_folder_id: str | None = Field(None, alias="parentFolderId")
    web_link: str | None = Field(None, alias="webLink")


class MessageCreate(BaseModel):
    """Request model for creating a draft message."""

    subject: str | None = None
    body: ItemBody | None = None
    to_recipients: list[Recipient] = Field(default_factory=list, alias="toRecipients")
    cc_recipients: list[Recipient] = Field(default_factory=list, alias="ccRecipients")
    importance: str = "normal"

    def to_api_payload(self) -> dict[str, Any]:
        """Convert to Microsoft Graph API payload."""
        payload: dict[str, Any] = {}

        if self.subject:
            payload["subject"] = self.subject

        if self.body:
            payload["body"] = {
                "contentType": self.body.content_type,
                "content": self.body.content,
            }

        if self.to_recipients:
            payload["toRecipients"] = [
                {"emailAddress": {"address": r.email_address.address, "name": r.email_address.name}}
                for r in self.to_recipients
            ]

        if self.cc_recipients:
            payload["ccRecipients"] = [
                {"emailAddress": {"address": r.email_address.address, "name": r.email_address.name}}
                for r in self.cc_recipients
            ]

        if self.importance != "normal":
            payload["importance"] = self.importance

        return payload


class DateTimeTimeZone(BaseModel):
    """Date/time with timezone for calendar events."""

    date_time: str = Field(alias="dateTime")
    time_zone: str = Field(alias="timeZone")

    @property
    def as_datetime(self) -> datetime:
        """Parse the datetime string (note: timezone info may need handling)."""
        # Graph API returns ISO format without timezone offset in dateTime field
        return datetime.fromisoformat(self.date_time.replace("Z", "+00:00"))


class Attendee(BaseModel):
    """Calendar event attendee."""

    type: str  # "required", "optional", "resource"
    status: dict[str, str] | None = None  # {"response": "accepted", "time": "..."}
    email_address: EmailAddress = Field(alias="emailAddress")


class Location(BaseModel):
    """Event location."""

    display_name: str = Field("", alias="displayName")
    location_type: str | None = Field(None, alias="locationType")
    unique_id: str | None = Field(None, alias="uniqueId")
    unique_id_type: str | None = Field(None, alias="uniqueIdType")


class PatternedRecurrence(BaseModel):
    """Recurrence pattern for events."""

    pattern: dict[str, Any]
    range: dict[str, Any]


class Calendar(BaseModel):
    """Calendar in Exchange."""

    id: str
    name: str
    color: str | None = None
    change_key: str | None = Field(None, alias="changeKey")
    can_share: bool = Field(False, alias="canShare")
    can_view_private_items: bool = Field(False, alias="canViewPrivateItems")
    can_edit: bool = Field(False, alias="canEdit")
    owner: EmailAddress | None = None
    is_default_calendar: bool = Field(False, alias="isDefaultCalendar")


class Event(BaseModel):
    """Calendar event."""

    id: str
    subject: str | None = None
    body: ItemBody | None = None
    body_preview: str | None = Field(None, alias="bodyPreview")
    start: DateTimeTimeZone | None = None
    end: DateTimeTimeZone | None = None
    location: Location | None = None
    locations: list[Location] = Field(default_factory=list)
    attendees: list[Attendee] = Field(default_factory=list)
    organizer: Recipient | None = None
    is_all_day: bool = Field(False, alias="isAllDay")
    is_cancelled: bool = Field(False, alias="isCancelled")
    is_organizer: bool = Field(False, alias="isOrganizer")
    recurrence: PatternedRecurrence | None = None
    series_master_id: str | None = Field(None, alias="seriesMasterId")
    show_as: str | None = Field(None, alias="showAs")  # "free", "tentative", "busy", "oof", etc.
    type: str | None = None  # "singleInstance", "occurrence", "exception", "seriesMaster"
    importance: str = "normal"
    sensitivity: str = "normal"
    categories: list[str] = Field(default_factory=list)
    web_link: str | None = Field(None, alias="webLink")
    online_meeting_url: str | None = Field(None, alias="onlineMeetingUrl")
    created_date_time: datetime | None = Field(None, alias="createdDateTime")
    last_modified_date_time: datetime | None = Field(None, alias="lastModifiedDateTime")


class ScheduleItem(BaseModel):
    """Individual schedule item in free/busy response."""

    status: str  # "free", "tentative", "busy", "oof", "workingElsewhere", "unknown"
    start: DateTimeTimeZone
    end: DateTimeTimeZone
    subject: str | None = None
    location: str | None = None
    is_private: bool = Field(False, alias="isPrivate")


class ScheduleInformation(BaseModel):
    """Schedule information for a user (free/busy response)."""

    schedule_id: str = Field(alias="scheduleId")  # Email address
    availability_view: str = Field("", alias="availabilityView")  # Encoded availability string
    schedule_items: list[ScheduleItem] = Field(default_factory=list, alias="scheduleItems")
    working_hours: dict[str, Any] | None = Field(None, alias="workingHours")
    error: dict[str, Any] | None = None

    @field_validator("schedule_items", mode="before")
    @classmethod
    def handle_null_schedule_items(cls, v: Any) -> list[Any]:
        """Handle null schedule_items from API."""
        return v or []


class FreeBusyRequest(BaseModel):
    """Request model for getSchedule API."""

    schedules: list[str]  # Email addresses
    start_time: DateTimeTimeZone = Field(alias="startTime")
    end_time: DateTimeTimeZone = Field(alias="endTime")
    availability_view_interval: int = Field(30, alias="availabilityViewInterval")

    def to_api_payload(self) -> dict[str, Any]:
        """Convert to Microsoft Graph API payload."""
        return {
            "schedules": self.schedules,
            "startTime": {
                "dateTime": self.start_time.date_time,
                "timeZone": self.start_time.time_zone,
            },
            "endTime": {
                "dateTime": self.end_time.date_time,
                "timeZone": self.end_time.time_zone,
            },
            "availabilityViewInterval": self.availability_view_interval,
        }
