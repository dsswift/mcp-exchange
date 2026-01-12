"""Exchange Online MCP Server with FastMCP tools."""

from __future__ import annotations

import json
import logging
import sys
from collections.abc import AsyncIterator
from contextlib import asynccontextmanager
from datetime import datetime
from typing import Any

from mcp.server.fastmcp import FastMCP

from .auth import GraphAuthenticator
from .client import ExchangeClient, GraphError, GraphNotFoundError
from .config import load_config

# Configure logging to stderr (stdout is reserved for MCP protocol)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    stream=sys.stderr,
)
logger = logging.getLogger(__name__)


@asynccontextmanager
async def lifespan(_server: FastMCP) -> AsyncIterator[dict[str, Any]]:
    """Manage server lifecycle and shared resources."""
    config = load_config()
    logger.info("Exchange Online MCP Server starting")

    authenticator = GraphAuthenticator(config)

    async with ExchangeClient(config, authenticator) as client:
        yield {"client": client, "config": config, "authenticator": authenticator}

    logger.info("Exchange Online MCP Server shutting down")


# Initialize FastMCP server
mcp = FastMCP(
    name="exchange",
    lifespan=lifespan,
)


# =========================================================================
# Formatting helpers
# =========================================================================


def format_folder(folder: Any) -> dict[str, Any]:
    """Format a mail folder for display."""
    return {
        "id": folder.id,
        "displayName": folder.display_name,
        "unreadItemCount": folder.unread_item_count,
        "totalItemCount": folder.total_item_count,
    }


def format_message(message: Any, include_body: bool = False) -> dict[str, Any]:
    """Format an email message for display."""
    result: dict[str, Any] = {
        "id": message.id,
        "subject": message.subject,
    }

    if message.sender and message.sender.email_address:
        result["sender"] = {
            "name": message.sender.email_address.name,
            "address": message.sender.email_address.address,
        }

    if message.received_date_time:
        result["receivedDateTime"] = message.received_date_time.isoformat()

    result["isRead"] = message.is_read
    result["hasAttachments"] = message.has_attachments
    result["importance"] = message.importance

    if message.body_preview:
        result["preview"] = message.body_preview[:200]

    if include_body and message.body:
        result["body"] = {
            "contentType": message.body.content_type,
            "content": message.body.content,
        }

    if message.to_recipients:
        result["toRecipients"] = [
            {"name": r.email_address.name, "address": r.email_address.address}
            for r in message.to_recipients
        ]

    if message.web_link:
        result["webLink"] = message.web_link

    return result


def format_calendar(calendar: Any) -> dict[str, Any]:
    """Format a calendar for display."""
    result: dict[str, Any] = {
        "id": calendar.id,
        "name": calendar.name,
        "isDefaultCalendar": calendar.is_default_calendar,
    }

    if calendar.owner:
        result["owner"] = calendar.owner.address

    if calendar.color:
        result["color"] = calendar.color

    return result


def format_event(event: Any) -> dict[str, Any]:
    """Format a calendar event for display."""
    result: dict[str, Any] = {
        "id": event.id,
        "subject": event.subject,
    }

    if event.start:
        result["start"] = {
            "dateTime": event.start.date_time,
            "timeZone": event.start.time_zone,
        }

    if event.end:
        result["end"] = {
            "dateTime": event.end.date_time,
            "timeZone": event.end.time_zone,
        }

    if event.location and event.location.display_name:
        result["location"] = event.location.display_name

    result["isAllDay"] = event.is_all_day
    result["showAs"] = event.show_as

    if event.organizer and event.organizer.email_address:
        result["organizer"] = {
            "name": event.organizer.email_address.name,
            "address": event.organizer.email_address.address,
        }

    if event.attendees:
        result["attendees"] = [
            {
                "name": a.email_address.name,
                "address": a.email_address.address,
                "type": a.type,
                "response": a.status.get("response") if a.status else None,
            }
            for a in event.attendees
        ]

    if event.body and event.body.content:
        result["body"] = {
            "contentType": event.body.content_type,
            "content": event.body.content,
        }

    if event.recurrence:
        result["recurrence"] = event.recurrence.model_dump()

    if event.web_link:
        result["webLink"] = event.web_link

    if event.online_meeting_url:
        result["onlineMeetingUrl"] = event.online_meeting_url

    return result


def format_schedule(schedule: Any) -> dict[str, Any]:
    """Format schedule information for display."""
    result: dict[str, Any] = {
        "email": schedule.schedule_id,
        "availabilityView": schedule.availability_view,
    }

    if schedule.error:
        result["error"] = schedule.error

    if schedule.schedule_items:
        result["scheduleItems"] = [
            {
                "status": item.status,
                "start": item.start.date_time,
                "end": item.end.date_time,
                "subject": item.subject if not item.is_private else "[Private]",
                "location": item.location if not item.is_private else None,
            }
            for item in schedule.schedule_items
        ]

    return result


# =========================================================================
# Email Tools
# =========================================================================


@mcp.tool()
async def list_mail_folders() -> str:
    """List all mail folders in the mailbox.

    Returns:
        JSON array of mail folders with id, name, and item counts.

    Examples:
        - list_mail_folders() - Get all folders (Inbox, Sent, Archive, etc.)
    """
    ctx = mcp.get_context()
    client: ExchangeClient = ctx.request_context.lifespan_context["client"]

    try:
        folders = await client.list_mail_folders()

        result = {
            "count": len(folders),
            "folders": [format_folder(f) for f in folders],
        }

        return json.dumps(result, indent=2)

    except GraphError as e:
        return json.dumps({"error": str(e)})


@mcp.tool()
async def search_emails(
    folder: str | None = None,
    sender: str | None = None,
    subject: str | None = None,
    from_date: str | None = None,
    to_date: str | None = None,
    is_read: bool | None = None,
    has_attachments: bool | None = None,
    limit: int = 25,
) -> str:
    """Search for emails with various filters.

    Args:
        folder: Folder name (e.g., 'Inbox', 'Archive') or folder ID. Default: Inbox.
        sender: Filter by sender email address.
        subject: Filter by subject (contains search).
        from_date: Filter emails received after this date (ISO format: YYYY-MM-DD).
        to_date: Filter emails received before this date (ISO format: YYYY-MM-DD).
        is_read: Filter by read status (true/false).
        has_attachments: Filter by attachment presence (true/false).
        limit: Maximum number of results (default 25, max 100).

    Returns:
        JSON array of matching emails with key fields.

    Examples:
        - search_emails() - Recent emails from Inbox
        - search_emails(sender="boss@company.com") - Emails from specific sender
        - search_emails(subject="meeting", is_read=False) - Unread emails about meetings
        - search_emails(from_date="2024-01-01", has_attachments=True) - Recent emails with attachments
    """
    ctx = mcp.get_context()
    client: ExchangeClient = ctx.request_context.lifespan_context["client"]

    limit = min(max(1, limit), 100)

    # Parse dates
    parsed_from_date = None
    parsed_to_date = None

    if from_date:
        try:
            parsed_from_date = datetime.fromisoformat(from_date)
        except ValueError:
            return json.dumps({"error": f"Invalid from_date format: {from_date}. Use YYYY-MM-DD."})

    if to_date:
        try:
            parsed_to_date = datetime.fromisoformat(to_date)
        except ValueError:
            return json.dumps({"error": f"Invalid to_date format: {to_date}. Use YYYY-MM-DD."})

    try:
        # Resolve folder name to ID if needed
        folder_id = None
        if folder:
            # Check if it's a well-known folder name
            well_known = ["inbox", "archive", "drafts", "sentitems", "deleteditems", "junkemail"]
            if folder.lower() in well_known:
                folder_id = folder.lower()
            else:
                # Try to find by display name, fall back to assuming it's a folder ID
                folder_obj = await client.get_folder_by_name(folder)
                folder_id = folder_obj.id if folder_obj else folder

        messages = await client.list_messages(
            folder_id=folder_id,
            sender=sender,
            subject=subject,
            from_date=parsed_from_date,
            to_date=parsed_to_date,
            is_read=is_read,
            has_attachments=has_attachments,
            limit=limit,
        )

        # Build filter description for response
        filters_applied = []
        if folder:
            filters_applied.append(f"folder={folder}")
        if sender:
            filters_applied.append(f"sender={sender}")
        if subject:
            filters_applied.append(f"subject contains '{subject}'")
        if from_date:
            filters_applied.append(f"from={from_date}")
        if to_date:
            filters_applied.append(f"to={to_date}")
        if is_read is not None:
            filters_applied.append(f"isRead={is_read}")
        if has_attachments is not None:
            filters_applied.append(f"hasAttachments={has_attachments}")

        result = {
            "count": len(messages),
            "filters": ", ".join(filters_applied) if filters_applied else "none",
            "messages": [format_message(m) for m in messages],
        }

        return json.dumps(result, indent=2, default=str)

    except GraphError as e:
        return json.dumps({"error": str(e)})


@mcp.tool()
async def get_email(message_id: str) -> str:
    """Get detailed information about a specific email.

    Args:
        message_id: The message ID.

    Returns:
        JSON object with full email details including body.
    """
    ctx = mcp.get_context()
    client: ExchangeClient = ctx.request_context.lifespan_context["client"]

    try:
        message = await client.get_message(message_id)
        return json.dumps(format_message(message, include_body=True), indent=2, default=str)

    except GraphNotFoundError:
        return json.dumps({"error": f"Message '{message_id}' not found"})
    except GraphError as e:
        return json.dumps({"error": str(e)})


@mcp.tool()
async def archive_email(message_id: str) -> str:
    """Archive an email (move to Archive folder).

    Args:
        message_id: The message ID to archive.

    Returns:
        JSON object confirming the archive or error message.
    """
    ctx = mcp.get_context()
    client: ExchangeClient = ctx.request_context.lifespan_context["client"]

    try:
        message = await client.archive_message(message_id)

        return json.dumps({
            "archived": True,
            "message_id": message.id,
            "subject": message.subject,
            "message": f"Email '{message.subject}' has been archived",
        }, indent=2)

    except GraphNotFoundError:
        return json.dumps({"error": f"Message '{message_id}' not found"})
    except GraphError as e:
        return json.dumps({"error": str(e)})


@mcp.tool()
async def delete_email(message_id: str) -> str:
    """Delete an email (moves to Deleted Items).

    Args:
        message_id: The message ID to delete.

    Returns:
        JSON object confirming deletion or error message.
    """
    ctx = mcp.get_context()
    client: ExchangeClient = ctx.request_context.lifespan_context["client"]

    try:
        # Get message details first for confirmation
        message = await client.get_message(message_id)
        subject = message.subject

        await client.delete_message(message_id)

        return json.dumps({
            "deleted": True,
            "message_id": message_id,
            "subject": subject,
            "message": f"Email '{subject}' has been moved to Deleted Items",
        }, indent=2)

    except GraphNotFoundError:
        return json.dumps({"error": f"Message '{message_id}' not found"})
    except GraphError as e:
        return json.dumps({"error": str(e)})


@mcp.tool()
async def create_draft(
    subject: str | None = None,
    body: str | None = None,
    body_type: str = "text",
    to_recipients: str | None = None,
    cc_recipients: str | None = None,
    importance: str = "normal",
) -> str:
    """Create a draft email.

    Args:
        subject: Email subject.
        body: Email body content.
        body_type: Body content type - "text" or "html" (default: text).
        to_recipients: Comma-separated list of recipient email addresses.
        cc_recipients: Comma-separated list of CC recipient email addresses.
        importance: Email importance - "low", "normal", or "high" (default: normal).

    Returns:
        JSON object with the created draft details.

    Examples:
        - create_draft(subject="Meeting follow-up", to_recipients="boss@company.com")
        - create_draft(subject="Report", body="<h1>Report</h1>", body_type="html", to_recipients="team@company.com")
    """
    ctx = mcp.get_context()
    client: ExchangeClient = ctx.request_context.lifespan_context["client"]

    # Parse recipient lists
    to_list = [e.strip() for e in to_recipients.split(",")] if to_recipients else None
    cc_list = [e.strip() for e in cc_recipients.split(",")] if cc_recipients else None

    try:
        message = await client.create_draft(
            subject=subject,
            body=body,
            body_type=body_type,
            to_recipients=to_list,
            cc_recipients=cc_list,
            importance=importance,
        )

        result = format_message(message)
        result["_created"] = True
        result["message"] = f"Draft created: '{subject}'"

        return json.dumps(result, indent=2, default=str)

    except GraphError as e:
        return json.dumps({"error": str(e)})


# =========================================================================
# Calendar Tools
# =========================================================================


@mcp.tool()
async def list_calendars() -> str:
    """List all calendars in the mailbox.

    Returns:
        JSON array of calendars with id, name, and owner.
    """
    ctx = mcp.get_context()
    client: ExchangeClient = ctx.request_context.lifespan_context["client"]

    try:
        calendars = await client.list_calendars()

        result = {
            "count": len(calendars),
            "calendars": [format_calendar(c) for c in calendars],
        }

        return json.dumps(result, indent=2)

    except GraphError as e:
        return json.dumps({"error": str(e)})


@mcp.tool()
async def list_events(
    calendar_id: str | None = None,
    start_date: str | None = None,
    end_date: str | None = None,
    limit: int = 25,
) -> str:
    """List calendar events in a date range.

    Args:
        calendar_id: Calendar ID (default: primary calendar).
        start_date: Filter events starting after this date (ISO format: YYYY-MM-DD).
        end_date: Filter events ending before this date (ISO format: YYYY-MM-DD).
        limit: Maximum number of results (default 25, max 100).

    Returns:
        JSON array of events with full details.

    Examples:
        - list_events() - Upcoming events from primary calendar
        - list_events(start_date="2024-01-15", end_date="2024-01-22") - Events in date range
    """
    ctx = mcp.get_context()
    client: ExchangeClient = ctx.request_context.lifespan_context["client"]

    limit = min(max(1, limit), 100)

    # Parse dates
    parsed_start = None
    parsed_end = None

    if start_date:
        try:
            parsed_start = datetime.fromisoformat(start_date)
        except ValueError:
            return json.dumps({"error": f"Invalid start_date format: {start_date}. Use YYYY-MM-DD."})

    if end_date:
        try:
            parsed_end = datetime.fromisoformat(end_date)
        except ValueError:
            return json.dumps({"error": f"Invalid end_date format: {end_date}. Use YYYY-MM-DD."})

    try:
        events = await client.list_events(
            calendar_id=calendar_id,
            start_date=parsed_start,
            end_date=parsed_end,
            limit=limit,
        )

        result = {
            "count": len(events),
            "calendar": calendar_id or "primary",
            "events": [format_event(e) for e in events],
        }

        return json.dumps(result, indent=2, default=str)

    except GraphError as e:
        return json.dumps({"error": str(e)})


@mcp.tool()
async def get_event(event_id: str) -> str:
    """Get detailed information about a specific calendar event.

    Args:
        event_id: The event ID.

    Returns:
        JSON object with full event details.
    """
    ctx = mcp.get_context()
    client: ExchangeClient = ctx.request_context.lifespan_context["client"]

    try:
        event = await client.get_event(event_id)
        return json.dumps(format_event(event), indent=2, default=str)

    except GraphNotFoundError:
        return json.dumps({"error": f"Event '{event_id}' not found"})
    except GraphError as e:
        return json.dumps({"error": str(e)})


@mcp.tool()
async def get_free_busy(
    emails: str,
    start_time: str,
    end_time: str,
    timezone: str = "UTC",
    interval_minutes: int = 30,
) -> str:
    """Get free/busy schedule for one or more users.

    Use this to find available meeting times by checking when people are free or busy.

    Args:
        emails: Comma-separated list of email addresses to check (max 20).
        start_time: Start of time range (ISO format: YYYY-MM-DDTHH:MM:SS).
        end_time: End of time range (ISO format: YYYY-MM-DDTHH:MM:SS).
        timezone: Timezone for the query (default: UTC).
        interval_minutes: Granularity of availability view in minutes (default: 30).

    Returns:
        JSON object with schedule information for each user.
        - availabilityView: encoded string where each character represents a time slot
          (0=free, 1=tentative, 2=busy, 3=out of office, 4=working elsewhere)
        - scheduleItems: list of busy times with subject/location (if shared)

    Examples:
        - get_free_busy(emails="chris@company.com", start_time="2024-01-15T09:00:00", end_time="2024-01-15T17:00:00")
        - get_free_busy(emails="chris@company.com,alex@company.com", start_time="2024-01-15T08:00:00", end_time="2024-01-15T18:00:00", timezone="America/New_York")
    """
    ctx = mcp.get_context()
    client: ExchangeClient = ctx.request_context.lifespan_context["client"]

    # Parse email list
    email_list = [e.strip() for e in emails.split(",")]

    if len(email_list) > 20:
        return json.dumps({"error": "Maximum 20 email addresses allowed"})

    # Parse times
    try:
        parsed_start = datetime.fromisoformat(start_time)
    except ValueError:
        return json.dumps({"error": f"Invalid start_time format: {start_time}. Use YYYY-MM-DDTHH:MM:SS."})

    try:
        parsed_end = datetime.fromisoformat(end_time)
    except ValueError:
        return json.dumps({"error": f"Invalid end_time format: {end_time}. Use YYYY-MM-DDTHH:MM:SS."})

    try:
        schedules = await client.get_free_busy(
            emails=email_list,
            start_time=parsed_start,
            end_time=parsed_end,
            timezone=timezone,
            interval_minutes=interval_minutes,
        )

        result = {
            "startTime": start_time,
            "endTime": end_time,
            "timezone": timezone,
            "intervalMinutes": interval_minutes,
            "schedules": [format_schedule(s) for s in schedules],
        }

        return json.dumps(result, indent=2, default=str)

    except GraphError as e:
        return json.dumps({"error": str(e)})


def run_server() -> None:
    """Run the MCP server."""
    mcp.run()
