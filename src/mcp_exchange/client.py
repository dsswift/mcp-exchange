"""Microsoft Graph API client for Exchange Online."""

from __future__ import annotations

import logging
from datetime import datetime
from typing import Any

import httpx

from .auth import AuthError, GraphAuthenticator
from .config import ExchangeConfig
from .models import (
    Calendar,
    Event,
    MailFolder,
    Message,
    ScheduleInformation,
)

logger = logging.getLogger(__name__)


class GraphError(Exception):
    """Base exception for Microsoft Graph API errors."""

    def __init__(self, message: str, status_code: int | None = None) -> None:
        super().__init__(message)
        self.status_code = status_code


class GraphAuthError(GraphError):
    """Authentication/authorization error."""


class GraphNotFoundError(GraphError):
    """Resource not found error."""


# Default fields to request from the API
MESSAGE_FIELDS = (
    "id,subject,bodyPreview,body,sender,from,toRecipients,ccRecipients,"
    "receivedDateTime,sentDateTime,hasAttachments,isRead,isDraft,"
    "importance,parentFolderId,webLink"
)

FOLDER_FIELDS = (
    "id,displayName,parentFolderId,childFolderCount,unreadItemCount,totalItemCount"
)

CALENDAR_FIELDS = (
    "id,name,color,canShare,canViewPrivateItems,canEdit,owner,isDefaultCalendar"
)

EVENT_FIELDS = (
    "id,subject,body,bodyPreview,start,end,location,locations,attendees,"
    "organizer,isAllDay,isCancelled,isOrganizer,recurrence,seriesMasterId,"
    "showAs,type,importance,sensitivity,categories,webLink,onlineMeetingUrl,"
    "createdDateTime,lastModifiedDateTime"
)


class ExchangeClient:
    """Async client for Microsoft Graph API - Exchange Online operations."""

    def __init__(self, config: ExchangeConfig, authenticator: GraphAuthenticator) -> None:
        """Initialize the client with configuration and authenticator."""
        self.config = config
        self.authenticator = authenticator
        self._client: httpx.AsyncClient | None = None

    async def __aenter__(self) -> ExchangeClient:
        """Enter async context."""
        self._client = httpx.AsyncClient(
            base_url=self.config.graph_base_url,
            headers={
                "Accept": "application/json",
                "Content-Type": "application/json",
            },
            timeout=self.config.timeout,
        )
        return self

    async def __aexit__(self, *args: Any) -> None:
        """Exit async context."""
        if self._client:
            await self._client.aclose()
            self._client = None

    @property
    def client(self) -> httpx.AsyncClient:
        """Get the HTTP client, ensuring it's initialized."""
        if self._client is None:
            raise RuntimeError(
                "Client not initialized. Use 'async with ExchangeClient(config)' context."
            )
        return self._client

    def _get_auth_header(self) -> dict[str, str]:
        """Get current auth header (refreshes token if needed)."""
        try:
            return self.authenticator.get_auth_header()
        except AuthError as e:
            raise GraphAuthError(str(e)) from e

    def _handle_error(self, response: httpx.Response) -> None:
        """Handle HTTP error responses."""
        status = response.status_code

        try:
            error_data = response.json()
            error_info = error_data.get("error", {})
            message = error_info.get("message", str(status))
            code = error_info.get("code", "")
        except Exception:
            message = response.text or f"HTTP {status}"
            code = ""

        if status == 401:
            raise GraphAuthError(
                f"Authentication failed: {message}",
                status_code=status,
            )
        elif status == 403:
            raise GraphAuthError(
                f"Permission denied: {message}",
                status_code=status,
            )
        elif status == 404:
            raise GraphNotFoundError(
                f"Resource not found: {message}",
                status_code=status,
            )
        else:
            raise GraphError(
                f"API error ({status}, {code}): {message}",
                status_code=status,
            )

    # =========================================================================
    # Mail Folder Operations
    # =========================================================================

    async def list_mail_folders(self) -> list[MailFolder]:
        """List all mail folders.

        Returns:
            List of mail folders.
        """
        logger.debug("Listing mail folders")
        response = await self.client.get(
            "/me/mailFolders",
            headers=self._get_auth_header(),
            params={"$select": FOLDER_FIELDS, "$top": 100},
        )

        if not response.is_success:
            self._handle_error(response)

        data = response.json()
        return [MailFolder.model_validate(item) for item in data.get("value", [])]

    async def get_folder_by_name(self, name: str) -> MailFolder | None:
        """Get a mail folder by display name.

        Args:
            name: Folder display name (e.g., "Archive", "Inbox").

        Returns:
            The folder if found, None otherwise.
        """
        folders = await self.list_mail_folders()
        name_lower = name.lower()
        return next(
            (f for f in folders if f.display_name.lower() == name_lower),
            None
        )

    # =========================================================================
    # Message Operations
    # =========================================================================

    async def list_messages(
        self,
        folder_id: str | None = None,
        sender: str | None = None,
        subject: str | None = None,
        from_date: datetime | None = None,
        to_date: datetime | None = None,
        is_read: bool | None = None,
        has_attachments: bool | None = None,
        limit: int = 25,
        skip: int = 0,
    ) -> list[Message]:
        """List messages with optional filters.

        Args:
            folder_id: Folder ID to list from (default: Inbox).
            sender: Filter by sender email address.
            subject: Filter by subject (contains).
            from_date: Filter messages received after this date.
            to_date: Filter messages received before this date.
            is_read: Filter by read/unread status.
            has_attachments: Filter by attachment presence.
            limit: Maximum number of results.
            skip: Number of results to skip.

        Returns:
            List of messages.
        """
        # Build the endpoint
        if folder_id:
            endpoint = f"/me/mailFolders/{folder_id}/messages"
        else:
            endpoint = "/me/mailFolders/inbox/messages"

        # Build filter parts
        filters: list[str] = []

        if sender:
            filters.append(f"from/emailAddress/address eq '{sender}'")

        if subject:
            # Use contains for subject search
            filters.append(f"contains(subject, '{subject}')")

        if from_date:
            filters.append(f"receivedDateTime ge {from_date.isoformat()}")

        if to_date:
            filters.append(f"receivedDateTime le {to_date.isoformat()}")

        if is_read is not None:
            filters.append(f"isRead eq {str(is_read).lower()}")

        if has_attachments is not None:
            filters.append(f"hasAttachments eq {str(has_attachments).lower()}")

        params: dict[str, Any] = {
            "$select": MESSAGE_FIELDS,
            "$top": limit,
            "$skip": skip,
            "$orderby": "receivedDateTime desc",
        }

        if filters:
            params["$filter"] = " and ".join(filters)

        logger.debug("Listing messages from %s with params: %s", endpoint, params)
        response = await self.client.get(
            endpoint,
            headers=self._get_auth_header(),
            params=params,
        )

        if not response.is_success:
            self._handle_error(response)

        data = response.json()
        return [Message.model_validate(item) for item in data.get("value", [])]

    async def get_message(self, message_id: str) -> Message:
        """Get a single message by ID.

        Args:
            message_id: Message ID.

        Returns:
            The message details.
        """
        logger.debug("Getting message: %s", message_id)
        response = await self.client.get(
            f"/me/messages/{message_id}",
            headers=self._get_auth_header(),
            params={"$select": MESSAGE_FIELDS},
        )

        if not response.is_success:
            self._handle_error(response)

        return Message.model_validate(response.json())

    async def move_message(self, message_id: str, destination_folder_id: str) -> Message:
        """Move a message to another folder.

        Args:
            message_id: Message ID to move.
            destination_folder_id: Destination folder ID.

        Returns:
            The moved message.
        """
        logger.debug("Moving message %s to folder %s", message_id, destination_folder_id)
        response = await self.client.post(
            f"/me/messages/{message_id}/move",
            headers=self._get_auth_header(),
            json={"destinationId": destination_folder_id},
        )

        if not response.is_success:
            self._handle_error(response)

        return Message.model_validate(response.json())

    async def archive_message(self, message_id: str) -> Message:
        """Archive a message (move to Archive folder).

        Args:
            message_id: Message ID to archive.

        Returns:
            The archived message.
        """
        # Use well-known folder name "archive"
        return await self.move_message(message_id, "archive")

    async def delete_message(self, message_id: str) -> bool:
        """Delete a message (moves to Deleted Items).

        Args:
            message_id: Message ID to delete.

        Returns:
            True if deleted successfully.
        """
        logger.debug("Deleting message: %s", message_id)
        response = await self.client.delete(
            f"/me/messages/{message_id}",
            headers=self._get_auth_header(),
        )

        if not response.is_success:
            self._handle_error(response)

        return True

    async def create_draft(
        self,
        subject: str | None = None,
        body: str | None = None,
        body_type: str = "text",
        to_recipients: list[str] | None = None,
        cc_recipients: list[str] | None = None,
        importance: str = "normal",
    ) -> Message:
        """Create a draft message.

        Args:
            subject: Message subject.
            body: Message body content.
            body_type: Body content type ("text" or "html").
            to_recipients: List of recipient email addresses.
            cc_recipients: List of CC recipient email addresses.
            importance: Message importance ("low", "normal", "high").

        Returns:
            The created draft message.
        """
        payload: dict[str, Any] = {}

        if subject:
            payload["subject"] = subject

        if body:
            payload["body"] = {
                "contentType": body_type,
                "content": body,
            }

        if to_recipients:
            payload["toRecipients"] = [
                {"emailAddress": {"address": email}} for email in to_recipients
            ]

        if cc_recipients:
            payload["ccRecipients"] = [
                {"emailAddress": {"address": email}} for email in cc_recipients
            ]

        if importance != "normal":
            payload["importance"] = importance

        logger.debug("Creating draft: %s", subject)
        response = await self.client.post(
            "/me/messages",
            headers=self._get_auth_header(),
            json=payload,
        )

        if not response.is_success:
            self._handle_error(response)

        return Message.model_validate(response.json())

    # =========================================================================
    # Calendar Operations
    # =========================================================================

    async def list_calendars(self) -> list[Calendar]:
        """List all calendars.

        Returns:
            List of calendars.
        """
        logger.debug("Listing calendars")
        response = await self.client.get(
            "/me/calendars",
            headers=self._get_auth_header(),
            params={"$select": CALENDAR_FIELDS},
        )

        if not response.is_success:
            self._handle_error(response)

        data = response.json()
        return [Calendar.model_validate(item) for item in data.get("value", [])]

    async def list_events(
        self,
        calendar_id: str | None = None,
        start_date: datetime | None = None,
        end_date: datetime | None = None,
        limit: int = 25,
        skip: int = 0,
    ) -> list[Event]:
        """List calendar events.

        Args:
            calendar_id: Calendar ID (default: primary calendar).
            start_date: Filter events starting after this date.
            end_date: Filter events ending before this date.
            limit: Maximum number of results.
            skip: Number of results to skip.

        Returns:
            List of events.
        """
        # Build endpoint
        endpoint = f"/me/calendars/{calendar_id}/events" if calendar_id else "/me/calendar/events"

        params: dict[str, Any] = {
            "$select": EVENT_FIELDS,
            "$top": limit,
            "$skip": skip,
            "$orderby": "start/dateTime",
        }

        # Build filter
        filters: list[str] = []

        if start_date:
            filters.append(f"start/dateTime ge '{start_date.isoformat()}'")

        if end_date:
            filters.append(f"end/dateTime le '{end_date.isoformat()}'")

        if filters:
            params["$filter"] = " and ".join(filters)

        logger.debug("Listing events from %s with params: %s", endpoint, params)
        response = await self.client.get(
            endpoint,
            headers=self._get_auth_header(),
            params=params,
        )

        if not response.is_success:
            self._handle_error(response)

        data = response.json()
        return [Event.model_validate(item) for item in data.get("value", [])]

    async def get_event(self, event_id: str) -> Event:
        """Get a single event by ID.

        Args:
            event_id: Event ID.

        Returns:
            The event details.
        """
        logger.debug("Getting event: %s", event_id)
        response = await self.client.get(
            f"/me/events/{event_id}",
            headers=self._get_auth_header(),
            params={"$select": EVENT_FIELDS},
        )

        if not response.is_success:
            self._handle_error(response)

        return Event.model_validate(response.json())

    async def get_free_busy(
        self,
        emails: list[str],
        start_time: datetime,
        end_time: datetime,
        timezone: str = "UTC",
        interval_minutes: int = 30,
    ) -> list[ScheduleInformation]:
        """Get free/busy schedule for one or more users.

        Args:
            emails: List of email addresses to check (up to 20).
            start_time: Start of time range to check.
            end_time: End of time range to check.
            timezone: Timezone for the query (default: UTC).
            interval_minutes: Granularity of availability view (default: 30).

        Returns:
            List of schedule information for each user.
        """
        if len(emails) > 20:
            raise GraphError("Maximum 20 email addresses allowed for getSchedule")

        payload = {
            "schedules": emails,
            "startTime": {
                "dateTime": start_time.strftime("%Y-%m-%dT%H:%M:%S"),
                "timeZone": timezone,
            },
            "endTime": {
                "dateTime": end_time.strftime("%Y-%m-%dT%H:%M:%S"),
                "timeZone": timezone,
            },
            "availabilityViewInterval": interval_minutes,
        }

        logger.debug("Getting free/busy for %s from %s to %s", emails, start_time, end_time)
        response = await self.client.post(
            "/me/calendar/getSchedule",
            headers=self._get_auth_header(),
            json=payload,
        )

        if not response.is_success:
            self._handle_error(response)

        data = response.json()
        return [ScheduleInformation.model_validate(item) for item in data.get("value", [])]
