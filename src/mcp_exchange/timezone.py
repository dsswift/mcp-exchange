"""Timezone service for consistent date/time handling."""

from __future__ import annotations

from datetime import datetime
from typing import TYPE_CHECKING
from zoneinfo import ZoneInfo

if TYPE_CHECKING:
    from .models import DateTimeTimeZone


class TimezoneService:
    """Centralized service for timezone-aware date/time operations.

    All datetime formatting and parsing should go through this service
    to ensure consistent timezone handling across the codebase.
    """

    def __init__(self, timezone: str = "America/Chicago") -> None:
        """Initialize the timezone service.

        Args:
            timezone: IANA timezone name (e.g., "America/Chicago", "America/New_York").
        """
        self.tz = ZoneInfo(timezone)
        self.timezone_name = timezone

    def format_datetime(self, dt: datetime, source_tz: str | None = None) -> str:
        """Format a datetime in the user's timezone.

        Args:
            dt: Datetime to format. If naive, assumed to be in source_tz or UTC.
            source_tz: IANA timezone name of the source datetime. Defaults to UTC.

        Returns:
            Human-readable datetime string in user's timezone (e.g., "2026-01-12 10:30 AM CST").
        """
        if dt.tzinfo is None:
            # Naive datetime - assume it's in the source timezone
            source = ZoneInfo(source_tz) if source_tz else ZoneInfo("UTC")
            dt = dt.replace(tzinfo=source)

        # Convert to user's timezone
        local_dt = dt.astimezone(self.tz)
        return local_dt.strftime("%Y-%m-%d %I:%M %p %Z")

    def format_graph_datetime(self, dt_tz: DateTimeTimeZone) -> str:
        """Format a Graph API DateTimeTimeZone object in the user's timezone.

        Args:
            dt_tz: DateTimeTimeZone object from Graph API response.

        Returns:
            Human-readable datetime string in user's timezone.
        """
        # Parse the datetime string (Graph API format: "2026-01-12T16:30:00.0000000")
        dt_str = dt_tz.date_time.split(".")[0]  # Remove fractional seconds
        dt = datetime.fromisoformat(dt_str)

        # Apply the source timezone from the Graph API
        source_tz = dt_tz.time_zone
        return self.format_datetime(dt, source_tz)

    def format_date(self, dt: datetime, source_tz: str | None = None) -> str:
        """Format just the date portion in the user's timezone.

        Args:
            dt: Datetime to format.
            source_tz: IANA timezone name of the source datetime.

        Returns:
            Date string in user's timezone (e.g., "2026-01-12").
        """
        if dt.tzinfo is None:
            source = ZoneInfo(source_tz) if source_tz else ZoneInfo("UTC")
            dt = dt.replace(tzinfo=source)

        local_dt = dt.astimezone(self.tz)
        return local_dt.strftime("%Y-%m-%d")

    def format_time(self, dt: datetime, source_tz: str | None = None) -> str:
        """Format just the time portion in the user's timezone.

        Args:
            dt: Datetime to format.
            source_tz: IANA timezone name of the source datetime.

        Returns:
            Time string in user's timezone (e.g., "10:30 AM CST").
        """
        if dt.tzinfo is None:
            source = ZoneInfo(source_tz) if source_tz else ZoneInfo("UTC")
            dt = dt.replace(tzinfo=source)

        local_dt = dt.astimezone(self.tz)
        return local_dt.strftime("%I:%M %p %Z")

    def parse_date(self, date_str: str) -> datetime:
        """Parse a date string as start of day in user's timezone.

        Args:
            date_str: Date string in YYYY-MM-DD format.

        Returns:
            Timezone-aware datetime at start of day in user's timezone.
        """
        dt = datetime.fromisoformat(date_str)
        # Make it timezone-aware in user's timezone
        return dt.replace(tzinfo=self.tz)

    def parse_datetime(self, dt_str: str) -> datetime:
        """Parse a datetime string in user's timezone.

        Args:
            dt_str: Datetime string in ISO format (YYYY-MM-DDTHH:MM:SS).

        Returns:
            Timezone-aware datetime in user's timezone.
        """
        dt = datetime.fromisoformat(dt_str)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=self.tz)
        return dt

    def get_day_bounds(self, date: datetime) -> tuple[datetime, datetime]:
        """Get start and end of day in user's timezone.

        Useful for filtering events/emails for a specific day.

        Args:
            date: A datetime representing the day (time portion ignored).

        Returns:
            Tuple of (start_of_day, end_of_day) as timezone-aware datetimes.
        """
        # Ensure we're working in user's timezone
        if date.tzinfo is None:
            date = date.replace(tzinfo=self.tz)
        else:
            date = date.astimezone(self.tz)

        start = date.replace(hour=0, minute=0, second=0, microsecond=0)
        end = date.replace(hour=23, minute=59, second=59, microsecond=999999)
        return start, end

    def to_utc(self, dt: datetime) -> datetime:
        """Convert a datetime to UTC.

        Args:
            dt: Datetime to convert. If naive, assumed to be in user's timezone.

        Returns:
            UTC datetime.
        """
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=self.tz)
        return dt.astimezone(ZoneInfo("UTC"))

    def to_utc_isoformat(self, dt: datetime) -> str:
        """Convert a datetime to UTC ISO format string.

        Args:
            dt: Datetime to convert.

        Returns:
            ISO format string in UTC (e.g., "2026-01-12T16:30:00Z").
        """
        utc_dt = self.to_utc(dt)
        return utc_dt.strftime("%Y-%m-%dT%H:%M:%SZ")
