"""Configuration module for Exchange Online MCP Server."""

from __future__ import annotations

import os
import sys
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv


@dataclass
class ExchangeConfig:
    """Configuration for Microsoft Graph API connection."""

    client_id: str
    tenant_id: str = "common"
    token_cache_path: Path = Path.home() / ".mcp-exchange" / "token_cache.json"
    timeout: int = 30
    timezone: str = "America/Chicago"

    # Microsoft Graph API scopes for Exchange Online
    scopes: tuple[str, ...] = (
        "User.Read",
        "Mail.ReadWrite",
        "Mail.Send",
        "Calendars.Read",
        "Calendars.Read.Shared",
    )

    @property
    def authority(self) -> str:
        """Get the Azure AD authority URL."""
        return f"https://login.microsoftonline.com/{self.tenant_id}"

    @property
    def graph_base_url(self) -> str:
        """Get the Microsoft Graph API base URL."""
        return "https://graph.microsoft.com/v1.0"


def load_config() -> ExchangeConfig:
    """Load configuration from environment variables.

    Returns:
        ExchangeConfig: Validated configuration object.

    Raises:
        SystemExit: If required environment variables are missing.
    """
    load_dotenv()

    client_id = os.getenv("EXCHANGE_CLIENT_ID")

    if not client_id:
        print("Configuration error: EXCHANGE_CLIENT_ID environment variable is required",
              file=sys.stderr)
        print("\nSee .env.sample for required configuration.", file=sys.stderr)
        sys.exit(1)

    tenant_id = os.getenv("EXCHANGE_TENANT_ID", "common")

    # Token cache path
    token_cache_str = os.getenv("EXCHANGE_TOKEN_CACHE")
    if token_cache_str:
        token_cache_path = Path(token_cache_str).expanduser()
    else:
        token_cache_path = Path.home() / ".mcp-exchange" / "token_cache.json"

    # Timeout
    timeout_str = os.getenv("EXCHANGE_TIMEOUT", "30")
    try:
        timeout = int(timeout_str)
    except ValueError:
        print(
            f"Warning: Invalid EXCHANGE_TIMEOUT value '{timeout_str}', using default 30",
            file=sys.stderr,
        )
        timeout = 30

    # Timezone (IANA format, e.g., "America/Chicago", "America/New_York")
    timezone = os.getenv("EXCHANGE_TIMEZONE", "America/Chicago")

    return ExchangeConfig(
        client_id=client_id,
        tenant_id=tenant_id,
        token_cache_path=token_cache_path,
        timeout=timeout,
        timezone=timezone,
    )
