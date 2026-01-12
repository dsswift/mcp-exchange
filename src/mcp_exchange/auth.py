"""Authentication module for Microsoft Graph API using MSAL."""

from __future__ import annotations

import json
import logging
import sys
from typing import TYPE_CHECKING

import msal

if TYPE_CHECKING:
    from .config import ExchangeConfig

logger = logging.getLogger(__name__)


class AuthError(Exception):
    """Authentication error."""


class GraphAuthenticator:
    """Handles Microsoft Graph API authentication using MSAL device code flow."""

    def __init__(self, config: ExchangeConfig) -> None:
        """Initialize the authenticator.

        Args:
            config: Exchange configuration with client_id and tenant_id.
        """
        self.config = config
        self._app: msal.PublicClientApplication | None = None
        self._token_cache: msal.SerializableTokenCache | None = None

    def _ensure_cache_dir(self) -> None:
        """Ensure the token cache directory exists."""
        cache_dir = self.config.token_cache_path.parent
        cache_dir.mkdir(parents=True, exist_ok=True)

    def _load_token_cache(self) -> msal.SerializableTokenCache:
        """Load or create the token cache.

        Returns:
            Token cache instance, loaded from disk if available.
        """
        cache = msal.SerializableTokenCache()

        if self.config.token_cache_path.exists():
            try:
                cache_data = self.config.token_cache_path.read_text()
                cache.deserialize(cache_data)
                logger.debug("Loaded token cache from %s", self.config.token_cache_path)
            except (json.JSONDecodeError, OSError) as e:
                logger.warning("Failed to load token cache: %s", e)

        return cache

    def _save_token_cache(self) -> None:
        """Save the token cache to disk if it has changed."""
        if self._token_cache and self._token_cache.has_state_changed:
            self._ensure_cache_dir()
            try:
                self.config.token_cache_path.write_text(self._token_cache.serialize())
                logger.debug("Saved token cache to %s", self.config.token_cache_path)
            except OSError as e:
                logger.warning("Failed to save token cache: %s", e)

    def _get_app(self) -> msal.PublicClientApplication:
        """Get or create the MSAL application instance.

        Returns:
            MSAL PublicClientApplication configured with token cache.
        """
        if self._app is None:
            self._token_cache = self._load_token_cache()
            self._app = msal.PublicClientApplication(
                client_id=self.config.client_id,
                authority=self.config.authority,
                token_cache=self._token_cache,
            )
        return self._app

    def get_access_token(self) -> str:
        """Get an access token, using cache or device code flow.

        First attempts to acquire a token silently from cache.
        If that fails, initiates device code flow for user authentication.

        Returns:
            Access token string.

        Raises:
            AuthError: If authentication fails.
        """
        app = self._get_app()
        scopes = list(self.config.scopes)

        # Try to get token silently from cache
        accounts = app.get_accounts()
        if accounts:
            logger.debug("Found %d cached account(s), attempting silent auth", len(accounts))
            result = app.acquire_token_silent(scopes=scopes, account=accounts[0])
            if result and "access_token" in result:
                logger.info("Acquired token silently from cache")
                self._save_token_cache()
                return result["access_token"]

        # Fall back to device code flow
        logger.info("No cached token available, initiating device code flow")
        flow = app.initiate_device_flow(scopes=scopes)

        if "user_code" not in flow:
            error_msg = flow.get("error_description", "Failed to initiate device code flow")
            raise AuthError(error_msg)

        # Print device code instructions to stderr (stdout is for MCP protocol)
        print("\n" + "=" * 60, file=sys.stderr)
        print("AUTHENTICATION REQUIRED", file=sys.stderr)
        print("=" * 60, file=sys.stderr)
        print(f"\nTo sign in, visit: {flow['verification_uri']}", file=sys.stderr)
        print(f"Enter this code: {flow['user_code']}", file=sys.stderr)
        print("\nWaiting for authentication...", file=sys.stderr)

        result = app.acquire_token_by_device_flow(flow)

        if "access_token" not in result:
            error_msg = result.get("error_description", "Authentication failed")
            raise AuthError(error_msg)

        logger.info("Successfully authenticated via device code flow")
        self._save_token_cache()

        return result["access_token"]

    def get_auth_header(self) -> dict[str, str]:
        """Get the authorization header for API requests.

        Returns:
            Dictionary with Authorization header.
        """
        token = self.get_access_token()
        return {"Authorization": f"Bearer {token}"}

    def clear_cache(self) -> None:
        """Clear the token cache (logout)."""
        if self.config.token_cache_path.exists():
            try:
                self.config.token_cache_path.unlink()
                logger.info("Cleared token cache")
            except OSError as e:
                logger.warning("Failed to clear token cache: %s", e)

        self._app = None
        self._token_cache = None
