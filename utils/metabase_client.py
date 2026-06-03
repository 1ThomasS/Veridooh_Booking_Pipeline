"""
utils/metabase_client.py
------------------------
Thin wrapper around Veridooh's Metabase query API.

Configure METABASE_URL and METABASE_TOKEN via environment variables (or
pass them directly to MetabaseClient.__init__). Everything else is auto-handled.

Usage:
    from utils.metabase_client import MetabaseClient

    client = MetabaseClient()
    fmt = client.get_panel_format("P12345")          # -> "Classic Portrait"
    dims = client.get_panel_dimensions("P12345")     # -> "1080x1920"
    row  = client.get_panel_full("P12345")           # -> dict of all panel fields
"""

import os
import requests
from functools import lru_cache


# ---------------------------------------------------------------------------
# Configuration — fill these in or set environment variables
# ---------------------------------------------------------------------------

DEFAULT_URL   = os.getenv("METABASE_URL",   "https://metabase.veridooh.com")
DEFAULT_TOKEN = os.getenv("METABASE_TOKEN", "")

# The Metabase question (saved query) ID that exposes the Master Site List.
# Run the query manually once, grab the ID from the URL (/question/<ID>), set here.
# TODO: confirm this ID with your team
PANEL_LOOKUP_QUESTION_ID = int(os.getenv("METABASE_QUESTION_ID", "0"))

# Column names as they appear in your Metabase panel lookup query.
# TODO: update these to match your actual Metabase column headers
COL_PANEL_ID  = "Panel ID"
COL_FORMAT    = "Format"
COL_WIDTH     = "Width"
COL_HEIGHT    = "Height"
COL_DIMENSION = "Screen Resolution"   # if width/height are combined in one column


class MetabaseClient:
    """
    Query the Metabase Master Site List for panel metadata.

    Parameters
    ----------
    url   : str  Metabase base URL (default: METABASE_URL env var)
    token : str  Session token or API key (default: METABASE_TOKEN env var)
    """

    def __init__(self, url: str = DEFAULT_URL, token: str = DEFAULT_TOKEN):
        if not token:
            raise ValueError(
                "Metabase token not set. "
                "Export METABASE_TOKEN=<your-token> or pass token= directly."
            )
        self.base_url = url.rstrip("/")
        self.session  = requests.Session()
        self.session.headers.update({
            "X-Metabase-Session": token,
            "Content-Type": "application/json",
        })
        self._panel_cache: dict[str, dict] = {}

    # ------------------------------------------------------------------
    # Public interface
    # ------------------------------------------------------------------

    def get_panel_format(self, panel_id: str) -> str | None:
        """Return the canonical Format string for a panel, or None if not found."""
        row = self._get_row(panel_id)
        return row.get(COL_FORMAT) if row else None

    def get_panel_dimensions(self, panel_id: str) -> str | None:
        """
        Return the canonical WxH dimension string for a panel, e.g. '1080x1920'.
        Returns None if the panel is not found in Metabase.
        """
        row = self._get_row(panel_id)
        if not row:
            return None

        if COL_DIMENSION in row and row[COL_DIMENSION]:
            return _normalise_dim(str(row[COL_DIMENSION]))

        if COL_WIDTH in row and COL_HEIGHT in row:
            w, h = row[COL_WIDTH], row[COL_HEIGHT]
            if w and h:
                return f"{int(w)}x{int(h)}"

        return None

    def get_panel_full(self, panel_id: str) -> dict | None:
        """Return the full Metabase row dict for a panel, or None."""
        return self._get_row(panel_id)

    def prefetch(self, panel_ids: list[str]) -> None:
        """
        Bulk-load all rows for a list of Panel IDs in a single query.
        Call this once at the start of a run to avoid per-row round trips.

        TODO: implement once you confirm the Metabase question accepts a
        filter parameter. For now falls back to per-row fetching.
        """
        # Placeholder — swap for a bulk API call once you know the question schema
        for pid in panel_ids:
            self._get_row(pid)

    # ------------------------------------------------------------------
    # Internal
    # ------------------------------------------------------------------

    def _get_row(self, panel_id: str) -> dict | None:
        """Fetch a single panel row, using the in-memory cache."""
        pid = panel_id.strip()
        if pid in self._panel_cache:
            return self._panel_cache[pid]

        row = self._fetch_from_metabase(pid)
        self._panel_cache[pid] = row    # cache even if None to avoid re-fetching
        return row

    def _fetch_from_metabase(self, panel_id: str) -> dict | None:
        """
        Execute the Metabase question with a Panel ID filter.

        TODO: once you know whether your question uses a template variable
        or a parameter filter, update the payload below accordingly.

        Metabase REST endpoint for a saved question:
            POST /api/card/<question_id>/query
        with body:
            {"parameters": [{"type": "category", "target": [...], "value": "<panel_id>"}]}
        """
        if PANEL_LOOKUP_QUESTION_ID == 0:
            raise RuntimeError(
                "METABASE_QUESTION_ID not set. "
                "Find the question ID from the Metabase URL and set the env var."
            )

        url = f"{self.base_url}/api/card/{PANEL_LOOKUP_QUESTION_ID}/query"
        payload = {
            "parameters": [
                {
                    "type":   "category",
                    "target": ["variable", ["template-tag", "panel_id"]],
                    "value":  panel_id,
                }
            ]
        }

        try:
            resp = self.session.post(url, json=payload, timeout=10)
            resp.raise_for_status()
            data = resp.json()

            rows = data.get("data", {}).get("rows", [])
            cols = [c["name"] for c in data.get("data", {}).get("cols", [])]

            if not rows:
                return None

            # Return the first matching row as a dict
            return dict(zip(cols, rows[0]))

        except requests.RequestException as exc:
            print(f"[MetabaseClient] Warning: could not fetch panel {panel_id!r}: {exc}")
            return None


# ------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------

def _normalise_dim(raw: str) -> str:
    """Normalise a raw dimension string to 'WxH' with no spaces or case variation."""
    cleaned = raw.strip().lower().replace(" ", "").replace("px", "")
    # Support 'x' or 'X' separator
    if "x" in cleaned:
        parts = cleaned.split("x", 1)
        try:
            w, h = int(parts[0]), int(parts[1])
            return f"{w}x{h}"
        except ValueError:
            pass
    return raw.strip()
