from datetime import datetime
import re


def normalize_name(name: str) -> str:
    """Return a canonical representation of a person name.

    Removes extra whitespace, punctuation (except hyphen), converts to
    uppercase so matches are case-insensitive.
    """
    if not isinstance(name, str):
        return ""
    # Replace multiple spaces, strip, remove common punctuation
    cleaned = re.sub(r"[\.,']", "", name).strip().upper()
    # Collapse consecutive whitespace to single space
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned


def format_output_filename(prefix: str, extension: str = "xlsx", timestamp: str | None = None) -> str:
    """Return a consistent timestamped output filename.

    Example: format_output_filename("master_names_with_ids") ->
    'master_names_with_ids_20250804T093045.xlsx'
    """
    if timestamp is None:
        timestamp = datetime.now().strftime("%Y%m%dT%H%M%S")
    # Ensure prefix is filesystem-friendly (spaces -> underscore)
    safe_prefix = re.sub(r"[^A-Za-z0-9_-]+", "_", prefix.strip()).lower()
    return f"{safe_prefix}_{timestamp}.{extension.lstrip('.')}"