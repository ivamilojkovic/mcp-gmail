"""
Gmail MCP Server Implementation

This module provides a Model Context Protocol server for interacting with Gmail.
It exposes Gmail messages as resources and provides tools for composing and sending emails.
"""

import re
from datetime import datetime
from typing import Optional

from mcp.server.fastmcp import FastMCP

from mcp_gmail.config import settings
from mcp_gmail.gmail import (
    batch_modify_messages_labels,
    create_draft,
    get_gmail_service,
    get_headers_dict,
    get_labels,
    get_message,
    get_thread,
    modify_message_labels,
    parse_message_body,
    search_messages,
)
from mcp_gmail.gmail import send_email as gmail_send_email

# Initialize the Gmail service
service = get_gmail_service(
    credentials_path=settings.credentials_path, token_path=settings.token_path, scopes=settings.scopes
)

mcp = FastMCP(
    "Gmail MCP Server",
    instructions="Access and interact with Gmail. You can get messages, threads, search emails, and send or compose new messages.",  # noqa: E501
)

EMAIL_PREVIEW_LENGTH = 200


# Helper functions
def format_message(message):
    """Format a Gmail message for display."""
    headers = get_headers_dict(message)
    body = parse_message_body(message)

    # Extract relevant headers
    from_header = headers.get("From", "Unknown")
    to_header = headers.get("To", "Unknown")
    subject = headers.get("Subject", "No Subject")
    date = headers.get("Date", "Unknown Date")

    return f"""
From: {from_header}
To: {to_header}
Subject: {subject}
Date: {date}

{body}
"""


def validate_date_format(date_str):
    """
    Validate that a date string is in the format YYYY/MM/DD.

    Args:
        date_str: The date string to validate

    Returns:
        bool: True if valid, False otherwise
    """
    if not date_str:
        return True

    # Check format with regex
    if not re.match(r"^\d{4}/\d{2}/\d{2}$", date_str):
        return False

    # Validate the date is a real date
    try:
        datetime.strptime(date_str, "%Y/%m/%d")
        return True
    except ValueError:
        return False


# Resources
@mcp.resource("gmail://messages/{message_id}")
def get_email_message(message_id: str) -> str:
    """
    Get the content of an email message by its ID.

    Args:
        message_id: The Gmail message ID

    Returns:
        The formatted email content
    """
    message = get_message(service, message_id, user_id=settings.user_id)
    formatted_message = format_message(message)
    return formatted_message


@mcp.resource("gmail://threads/{thread_id}")
def get_email_thread(thread_id: str) -> str:
    """
    Get all messages in an email thread by thread ID.

    Args:
        thread_id: The Gmail thread ID

    Returns:
        The formatted thread content with all messages
    """
    thread = get_thread(service, thread_id, user_id=settings.user_id)
    messages = thread.get("messages", [])

    result = f"Email Thread (ID: {thread_id})\n"
    for i, message in enumerate(messages, 1):
        result += f"\n--- Message {i} ---\n"
        result += format_message(message)

    return result


# Tools
@mcp.tool()
def compose_email(
    to: str, subject: str, body: str, cc: Optional[str] = None, bcc: Optional[str] = None
) -> str:
    """
    Compose a new email draft.

    Args:
        to: Recipient email address
        subject: Email subject
        body: Email body content
        cc: Carbon copy recipients (optional)
        bcc: Blind carbon copy recipients (optional)

    Returns:
        The ID of the created draft and its content
    """
    sender = service.users().getProfile(userId=settings.user_id).execute().get("emailAddress")
    draft = create_draft(
        service, sender=sender, to=to, subject=subject, body=body, user_id=settings.user_id, cc=cc, bcc=bcc
    )

    draft_id = draft.get("id")
    return f"""
Email draft created with ID: {draft_id}
To: {to}
Subject: {subject}
CC: {cc or ""}
BCC: {bcc or ""}
Body: {body[:EMAIL_PREVIEW_LENGTH]}{"..." if len(body) > EMAIL_PREVIEW_LENGTH else ""}
"""


@mcp.tool()
def send_email(
    to: str, subject: str, body: str, cc: Optional[str] = None, bcc: Optional[str] = None
) -> str:
    """
    Compose and send an email.

    Args:
        to: Recipient email address
        subject: Email subject
        body: Email body content
        cc: Carbon copy recipients (optional)
        bcc: Blind carbon copy recipients (optional)

    Returns:
        Content of the sent email
    """
    sender = service.users().getProfile(userId=settings.user_id).execute().get("emailAddress")
    message = gmail_send_email(
        service, sender=sender, to=to, subject=subject, body=body, user_id=settings.user_id, cc=cc, bcc=bcc
    )

    message_id = message.get("id")
    return f"""
Email sent successfully with ID: {message_id}
To: {to}
Subject: {subject}
CC: {cc or ""}
BCC: {bcc or ""}
Body: {body[:EMAIL_PREVIEW_LENGTH]}{"..." if len(body) > EMAIL_PREVIEW_LENGTH else ""}
"""


@mcp.tool()
def search_emails(
    from_email: Optional[str] = None,
    to_email: Optional[str] = None,
    subject: Optional[str] = None,
    has_attachment: bool = False,
    is_unread: bool = False,
    after_date: Optional[str] = None,
    before_date: Optional[str] = None,
    label: Optional[str] = None,
    max_results: int = 10,
) -> str:
    """
    Search for emails using specific search criteria.

    Args:
        from_email: Filter by sender email
        to_email: Filter by recipient email
        subject: Filter by subject text
        has_attachment: Filter for emails with attachments
        is_unread: Filter for unread emails
        after_date: Filter for emails after this date (format: YYYY/MM/DD)
        before_date: Filter for emails before this date (format: YYYY/MM/DD)
        label: Filter by Gmail label
        max_results: Maximum number of results to return

    Returns:
        Formatted list of matching emails
    """
    # Validate date formats
    if after_date and not validate_date_format(after_date):
        return f"Error: after_date '{after_date}' is not in the required format YYYY/MM/DD"

    if before_date and not validate_date_format(before_date):
        return f"Error: before_date '{before_date}' is not in the required format YYYY/MM/DD"

    # Use search_messages to find matching emails
    messages = search_messages(
        service,
        user_id=settings.user_id,
        from_email=from_email,
        to_email=to_email,
        subject=subject,
        has_attachment=has_attachment,
        is_unread=is_unread,
        after=after_date,
        before=before_date,
        labels=[label] if label else None,
        max_results=max_results,
    )

    result = f"Found {len(messages)} messages matching criteria:\n"

    for msg_info in messages:
        msg_id = msg_info.get("id")
        message = get_message(service, msg_id, user_id=settings.user_id)
        headers = get_headers_dict(message)

        from_header = headers.get("From", "Unknown")
        subject = headers.get("Subject", "No Subject")
        date = headers.get("Date", "Unknown Date")

        result += f"\nMessage ID: {msg_id}\n"
        result += f"From: {from_header}\n"
        result += f"Subject: {subject}\n"
        result += f"Date: {date}\n"

    return result

@mcp.tool()
def search_unlabeled_emails(
    after_date: Optional[str] = None, 
    before_date: Optional[str] = None,
    max_results: int = 50
    ) -> str:
    """
    Search for emails that have NO user-created labels.
    System labels (INBOX, UNREAD, IMPORTANT etc.) are ignored.
    """
    # Validate dates
    if after_date and not validate_date_format(after_date):
        return f"Error: after_date '{after_date}' is not in YYYY/MM/DD format"
    if before_date and not validate_date_format(before_date):
        return f"Error: before_date '{before_date}' is not in YYYY/MM/DD format"

    # Fetch ALL messages (with optional date filtering)
    messages = search_messages(
        service,
        user_id=settings.user_id,
        after=after_date,
        before=before_date,
        max_results=max_results,
    )

    # Get all user-defined labels
    all_labels = get_labels(service, user_id=settings.user_id)
    user_label_ids = {label["id"] for label in all_labels if label.get("type") == "user"}

    result = "Messages without user labels:\n"
    count = 0

    for msg_info in messages:
        msg_id = msg_info["id"]

        # Fetch full message metadata
        message = get_message(service, msg_id, user_id=settings.user_id)
        labels = set(message.get("labelIds", []))

        # Check if message has zero user labels
        if labels.isdisjoint(user_label_ids):
            count += 1
            headers = get_headers_dict(message)
            result += (
                f"\nMessage ID: {msg_id}\n"
                f"From: {headers.get('From', 'Unknown')}\n"
                f"Subject: {headers.get('Subject', 'No Subject')}\n"
                f"Date: {headers.get('Date', 'Unknown')}\n"
            )

    return f"Found {count} unlabeled messages:\n" + result


@mcp.tool()
def get_available_labels() -> dict:
    """
    Get all available Gmail labels for the user.

    Returns:
        {
            "labels": [
                {
                    "id": str,
                    "name": str,
                    "type": str
                },
                ...
            ]
        }
    """
    labels = get_labels(service, user_id=settings.user_id)

    data = {
        "labels": []
    }

    for label in labels:
        data["labels"].append({
            "id": label.get("id", "Unknown"),
            "name": label.get("name", "Unknown"),
            "type": label.get("type", "user")
        })

    return data


@mcp.tool()
def add_label_to_message(message_id: str, label_id: str) -> str:
    """
    Add a label to a message.

    Args:
        message_id: The Gmail message ID
        label_id: The Gmail label ID to add (use list_available_labels to find label IDs)

    Returns:
        Confirmation message
    """
    # Add the specified label
    result = modify_message_labels(
        service, user_id=settings.user_id, message_id=message_id, remove_labels=[], add_labels=[label_id]
    )

    # Get message details to show what was modified
    headers = get_headers_dict(result)
    subject = headers.get("Subject", "No Subject")

    # Get the label name for the confirmation message
    label_name = label_id
    labels = get_labels(service, user_id=settings.user_id)
    for label in labels:
        if label.get("id") == label_id:
            label_name = label.get("name", label_id)
            break

    return f"""
Label added to message:
ID: {message_id}
Subject: {subject}
Added Label: {label_name} ({label_id})
"""


@mcp.tool()
def get_email_metadata(message_ids: list[str]) -> list[dict]:
    """
    Get structured header metadata for multiple emails.

    Returns From, Subject, Date, and unsubscribe-relevant headers only.
    Does NOT return email body or content.

    Args:
        message_ids: List of Gmail message IDs

    Returns:
        List of dicts with id, from, subject, date, list_unsubscribe, list_unsubscribe_post
    """
    results = []
    for msg_id in message_ids:
        try:
            message = get_message(service, msg_id, user_id=settings.user_id)
            headers = get_headers_dict(message)
            results.append({
                "id": msg_id,
                "from": headers.get("From", ""),
                "subject": headers.get("Subject", ""),
                "date": headers.get("Date", ""),
                "list_unsubscribe": headers.get("List-Unsubscribe", ""),
                "list_unsubscribe_post": headers.get("List-Unsubscribe-Post", ""),
            })
        except Exception as e:
            results.append({"id": msg_id, "error": str(e)})
    return results


@mcp.tool()
def categorize_emails_from_sender(
    sender_email: str,
    category_label_id: str = "CATEGORY_PROMOTIONS",
    max_results: int = 500,
) -> dict:
    """
    Move all emails from a sender into a Gmail category tab.

    Args:
        sender_email: The sender's email address to filter by
        category_label_id: Gmail system category label ID to apply.
            Valid values: CATEGORY_PROMOTIONS, CATEGORY_SOCIAL,
            CATEGORY_UPDATES, CATEGORY_FORUMS (default: CATEGORY_PROMOTIONS)
        max_results: Maximum number of messages to categorize (default: 500)

    Returns:
        {"categorized_count": int, "category_label_id": str}
    """
    messages = search_messages(
        service,
        user_id=settings.user_id,
        from_email=sender_email,
        max_results=max_results,
    )

    if not messages:
        return {"categorized_count": 0, "category_label_id": category_label_id}

    message_ids = [m["id"] for m in messages]
    batch_modify_messages_labels(
        service,
        user_id=settings.user_id,
        message_ids=message_ids,
        add_labels=[category_label_id],
    )

    return {"categorized_count": len(message_ids), "category_label_id": category_label_id}


if __name__ == "__main__":
    mcp.run(transport="stdio")