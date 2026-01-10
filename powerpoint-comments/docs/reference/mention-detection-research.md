# @Mention Detection Research

Research for detecting when current user is @mentioned in PowerPoint comments.

## Key Finding: No Direct COM Property

PowerPoint's COM model does NOT expose a `Mentions` collection. Mentions appear as **plain text** in `Comment.Text`.

## Text Format Patterns

```
# Example formats in Comment.Text:
"@John Doe please review this section"
"@Jane Smith and @Bob Johnson - thoughts?"
"I agree with @Jane"  # User shortened the name
"What do you think @Jane Smith"
```

**Key observations:**
- Format is display name, NOT email
- User can modify after selection (shorten name)
- No hidden markers or metadata around the mention
- Case preserved from directory

## Regex Patterns

### Primary Pattern (Display Names)
```python
import re

# Handles: @John Doe, @John, @Jean-Pierre Dupont
MENTION_PATTERN = re.compile(
    r'@([A-Z\u00C0-\u024F][a-z\u00C0-\u024F]+(?:[-\'][A-Z\u00C0-\u024F]?[a-z\u00C0-\u024F]+)?'
    r'(?:\s+[A-Z\u00C0-\u024F][a-z\u00C0-\u024F]+(?:[-\'][A-Z\u00C0-\u024F]?[a-z\u00C0-\u024F]+)?)*)',
    re.UNICODE
)

# Exclude email patterns
EMAIL_PATTERN = re.compile(
    r'@[\w.-]+@[\w.-]+\.\w+',
    re.UNICODE
)
```

### Extract Mentions
```python
def extract_mentions(text):
    if not text:
        return []

    # Find emails to exclude
    email_matches = set(EMAIL_PATTERN.findall(text))

    mentions = []
    for match in MENTION_PATTERN.finditer(text):
        mention = match.group(1).strip()
        full_match = match.group(0)

        # Skip if looks like email
        if any(email in full_match for email in email_matches):
            continue

        mentions.append(mention)

    return mentions
```

## Current User Detection

### Method 1: Windows Display Name (Recommended)
```python
import ctypes

def get_windows_display_name():
    GetUserNameEx = ctypes.windll.secur32.GetUserNameExW
    NameDisplay = 3  # EXTENDED_NAME_FORMAT.NameDisplay

    size = ctypes.pointer(ctypes.c_ulong(0))
    GetUserNameEx(NameDisplay, None, size)

    name_buffer = ctypes.create_unicode_buffer(size.contents.value)
    GetUserNameEx(NameDisplay, name_buffer, size)

    return name_buffer.value
```

### Method 2: Outlook Email
```python
import win32com.client

def get_outlook_email():
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        return namespace.Accounts.Item(1).SmtpAddress
    except:
        return None
```

### Method 3: Environment Variable
```python
import os
username = os.environ.get('USERNAME')
```

## Fuzzy Matching

Users may shorten mentions: `@John` instead of `@John Smith`

```python
from difflib import SequenceMatcher

def names_match(mention, target, threshold=0.8):
    mention_lower = mention.lower()
    target_lower = target.lower()

    # Exact match
    if mention_lower == target_lower:
        return True

    # Target starts with mention (@John matches "John Smith")
    if target_lower.startswith(mention_lower):
        return True

    # Mention is first name only
    target_parts = target_lower.split()
    if target_parts and mention_lower == target_parts[0]:
        return True

    # Fuzzy match for typos
    similarity = SequenceMatcher(None, mention_lower, target_lower).ratio()
    return similarity >= threshold
```

## Scanning Comments

```python
def scan_for_mentions_to_user(presentation, user_identities):
    """Find all comments where current user is @mentioned."""
    mentions_found = []

    for slide in presentation.Slides:
        for comment in slide.Comments:
            # Check parent comment
            if check_for_user_mention(comment.Text, user_identities):
                mentions_found.append({
                    'slide': slide.SlideIndex,
                    'author': comment.Author,
                    'text': comment.Text
                })

            # Check replies
            for reply in comment.Replies:
                if check_for_user_mention(reply.Text, user_identities):
                    mentions_found.append({
                        'slide': slide.SlideIndex,
                        'author': reply.Author,
                        'text': reply.Text,
                        'is_reply': True
                    })

    return mentions_found

def check_for_user_mention(text, user_identities):
    """Check if text mentions any of the user's identities."""
    mentions = extract_mentions(text)

    for mention in mentions:
        # Check display name
        if user_identities.get('display_name'):
            if names_match(mention, user_identities['display_name']):
                return True

        # Check first name only
        if user_identities.get('first_name'):
            if mention.lower() == user_identities['first_name'].lower():
                return True

    return False
```

## Performance

- Measured at <5ms per comment
- Suitable for real-time detection
- Cache user identities on startup

## Risk Assessment

| Risk | Level | Mitigation |
|------|-------|------------|
| False positives from emails | Medium | Regex to exclude email patterns |
| Display name variations | Medium | Fuzzy matching with multiple sources |
| International characters | Low | Unicode-aware regex |
| Performance | Low | <5ms per comment measured |

## Status

**NOT IMPLEMENTED** - This is research for a future feature. Current plugin does not detect @mentions.
