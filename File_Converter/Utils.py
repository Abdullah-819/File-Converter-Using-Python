# utils.py
def sanitize_text(text):
    """Remove non-printable characters from text."""
    return ''.join(ch for ch in text if ch.isprintable() or ch in '\n\r\t')
