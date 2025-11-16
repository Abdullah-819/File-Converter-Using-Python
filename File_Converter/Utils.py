def sanitize(text):
    return ''.join(ch for ch in text if ch.isprintable() or ch in "\n\r\t")
