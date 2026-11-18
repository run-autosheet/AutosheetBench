"""Claude's visible reply, used when its transcript export is unavailable."""


def visible_reply(capture, prompt):
    """Accept only text copied with Claude's own "Copy message" control.

    Accessibility trees may mix prompts, tools and the composer. They are kept as
    diagnostics but never used as the assistant's reply. Completion is never
    inferred from this text.
    """
    if capture.get("status") != "ok" or capture.get("source") != "claude-copy-message":
        return None
    normalize = lambda value: " ".join(value.split())
    text = str(capture.get("text") or "").strip()
    return text if text and normalize(text) != normalize(prompt) else None
