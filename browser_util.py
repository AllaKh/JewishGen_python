"""One place to launch the browser for ALL scrapers.

Prefer the REAL installed Google Chrome (channel="chrome"); if it isn't present, fall back to
the bundled Playwright Chromium. This way every database in the app behaves the same — real
Chrome when the user has it, and a self-contained Chromium in the packaged install otherwise
(so a distributed copy still works on a machine without Chrome).
"""


async def launch_persistent(pw, user_data_dir, **kw):
    """launch_persistent_context, real Chrome first, Chromium fallback."""
    last = None
    for channel in ("chrome", None):
        try:
            k = dict(kw)
            if channel:
                k["channel"] = channel
            return await pw.chromium.launch_persistent_context(str(user_data_dir), **k)
        except Exception as e:
            last = e
    raise last


async def launch(pw, **kw):
    """chromium.launch (non-persistent), real Chrome first, Chromium fallback."""
    last = None
    for channel in ("chrome", None):
        try:
            k = dict(kw)
            if channel:
                k["channel"] = channel
            return await pw.chromium.launch(**k)
        except Exception as e:
            last = e
    raise last
