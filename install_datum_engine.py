"""Copy datum-engine.js into the web app, unchanged.

The engine is a clean ES module with a single export block at the bottom, so it
drops in as-is -- no edits. Copying it with a script rather than retyping it
guarantees the 957 lines arrive byte-identical.

USAGE
    python install_datum_engine.py "C:\\path\\to\\datum-engine.js"

    # if you dropped it somewhere obvious, it will find it itself:
    python install_datum_engine.py

Destination:
    C:\\CMS_AI\\webapp\\frontend\\src\\lib\\datumEngine.js
"""

import shutil
import sys
from pathlib import Path

DEST = Path(r"C:\CMS_AI\webapp\frontend\src\lib\datumEngine.js")

# Exact locations first, cheapest to check.
SEARCH = [
    Path(r"C:\CMS_AI\datum-engine.js"),
    Path(r"C:\CMS_Local_Workspace\datum-engine.js"),
    Path.home() / "Downloads" / "datum-engine.js",
    Path.home() / "Desktop" / "datum-engine.js",
    Path.home() / "Documents" / "datum-engine.js",
]

# Then sweep the Claude desktop session folders, where an uploaded file actually
# lands on disk. Session ids are random, so this has to be a glob.
GLOB_ROOTS = [
    Path.home() / "AppData" / "Roaming" / "Claude" / "local-agent-mode-sessions",
    Path.home() / "Downloads",
]


def find_engine() -> Path | None:
    for cand in SEARCH:
        if cand.is_file():
            return cand

    for root in GLOB_ROOTS:
        if not root.is_dir():
            continue
        try:
            # Newest first: if the file was uploaded more than once, take the
            # most recent copy.
            hits = sorted(
                root.rglob("datum-engine.js"),
                key=lambda p: p.stat().st_mtime,
                reverse=True,
            )
        except OSError:
            continue
        for hit in hits:
            if hit.is_file():
                return hit
    return None


def main() -> None:
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    quiet = "--quiet" in sys.argv[1:]

    # Already installed and non-trivial? Leave it alone. This lets the launcher
    # call us on every start without doing any work.
    if "--if-missing" in sys.argv[1:] and DEST.is_file() and DEST.stat().st_size > 10_000:
        if not quiet:
            print(f"Datum engine already installed: {DEST}")
        return

    src = None
    if args:
        src = Path(args[0])
        if not src.is_file():
            raise SystemExit(f"Not a file: {src}")
    else:
        src = find_engine()
        if src is None:
            msg = (
                "Could not find datum-engine.js anywhere.\n"
                "Drop it in C:\\CMS_AI\\ and start again, or pass the path:\n"
                '  python install_datum_engine.py "C:\\path\\to\\datum-engine.js"'
            )
            if "--if-missing" in sys.argv[1:]:
                # Non-fatal on startup: the app still runs, the Machining tab
                # just falls back to the server's coarse estimate.
                print("WARNING: " + msg)
                print("The Machining tab will use the pattern estimate until this is fixed.")
                return
            raise SystemExit(msg)

    if src.resolve() == DEST.resolve():
        if not quiet:
            print("Source and destination are the same file. Nothing to do.")
        return

    text = src.read_text(encoding="utf-8", errors="strict")

    # Sanity-check it is the engine and not the console demo, which inlines a
    # copy of the engine and would compile but drag in React and three.js.
    if "export {" not in text or "parseSTL" not in text:
        raise SystemExit(
            f"{src} does not look like datum-engine.js "
            "(no export block / no parseSTL). Nothing written."
        )
    if "import React" in text or "from \"react\"" in text:
        raise SystemExit(
            f"{src} looks like datum-console.jsx, not the engine. "
            "Use datum-engine.js -- the console inlines its own copy and pulls "
            "in React + three.js. Nothing written."
        )

    DEST.parent.mkdir(parents=True, exist_ok=True)
    if DEST.exists():
        backup = DEST.with_suffix(".js.bak")
        shutil.copy2(DEST, backup)
        print(f"Existing file backed up to {backup}")

    DEST.write_text(text, encoding="utf-8")

    print(f"Copied {src}")
    print(f"    -> {DEST}")
    print(f"       {len(text):,} chars, {text.count(chr(10)):,} lines")
    print()
    print("The adapter at src/lib/datumAdapter.ts imports it. Nothing else to do.")


if __name__ == "__main__":
    main()
