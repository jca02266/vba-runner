#!/usr/bin/env python3
"""Expand git word-diff changes into vertically aligned old/new lines.

Usage examples:
    git diff --color=always --word-diff=color | git-word-diff-align.py
    git log -p --color=always --word-diff | git-word-diff-align.py

The filter deliberately only rewrites lines inside a diff hunk.  Headers,
commit messages, and other output from git are passed through unchanged.
"""

from __future__ import annotations

import re
import sys
from dataclasses import dataclass


ANSI = re.compile(r"\x1b\[[0-9;]*m")
SGR_RED = {31, 91}
SGR_GREEN = {32, 92}


@dataclass
class Chunk:
    text: str
    kind: str = "same"


def sgr_kind(sequence: str) -> str | None:
    """Return the word-diff side represented by an SGR sequence."""
    values = sequence[2:-1]
    codes = {int(value or 0) for value in values.split(";")}
    if codes & SGR_RED:
        return "old"
    if codes & SGR_GREEN:
        return "new"
    return None


def parse_color(line: str) -> list[Chunk]:
    """Split an ANSI-colored line into old, new, and unchanged chunks."""
    chunks: list[Chunk] = []
    position = 0
    kind = "same"
    for match in ANSI.finditer(line):
        if match.start() > position:
            chunks.append(Chunk(line[position : match.start()], kind))
        detected = sgr_kind(match.group())
        if detected:
            kind = detected
        elif match.group() == "\x1b[m" or match.group() == "\x1b[0m":
            kind = "same"
        chunks.append(Chunk(match.group(), "control"))
        position = match.end()
    if position < len(line):
        chunks.append(Chunk(line[position:], kind))
    return chunks


def parse_markers(line: str) -> list[Chunk]:
    """Split --word-diff[=plain] markers without interpreting their contents."""
    chunks: list[Chunk] = []
    position = 0
    marker = re.compile(r"\[-(.*?)-\]|\{\+(.*?)\+\}")
    for match in marker.finditer(line):
        if match.start() > position:
            chunks.append(Chunk(line[position : match.start()]))
        if match.group(1) is not None:
            chunks.append(Chunk(match.group(1), "old"))
        else:
            chunks.append(Chunk(match.group(2), "new"))
        position = match.end()
    if position < len(line):
        chunks.append(Chunk(line[position:]))
    return chunks


def remove_markers(line: str) -> str:
    """Remove plain word-diff delimiters before interpreting ANSI colors."""
    line = re.sub(r"\[-(.*?)-\]", r"\1", line)
    return re.sub(r"\{\+(.*?)\+\}", r"\1", line)


def has_color_change(chunks: list[Chunk]) -> bool:
    return any(chunk.kind in {"old", "new"} for chunk in chunks)


def render(chunks: list[Chunk], side: str) -> str:
    """Render one side while retaining all ANSI controls from the input."""
    return "".join(
        chunk.text
        for chunk in chunks
        if chunk.kind == "control" or chunk.kind == "same" or chunk.kind == side
    )


def transform_line(line: str) -> list[str]:
    newline = "\n" if line.endswith("\n") else ""
    body = line[:-1] if newline else line
    chunks = parse_color(remove_markers(body))
    if not has_color_change(chunks):
        chunks = parse_markers(body)
    if not has_color_change(chunks):
        return [line]

    old = render(chunks, "old")
    new = render(chunks, "new")
    # A missing final newline still needs a separator between the generated
    # old and new lines.
    return [f"- {old}\n", f"+ {new}{newline}"]


def main() -> int:
    in_hunk = False
    for line in sys.stdin:
        plain = ANSI.sub("", line)
        if plain.startswith("@@"):
            in_hunk = True
            sys.stdout.write(line)
            continue
        if in_hunk and (plain.startswith("diff --git ") or plain.startswith("commit ")):
            in_hunk = False
        if in_hunk:
            sys.stdout.writelines(transform_line(line))
        else:
            sys.stdout.write(line)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
