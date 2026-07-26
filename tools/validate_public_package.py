#!/usr/bin/env python3
"""Validate the distributable skill and reject private or non-portable assets."""

from __future__ import annotations

import json
import re
import subprocess
import sys
import zipfile
from pathlib import Path
from urllib.parse import unquote, urlsplit
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
SKILL = ROOT / "artifact-template-academic-word"
ASSETS = SKILL / "assets"
FORBIDDEN_PARTS = ("customXml/", "word/embeddings/", "word/comments", "word/people")
FORBIDDEN_BINARY_TEXT = ("461453258", "NE.Rep", "NoteExpress", "marong", "Jasper")
REQUIRED_AUTHORITY_PARTS = {
    "word/document.xml",
    "word/styles.xml",
    "word/numbering.xml",
    "word/theme/theme1.xml",
    "word/fontTable.xml",
    "word/webSettings.xml",
    "word/settings.xml",
    "word/footnotes.xml",
    "word/endnotes.xml",
}
MARKDOWN_FILES = (
    ROOT / "README.md",
    ROOT / "README_EN.md",
    ROOT / "CHANGELOG.md",
    ROOT / "THIRD_PARTY_NOTICES.md",
    ROOT / "CONTRIBUTING.md",
    ROOT / "SECURITY.md",
    SKILL / "SKILL.md",
    SKILL / "references" / "request-schema.md",
)


def fail(message: str) -> None:
    raise SystemExit(f"ERROR: {message}")


def check_skill() -> None:
    skill_md = SKILL / "SKILL.md"
    content = skill_md.read_text(encoding="utf-8")
    match = re.match(r"\A---\n(.*?)\n---", content, re.DOTALL)
    if not match or "name: artifact-template-academic-word" not in match.group(1):
        fail("SKILL.md frontmatter is invalid")
    for path in (SKILL / "artifact-template.json", SKILL / "references/style-contract.json"):
        json.loads(path.read_text(encoding="utf-8"))
    if json.loads((SKILL / "artifact-template.json").read_text(encoding="utf-8"))["reference"] != "assets/clean-report.docx":
        fail("artifact-template.json must reference the clean report master")
    if "assets/style-authority.dotx" not in content:
        fail("SKILL.md does not reference the sanitized authority")


def xml_text(data: bytes) -> str:
    try:
        return " ".join(text for text in ET.fromstring(data).itertext() if text)
    except ET.ParseError:
        return ""


def check_office_asset(path: Path, authority: bool = False) -> None:
    with zipfile.ZipFile(path) as package:
        names = set(package.namelist())
        for name in names:
            if name.startswith(FORBIDDEN_PARTS):
                fail(f"{path.name} contains forbidden part {name}")
        if authority and not REQUIRED_AUTHORITY_PARTS.issubset(names):
            fail(f"{path.name} lacks required authority parts")
        combined = []
        for name in names:
            if name.endswith((".xml", ".rels")):
                combined.append(package.read(name).decode("utf-8", errors="ignore"))
        text = "\n".join(combined)
        for marker in FORBIDDEN_BINARY_TEXT:
            if marker.casefold() in text.casefold():
                fail(f"{path.name} contains forbidden marker {marker}")
        document_text = xml_text(package.read("word/document.xml"))
        if authority and document_text.strip():
            fail("style-authority.dotx must have an empty document body")
        content_types = package.read("[Content_Types].xml").decode("utf-8")
        expected = "template.main+xml" if authority else "document.main+xml"
        if expected not in content_types:
            fail(f"{path.name} has the wrong main content type")


def check_documentation() -> None:
    for path in MARKDOWN_FILES:
        if not path.is_file():
            fail(f"required documentation is missing: {path.relative_to(ROOT)}")
        content = path.read_text(encoding="utf-8")
        for match in re.finditer(r"!?\[[^\]]*]\(([^)\s]+)", content):
            reference = match.group(1)
            parsed = urlsplit(reference)
            if parsed.scheme or reference.startswith(("#", "//", "mailto:", "tel:")):
                continue
            target = (path.parent / unquote(parsed.path)).resolve()
            if not target.is_relative_to(ROOT.resolve()):
                fail(f"{path.relative_to(ROOT)} links outside the package: {reference}")
            if not target.exists():
                fail(f"{path.relative_to(ROOT)} contains a broken local link: {reference}")

    if "[English](README_EN.md)" not in (ROOT / "README.md").read_text(encoding="utf-8"):
        fail("README.md must link to the English documentation")
    if "[中文](README.md)" not in (ROOT / "README_EN.md").read_text(encoding="utf-8"):
        fail("README_EN.md must link to the Chinese documentation")


def main() -> None:
    check_skill()
    check_documentation()
    check_office_asset(ASSETS / "style-authority.dotx", authority=True)
    check_office_asset(ASSETS / "clean-report.docx")
    check_office_asset(ASSETS / "clean-thesis.docx")
    distributable_paths = subprocess.run(
        ["git", "ls-files", "--cached", "--others", "--exclude-standard", "-z"],
        cwd=ROOT,
        check=True,
        capture_output=True,
    ).stdout.decode("utf-8").split("\0")
    forbidden_files = []
    for relative_path in filter(None, distributable_paths):
        path = ROOT / relative_path
        if not path.is_file():
            continue
        if path.name.startswith("~$") or path.suffix in {".pyc", ".pdf"} or "__pycache__" in path.parts:
            forbidden_files.append(relative_path)
    if forbidden_files:
        fail(f"forbidden distributable files: {forbidden_files}")
    print("Public package validation passed")


if __name__ == "__main__":
    main()
