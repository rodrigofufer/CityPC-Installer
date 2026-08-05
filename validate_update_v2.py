#!/usr/bin/env python3
"""Static, secret-safe validator for the CityPC preparation bundle updater."""

from __future__ import annotations

import hashlib
import json
import re
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent
MANIFEST = ROOT / "preparacion-bundle-v2.json"
CHANNEL = ROOT / "preparacion-channel-v2.json"
ENGINE = ROOT / "CityPC_Updater.ps1"
EXPECTED = {
    "instalador": ("Instalador_CityPC.bat", ROOT / "version_instalador.txt"),
    "anclados": ("Anclados_y_Limpieza.bat", ROOT / "version_anclados.txt"),
    "onedrive": ("Reactivar_OneDrive.bat", ROOT / "version_reactivar_onedrive.txt"),
}


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def fail(message: str) -> None:
    raise AssertionError(message)


def main() -> int:
    manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
    channel = json.loads(CHANNEL.read_text(encoding="utf-8"))
    if manifest["schema"] != 3 or manifest["bundleId"] != "preparacion":
        fail("invalid manifest identity")
    if channel["schema"] != 1 or channel["bundleId"] != "preparacion":
        fail("invalid channel identity")
    if manifest["tag"] != channel["tag"] or not re.fullmatch(r"preparacion-v\d+", manifest["tag"]):
        fail("channel must point to immutable preparacion-vN tag")
    if channel["manifestFile"] != MANIFEST.name:
        fail("channel manifest filename mismatch")
    if channel["manifestSha256"] != sha256(MANIFEST):
        fail("channel manifest sha mismatch")
    if manifest["updater"]["sha256"] != sha256(ENGINE):
        fail("engine sha mismatch")
    if not (manifest["updater"]["minBytes"] <= ENGINE.stat().st_size <= manifest["updater"]["maxBytes"]):
        fail("engine size outside manifest bounds")

    tools = {entry["id"]: entry for entry in manifest["tools"]}
    if set(tools) != set(EXPECTED):
        fail("bundle must contain exactly installer, anclados and onedrive")
    embedded_engine_hashes: set[str] = set()
    for tool_id, (filename, version_file) in EXPECTED.items():
        entry = tools[tool_id]
        if entry["file"] != filename:
            fail(f"wrong filename for {tool_id}")
        path = ROOT / filename
        raw = path.read_bytes()
        text = raw.decode("ascii")
        if b"\r\n" not in raw:
            fail(f"{filename} must keep CRLF line endings")
        match = re.search(r'^set "LOCAL_VER=(\d+)"\s*$', text, re.M)
        if not match or int(match.group(1)) != entry["version"]:
            fail(f"LOCAL_VER mismatch in {filename}")
        if int(version_file.read_text(encoding="ascii").strip()) != entry["version"]:
            fail(f"independent version file mismatch for {tool_id}")
        if entry["sha256"] != sha256(path):
            fail(f"artifact sha mismatch for {filename}")
        if not (entry["minBytes"] <= path.stat().st_size <= entry["maxBytes"]):
            fail(f"artifact size outside bounds for {filename}")
        if f'set "CITYPC_TOOL_ID={tool_id}"' not in text:
            fail(f"tool id missing in {filename}")
        if f'set "CITYPC_RELEASE_TAG={manifest["tag"]}"' not in text:
            fail(f"immutable bootstrap tag missing in {filename}")
        hash_match = re.search(r'^set "CITYPC_UPDATER_SHA256=([0-9a-f]{64})"\s*$', text, re.M)
        if not hash_match:
            fail(f"embedded updater hash missing in {filename}")
        embedded_engine_hashes.add(hash_match.group(1))
        for required in (
            "--citypc-update-resume",
            'if "!CITYPC_UPDATE_RC!"=="20" exit',
            'if "!CITYPC_UPDATE_RC!"=="30" exit',
            'if "!CITYPC_UPDATE_RC!"=="31" exit',
            "Get-FileHash",
            "CityPC_Updater.ps1",
            "TimeoutSec 12",
            "1..2|ForEach-Object",
            "Reinicio sin actualizador valido. Se detiene para rollback.",
        ):
            if required not in text:
                fail(f"{required!r} missing in {filename}")
        for forbidden in ("GITHUB_RAW=", "/version.txt", "DownloadFile("):
            if forbidden in text:
                fail(f"legacy updater token {forbidden!r} remains in {filename}")
    if embedded_engine_hashes != {sha256(ENGINE)}:
        fail("all BATs must pin the exact common updater")

    phase_a = ROOT / "transition" / "phase-a-main"
    phase_b = ROOT / "transition" / "phase-b-preposition-main"
    phase_c = ROOT / "transition" / "phase-c-activate-main"
    bridge_text = (phase_a / "Anclados_y_Limpieza.bat").read_text(encoding="ascii")
    if 'set "LOCAL_VER=48"' not in bridge_text or 'set "CITYPC_TOOL_ID=anclados"' not in bridge_text:
        fail("phase A must provide exact Anclados V48 bridge")
    if sha256(ENGINE) not in bridge_text:
        fail("phase A bridge must pin current updater hash")
    if (phase_a / "version.txt").read_text(encoding="ascii").strip() != "48":
        fail("phase A must keep legacy version.txt=48")
    if (phase_b / "version.txt").read_text(encoding="ascii").strip() != "48":
        fail("phase B must preposition V49 while keeping legacy version.txt=48")
    for filename, _version_file in EXPECTED.values():
        if sha256(phase_b / filename) != sha256(ROOT / filename):
            fail(f"phase B drift: {filename}")
    if sorted(path.name for path in phase_c.iterdir()) != ["version.txt"]:
        fail("phase C must change only version.txt")
    if (phase_c / "version.txt").read_text(encoding="ascii").strip() != "49":
        fail("phase C must activate legacy version.txt=49")

    legacy = tools["onedrive"]["legacySha256"]
    if not re.fullmatch(r"[0-9a-f]{64}", legacy):
        fail("legacy OneDrive hash missing")
    if tools["instalador"]["legacySha256"] or tools["anclados"]["legacySha256"]:
        fail("legacySha256 is only allowed for OneDrive")

    engine = ENGINE.read_text(encoding="ascii")
    required_engine_tokens = (
        "$script:DownloadDeadlineUtc",
        "AddSeconds(90)",
        "Assert-Manifest",
        "manifestSha256",
        "legacySha256",
        "Get-BatchVersion",
        "-gt [int]$tool.version",
        "transactions",
        "Open-UpdateLock",
        "Restore-Bundle",
        "Rollback no verificable",
        "durableBackupPath",
        "restartCount -ne 1",
        "resume.ok",
        "ExitResumeFailed = 30",
        "ExitUnsafeState = 31",
        "quarantined",
        "Start-CallerRecovery",
        "AddMinutes(15)",
        "preparacion-v[0-9]",
        "refs/tags/",
    )
    for token in required_engine_tokens:
        if token not in engine:
            fail(f"engine guard missing: {token}")
    if engine.count("if ($CurrentVersion -le 0)") != 1:
        fail("CurrentVersion must be checked exactly once")

    timeout_block = re.search(
        r"if \(-not \(Wait-ForProcessExit .*?\n\s*}\n\n\s*\$stateRoot",
        engine,
        re.S,
    )
    if not timeout_block:
        fail("parent timeout block not found")
    if "PARENT_TIMEOUT_NO_RELAUNCH" not in timeout_block.group(0):
        fail("parent timeout static sentinel missing")
    if "Start-CallerRecovery" in timeout_block.group(0):
        fail("parent timeout must never relaunch while the parent is alive")

    restore_block = re.search(
        r"function Restore-Bundle \{.*?\n}\n\nfunction Start-CallerRecovery",
        engine,
        re.S,
    )
    if not restore_block:
        fail("Restore-Bundle block not found")
    restore_text = restore_block.group(0)
    for token in (
        "PREVIOUS_CORRUPT_DURABLE_FALLBACK",
        "Test-FileSha256 -Path $previous -ExpectedHash $oldHash",
        "Test-FileSha256 -Path $durable -ExpectedHash $oldHash",
        "Rollback sin copia verificada",
        "Rollback temporal no verificable",
    ):
        if token not in restore_text:
            fail(f"verified durable rollback guard missing: {token}")
    if "raw.githubusercontent.com/rodrigofufer/CityPC-Installer/main/CityPC_Updater.ps1" in engine:
        fail("engine artifact must never be fetched from mutable main")

    combined = "\n".join(
        (ROOT / name).read_text(encoding="ascii", errors="ignore")
        for name in ["Instalador_CityPC.bat", "Anclados_y_Limpieza.bat", "Reactivar_OneDrive.bat", "CityPC_Updater.ps1"]
    )
    forbidden_secret_patterns = (
        r"(?i)wifi_password\s*=",
        r"(?i)x-citypc-usb-token\s*[:=]",
        r"(?i)bearer\s+[a-z0-9._-]{20,}",
    )
    for pattern in forbidden_secret_patterns:
        if re.search(pattern, combined):
            fail(f"possible secret found: {pattern}")

    print(
        "OK updater-v2 static: bundle=49 tools=3 immutable_tag sha256 rollback "
        "durable-fallback parent-timeout-no-relaunch sentinel offline-safe"
    )
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"FAIL updater-v2 static: {exc}", file=sys.stderr)
        raise SystemExit(1)
