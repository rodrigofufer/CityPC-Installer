#!/usr/bin/env python3
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent
VERSION = 8
COMMON_BAT = ROOT / "Diagnostico_CityPC.bat"
WIFI_HELPER = ROOT / "usbdiag-wifi-readiness.ps1"
BUNDLE_COMMITTER = ROOT / "usbdiag-bundle-commit.ps1"
VERSION_FILE = ROOT / "version_diagnostico.txt"
MANIFEST_FILE = ROOT / "update-manifest.json"
CHANNEL_EXAMPLE = ROOT / "update-channel.json.example"
SHARED_EXAMPLE = ROOT / "usbdiag.shared.local.cmd.example"
WIFI_EXAMPLE = ROOT / "wifi.local.cmd.example"

REQUIRED_PACKAGE_FILES = [
    COMMON_BAT,
    WIFI_HELPER,
    BUNDLE_COMMITTER,
    VERSION_FILE,
    MANIFEST_FILE,
    CHANNEL_EXAMPLE,
    SHARED_EXAMPLE,
    WIFI_EXAMPLE,
    ROOT / ".gitignore",
    ROOT / "README.md",
    ROOT / "INSTALACION.md",
    ROOT / "MIGRACION-V8.md",
    ROOT / "validate_v8_static.py",
    ROOT / "test_v8_bundle_updater.py",
]

COMMON_PAIR_FILES = sorted(
    {path.name for path in REQUIRED_PACKAGE_FILES} | {"usbdiag.shared.local.cmd"}
)


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def parse_cmd_assignments(path: Path) -> dict[str, str]:
    assignments: dict[str, str] = {}
    for line in read_text(path).splitlines():
        match = re.fullmatch(r'\s*set\s+"([A-Z0-9_]+)=([^"\r\n]*)"\s*', line, re.I)
        if not match:
            continue
        assignments[match.group(1).upper()] = match.group(2)
    return assignments


def validate_examples() -> None:
    shared = parse_cmd_assignments(SHARED_EXAMPLE)
    wifi = parse_cmd_assignments(WIFI_EXAMPLE)
    require(
        set(shared) == {"WEBHOOK_URL", "CITYPC_USB_TOKEN"},
        "shared example must contain only WEBHOOK_URL and CITYPC_USB_TOKEN",
    )
    require(
        set(wifi) == {"WIFI_SSID", "WIFI_PASS"},
        "wifi example must contain only WIFI_SSID and WIFI_PASS",
    )
    for name, value in {**shared, **wifi}.items():
        require("REEMPLAZAR_" in value, f"{name} example is not an obvious placeholder")
    require(
        shared["WEBHOOK_URL"].startswith("https://REEMPLAZAR_HOST/"),
        "shared example must use a non-live HTTPS placeholder host",
    )


def validate_bat() -> None:
    raw = COMMON_BAT.read_bytes()
    text = raw.decode("ascii")
    require(raw.startswith(b"@echo off\r\n"), "BAT must be ASCII with CRLF")
    require(b"\n" not in raw.replace(b"\r\n", b""), "BAT contains non-CRLF newlines")
    require('set "LOCAL_VER=8"' in text, "BAT LOCAL_VER is not 8")
    require(
        'set "REMOTE_FILE=Diagnostico_CityPC.bat"' in text,
        "updater does not target the common BAT",
    )
    require("Diagnostico_JP.bat" not in text, "BAT still references JP branch file")
    require("Diagnostico_Urano.bat" not in text, "BAT still references Urano branch file")
    require("ping 8.8.8.8" not in text, "legacy ICMP readiness check remains")

    embedded_secret = re.compile(
        r'(?mi)^\s*set\s+"(?:WEBHOOK_URL|WIFI_SSID|WIFI_PASS|CITYPC_USB_TOKEN)=.+"\s*$'
    )
    require(not embedded_secret.search(text), "BAT embeds local URL or credential")
    urls = re.findall(r"https://[^\s\"']+", text)
    require(
        urls
        == [
            "https://raw.githubusercontent.com/rodrigofufer/CityPC-Installer/main",
            "https://raw.githubusercontent.com/rodrigofufer/CityPC-Installer/!BUNDLE_COMMIT!",
            "https://github.com/rodrigofufer/CityPC-Installer/raw/!BUNDLE_COMMIT!",
        ],
        "BAT contains an unexpected URL",
    )

    required = [
        'set "SHARED_CONFIG=%~dp0usbdiag.shared.local.cmd"',
        'set "WIFI_CONFIG=%~dp0wifi.local.cmd"',
        'set "INSTANCE_LOCK=%temp%\\citypc_usbdiag_v8.lock"',
        'mkdir "%INSTANCE_LOCK%"',
        'set "WIFI_XML=%temp%\\citypc_diag_%UPDATE_NONCE%_wifi.xml"',
        'set "WIFI_PROFILE_NAME=CityPC-Diagnostico-%UPDATE_NONCE%"',
        'set "WIFI_HELPER=%~dp0usbdiag-wifi-readiness.ps1"',
        'set "BUNDLE_COMMITTER=%~dp0usbdiag-bundle-commit.ps1"',
        'set "REMOTE_HELPER_FILE=usbdiag-wifi-readiness.ps1"',
        'set "REMOTE_COMMITTER_FILE=usbdiag-bundle-commit.ps1"',
        'set "CHANNEL_FILE=update-channel.json"',
        'set "UPDATE_TRANSACTION_DIR=%~dp0.usbdiag-update-transaction"',
        '-Action Connect -ReadyTimeoutSeconds 45',
        'if not "%WIFI_RC%"=="0" goto :ERROR_WIFI',
        'SSID, IPv4, DNS y HTTPS',
        '-Action Cleanup',
        ':ERROR_WIFI_CLEANUP',
        'No se ejecutara el diagnostico ni se reportara un falso OK.',
        "citypc.usbdiag.update-channel.v1",
        "^[0-9a-f]{40}$",
        "manifestSha256",
        "-lt [int64]$env:LOCAL_VER",
        "-eq [int64]$env:LOCAL_VER",
        "no se permite downgrade",
        "bundle corrupto o incompleto. Reparando",
        "update-manifest.json",
        "Get-FileHash",
        "@($m.files.PSObject.Properties).Count -ne 3",
        "$manifestActual -cne $env:EXPECTED_MANIFEST_HASH",
        "Get-Anchored $env:REMOTE_HELPER_FILE $env:UPDATE_HELPER_PATH",
        "Get-Anchored $env:REMOTE_COMMITTER_FILE $env:UPDATE_COMMITTER_PATH",
        "$attempt -le 3",
        "AddSeconds(35)",
        "AddSeconds(100)",
        "BUNDLE_FALLBACK_RAW",
        '-File "!UPDATE_COMMITTER_PATH!"',
        '-Action Commit',
        '-ParentPid "!UPDATE_PARENT_PID!"',
        '-Token "!UPDATE_TRANSACTION_TOKEN!"',
        '-DownloadedHelper "!UPDATE_HELPER_PATH!"',
        '-DownloadedCommitter "!UPDATE_COMMITTER_PATH!"',
        '-ManifestExpectedSha256 "!EXPECTED_MANIFEST_HASH!"',
        '-ChannelPath "!UPDATE_CHANNEL_PATH!"',
        '-ValidatedVersionPath "!UPDATE_VALID_VERSION_PATH!"',
        '-CommitPath "!UPDATE_COMMIT_PATH!"',
        '-ManifestExpectedHashPath "!UPDATE_MANIFEST_EXPECTED_HASH_PATH!"',
        "set \"INSTANCE_LOCK_HELD=\"",
        ":CONFIRM_INSTALLED_BUNDLE",
        "--updated\" goto :CONFIRM_INSTALLED_BUNDLE",
        "-Action Confirm",
        "-Action Recover",
        "citypc_diag_recovered_",
        ":CAPTURE_PARENT_PID",
        "ParentProcessId",
        ":UPDATE_COMMIT_START_FAILED",
        ":CLEAN_STATELESS_UPDATE_TRANSACTION",
        ":ACQUIRE_INSTANCE_LOCK",
        "owner.pid",
        "Get-Process -Id ([int]$env:INSTANCE_LOCK_OWNER)",
        "USB_DIAG_PIPELINE_ACCEPTED_V1",
        "citypc_persisted",
        "^[0-9a-fA-F]{64}$",
        "citypc_accepted",
        "run_id",
    ]
    for needle in required:
        require(needle in text, f"BAT missing guard: {needle}")
    require(":UPDATE_RESTORE" not in text, "BAT still overwrites itself in-process")
    require(
        "UPDATE_BAT_BACKUP_PATH" not in text,
        "BAT still performs a partial in-process transaction",
    )
    require(text.count("$attempt -le 3") >= 2, "channel and bundle do not each retry three times")
    require(
        text.count('if not exist "%BUNDLE_COMMITTER%"') == 1,
        "normal startup must not require the installed committer before repair",
    )
    confirm_start = text.index("\r\n:CONFIRM_INSTALLED_BUNDLE\r\n")
    confirm_end = text.index("\r\n:CONFIRM_RECOVERED_BUNDLE\r\n", confirm_start)
    confirm_segment = text[confirm_start:confirm_end]
    require(
        'if not exist "%BUNDLE_COMMITTER%"' in confirm_segment,
        "installed committer is not required at the confirmation boundary",
    )
    update_start = text.index("\r\n:UPDATE_STARTUP_READY\r\n")
    update_end = text.index("\r\n:INICIO\r\n", update_start)
    update_segment = text[update_start:update_end]
    require(
        'if not exist "%BUNDLE_COMMITTER%"' not in update_segment,
        "normal update path blocks repair when the installed committer is absent",
    )
    labels = re.findall(r"(?m)^:([A-Z0-9_]+)\r?$", text, re.I)
    require(len(labels) == len(set(label.upper() for label in labels)), "BAT has duplicate labels")
    longest_line = max(len(line) for line in text.splitlines())
    require(longest_line < 8191, f"BAT command line exceeds cmd.exe limit: {longest_line}")


def validate_wifi_helper() -> None:
    text = read_text(WIFI_HELPER)
    require("https://" not in text, "Wi-Fi helper embeds a URL")
    require(
        not re.search(r'(?mi)^\s*\$env:(?:WIFI_SSID|WIFI_PASS|WEBHOOK_URL)\s*=\s*[\"\']', text),
        "Wi-Fi helper embeds local configuration",
    )
    required = [
        "Invoke-NetshBounded",
        "WaitForExit($TimeoutMilliseconds)",
        "GetConnectedSsid()",
        "$connectedSsid -cne $env:WIFI_SSID",
        "Get-NetIPAddress",
        "-AddressFamily IPv4",
        "169\\.254\\.",
        "Resolve-DnsName",
        "-QuickTimeout",
        "Invoke-WebRequest",
        "-TimeoutSec $timeout",
        "$env:GITHUB_CHANNEL_RAW",
        "$env:CHANNEL_FILE",
        "citypc.usbdiag.update-channel.v1",
        "manifestSha256",
        "$attempt -le 3",
        "ReadyTimeoutSeconds",
        "AddSeconds",
        "connectionMode",
        "manual",
        "Remove-DiagnosticWifiProfile",
        "$profileArgument = 'name=\"{0}\"' -f $env:WIFI_PROFILE_NAME",
        "$writer.WriteElementString('name', $namespace, $env:WIFI_PROFILE_NAME)",
        "('name=\"{0}\"' -f $env:WIFI_PROFILE_NAME)",
        "Remove-Item -LiteralPath $env:WIFI_XML",
        "Assert-TlsHandshake -Uri $webhookUri",
        "System.Net.Security.SslStream",
        "BeginAuthenticateAsClient",
        "EndAuthenticateAsClient",
        "IsAuthenticated",
        "IsEncrypted",
        "Exit-WithCode 46",
        "Exit-WithCode 47",
        "Exit-WithCode 48",
        "Exit-WithCode 49",
    ]
    for needle in required:
        require(needle in text, f"Wi-Fi helper missing guard: {needle}")
    require(
        "$ssidArgument = 'name=" not in text,
        "Wi-Fi helper may delete a pre-existing profile named after the SSID",
    )
    require("$env:GITHUB_RAW" not in text, "Wi-Fi helper uses obsolete GITHUB_RAW")
    require("$env:VERSION_FILE" not in text, "Wi-Fi helper uses obsolete VERSION_FILE")


def validate_bundle_committer() -> None:
    text = read_text(BUNDLE_COMMITTER)
    require("https://" not in text, "bundle committer embeds a URL")
    required = [
        "DownloadedBat",
        "DownloadedHelper",
        "DownloadedCommitter",
        "ManifestPath",
        "ManifestExpectedSha256",
        "ChannelPath",
        "ValidatedVersionPath",
        "CommitPath",
        "ManifestExpectedHashPath",
        "TransactionDir",
        "LockDir",
        "citypc.usbdiag.update-manifest.v1",
        "citypc.usbdiag.update-state.v1",
        "private configuration is not updateable",
        "Write-State([System.Collections.IDictionary]$State)",
        "Wait-ForParentExit",
        "Set-LockOwner",
        "owner.pid",
        "Invoke-Commit",
        "Invoke-Confirm",
        "Invoke-Recover",
        "Restore-FromState",
        "Get-Sha256",
        "phase = 'prepared'",
        "$state.phase = 'installing'",
        "installed-awaiting-confirmation",
        "phase = 'confirmed'",
        "$state.manifestSha256",
        "anchored manifest hash mismatch",
        "confirmation manifest hash mismatch",
        "confirmation token mismatch",
        "confirmation version mismatch",
        "Wait-ForParentExit",
        "@($HelperName, $CommitterName, $BatName)",
        "$attempt -le 20",
        "installed hash mismatch",
        "rollback could not be confirmed",
        "$rollbackConfirmed = $true",
        "$parentExited = $false",
        "if ($Action -ne 'Confirm' -and $parentExited)",
        "if ($Action -eq 'Commit' -and $rollbackConfirmed -and $parentExited)",
        "if ($Action -eq 'Commit' -or $Action -eq 'Recover') { Set-LockOwner }",
        "$backupExisted = Test-Path",
        "backupExisted = [bool]$backupExisted",
        "Move-Item -LiteralPath $TransactionDir -Destination $completedDir",
        "Start-Process -FilePath $env:ComSpec",
        "--updated",
        "$DownloadedCommitter",
    ]
    for needle in required:
        require(needle in text, f"bundle committer missing guard: {needle}")
    for private_name in ("usbdiag.shared.local.cmd", "wifi.local.cmd"):
        require(private_name in text, f"committer does not explicitly protect {private_name}")
    require(
        "Remove-Item -LiteralPath $TransactionDir -Recurse -Force" in text,
        "committer cannot remove a completed or rolled-back transaction",
    )
    prepared_index = text.index("phase = 'prepared'")
    wait_index = text.index("Wait-ForParentExit", prepared_index)
    installing_index = text.index("$state.phase = 'installing'", wait_index)
    copy_index = text.index("Copy-Item -LiteralPath (Join-Path $newDir $name)", installing_index)
    require(
        prepared_index < wait_index < installing_index < copy_index,
        "committer can mutate before durable prepared state and parent exit",
    )
    timeout_guard = text.index("parent BAT did not exit within 60 seconds")
    guarded_restart = text.index(
        "if ($Action -eq 'Commit' -and $rollbackConfirmed -and $parentExited)"
    )
    require(timeout_guard < guarded_restart, "parent timeout can relaunch a duplicate BAT")


def validate_channel_example() -> None:
    channel = json.loads(read_text(CHANNEL_EXAMPLE))
    require(
        channel.get("schema") == "citypc.usbdiag.update-channel.v1",
        "channel example schema mismatch",
    )
    require(channel.get("version") == VERSION, "channel example version mismatch")
    require(channel.get("commit") == "REEMPLAZAR_COMMIT_SHA40", "channel commit is not placeholder")
    require(
        channel.get("manifestSha256") == "REEMPLAZAR_SHA256_MANIFEST",
        "channel manifest hash is not placeholder",
    )


def validate_manifest() -> None:
    manifest = json.loads(read_text(MANIFEST_FILE))
    require(
        manifest.get("schema") == "citypc.usbdiag.update-manifest.v1",
        "manifest schema mismatch",
    )
    require(manifest.get("version") == VERSION, "manifest version mismatch")
    expected_files = {
        "Diagnostico_CityPC.bat": sha256(COMMON_BAT),
        "usbdiag-wifi-readiness.ps1": sha256(WIFI_HELPER),
        "usbdiag-bundle-commit.ps1": sha256(BUNDLE_COMMITTER),
    }
    require(manifest.get("files") == expected_files, "manifest hashes do not match common files")


def validate_pair(urano_root: Path, jp_root: Path) -> None:
    for root in (urano_root, jp_root):
        require(root.is_dir(), f"USB root does not exist: {root}")
        require(not (root / "Diagnostico_JP.bat").exists(), f"legacy JP BAT present in {root}")
        require(not (root / "Diagnostico_Urano.bat").exists(), f"legacy Urano BAT present in {root}")
        for relative in COMMON_PAIR_FILES + ["wifi.local.cmd"]:
            require((root / relative).is_file(), f"missing {relative} in {root}")

    for relative in COMMON_PAIR_FILES:
        urano_file = urano_root / relative
        jp_file = jp_root / relative
        require(
            urano_file.read_bytes() == jp_file.read_bytes(),
            f"forbidden branch drift in {relative}",
        )

    for root in (urano_root, jp_root):
        wifi = parse_cmd_assignments(root / "wifi.local.cmd")
        shared = parse_cmd_assignments(root / "usbdiag.shared.local.cmd")
        require(
            set(wifi) == {"WIFI_SSID", "WIFI_PASS"},
            f"wifi.local.cmd has forbidden keys in {root}",
        )
        require(
            set(shared) == {"WEBHOOK_URL", "CITYPC_USB_TOKEN"},
            f"shared config has forbidden keys in {root}",
        )
        require(
            all(value and "REEMPLAZAR_" not in value for value in wifi.values()),
            f"wifi.local.cmd still contains placeholders in {root}",
        )
        require(
            all(value and "REEMPLAZAR_" not in value for value in shared.values()),
            f"shared config still contains placeholders in {root}",
        )

    require(
        (urano_root / "Diagnostico_CityPC.bat").read_bytes() == COMMON_BAT.read_bytes(),
        "Urano BAT does not match staging hash",
    )
    require(
        (jp_root / "Diagnostico_CityPC.bat").read_bytes() == COMMON_BAT.read_bytes(),
        "JP BAT does not match staging hash",
    )
    print("OK: pair differs only where wifi.local.cmd is allowed to differ")


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate unified CityPC USB diagnosis V8")
    parser.add_argument(
        "--pair",
        nargs=2,
        metavar=("URANO_ROOT", "JP_ROOT"),
        help="also verify two provisioned USB roots",
    )
    args = parser.parse_args()

    for path in REQUIRED_PACKAGE_FILES:
        require(path.is_file(), f"missing package file: {path.name}")
    require(VERSION_FILE.read_text().strip() == str(VERSION), "version file is not 8")
    require(not (ROOT / "Diagnostico_JP.bat").exists(), "legacy JP BAT present")
    require(not (ROOT / "Diagnostico_Urano.bat").exists(), "legacy Urano BAT present")
    require(not (ROOT / "usbdiag.shared.local.cmd").exists(), "private shared config leaked")
    require(not (ROOT / "wifi.local.cmd").exists(), "private Wi-Fi config leaked")

    gitignore = read_text(ROOT / ".gitignore")
    for local_name in (
        "usbdiag.shared.local.cmd",
        "wifi.local.cmd",
        "*.local.cmd",
        ".usbdiag-update-transaction/",
        ".usbdiag-update-transaction.confirmed-*/",
    ):
        require(local_name in gitignore, f".gitignore missing {local_name}")

    validate_examples()
    validate_bat()
    validate_wifi_helper()
    validate_bundle_committer()
    validate_channel_example()
    validate_manifest()

    if args.pair:
        validate_pair(Path(args.pair[0]).resolve(), Path(args.pair[1]).resolve())

    print("OK: USB Diagnostico V8 unified static validation")
    print(f"COMMON_BAT_SHA256={sha256(COMMON_BAT)}")
    print(f"WIFI_HELPER_SHA256={sha256(WIFI_HELPER)}")
    print(f"BUNDLE_COMMITTER_SHA256={sha256(BUNDLE_COMMITTER)}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except AssertionError as exc:
        print(f"FAIL: {exc}", file=sys.stderr)
        raise SystemExit(1)
