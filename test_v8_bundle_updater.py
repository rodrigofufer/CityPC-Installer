#!/usr/bin/env python3
from __future__ import annotations

import json
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parent
VALIDATOR = ROOT / "validate_v8_static.py"


def run_validator(root: Path, *args: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, str(root / "validate_v8_static.py"), *args],
        cwd=root,
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        check=False,
    )


def provision_pair(root: Path, name: str, ssid: str) -> Path:
    destination = root / name
    shutil.copytree(ROOT, destination)
    (destination / "usbdiag.shared.local.cmd").write_text(
        '@echo off\nset "WEBHOOK_URL=https://example.invalid/test"\n'
        'set "CITYPC_USB_TOKEN=test-token"\n',
        encoding="ascii",
    )
    (destination / "wifi.local.cmd").write_text(
        f'@echo off\nset "WIFI_SSID={ssid}"\nset "WIFI_PASS=test-password"\n',
        encoding="ascii",
    )
    return destination


class BundleUpdaterContractTests(unittest.TestCase):
    def test_static_package_passes(self) -> None:
        result = run_validator(ROOT)
        self.assertEqual(result.returncode, 0, result.stdout)

    def test_manifest_covers_exact_runtime_bundle(self) -> None:
        manifest = json.loads((ROOT / "update-manifest.json").read_text(encoding="utf-8"))
        self.assertEqual(
            set(manifest["files"]),
            {
                "Diagnostico_CityPC.bat",
                "usbdiag-wifi-readiness.ps1",
                "usbdiag-bundle-commit.ps1",
            },
        )

    def test_normal_start_can_repair_absent_committer(self) -> None:
        bat = (ROOT / "Diagnostico_CityPC.bat").read_text(encoding="ascii")
        missing_guard = 'if not exist "%BUNDLE_COMMITTER%"'
        self.assertEqual(bat.count(missing_guard), 1)
        confirm_start = bat.index("\n:CONFIRM_INSTALLED_BUNDLE\n")
        confirm_end = bat.index("\n:CONFIRM_RECOVERED_BUNDLE\n", confirm_start)
        confirm = bat[confirm_start:confirm_end]
        update_start = bat.index("\n:UPDATE_STARTUP_READY\n")
        update_end = bat.index("\n:INICIO\n", update_start)
        normal_update = bat[update_start:update_end]
        self.assertIn(missing_guard, confirm)
        self.assertNotIn(missing_guard, normal_update)
        committer = (ROOT / "usbdiag-bundle-commit.ps1").read_text(encoding="utf-8")
        self.assertIn("$backupExisted = Test-Path", committer)
        self.assertIn("backupExisted = [bool]$backupExisted", committer)

    def test_equal_version_repairs_hash_drift(self) -> None:
        bat = (ROOT / "Diagnostico_CityPC.bat").read_text(encoding="ascii")
        self.assertIn('if "!VERSION_CHECK!"=="20"', bat)
        self.assertIn("Get-FileHash -LiteralPath $installed", bat)
        self.assertIn("bundle corrupto o incompleto. Reparando", bat)
        self.assertIn("hashes del bundle estan vigentes", bat)

    def test_channel_and_bundle_retries_are_bounded(self) -> None:
        bat = (ROOT / "Diagnostico_CityPC.bat").read_text(encoding="ascii")
        self.assertGreaterEqual(bat.count("$attempt -le 3"), 2)
        self.assertIn("AddSeconds(35)", bat)
        self.assertIn("AddSeconds(100)", bat)
        self.assertIn("BUNDLE_FALLBACK_RAW", bat)
        self.assertIn("$base=$bases[($attempt-1) %% $bases.Count]", bat)

    def test_bat_and_wifi_helper_share_channel_contract(self) -> None:
        bat = (ROOT / "Diagnostico_CityPC.bat").read_text(encoding="ascii")
        helper = (ROOT / "usbdiag-wifi-readiness.ps1").read_text(encoding="utf-8")
        self.assertIn('set "GITHUB_CHANNEL_RAW=', bat)
        self.assertIn('set "CHANNEL_FILE=update-channel.json"', bat)
        self.assertIn("$env:GITHUB_CHANNEL_RAW", helper)
        self.assertIn("$env:CHANNEL_FILE", helper)
        self.assertIn("citypc.usbdiag.update-channel.v1", helper)
        self.assertNotIn("$env:GITHUB_RAW", helper)
        self.assertNotIn("$env:VERSION_FILE", helper)

    def test_parent_timeout_cannot_unlock_or_duplicate(self) -> None:
        committer = (ROOT / "usbdiag-bundle-commit.ps1").read_text(encoding="utf-8")
        self.assertIn("$parentExited = $false", committer)
        self.assertIn("parent BAT did not exit within 60 seconds", committer)
        self.assertIn("if ($Action -ne 'Confirm' -and $parentExited)", committer)
        self.assertIn(
            "if ($Action -eq 'Commit' -and $rollbackConfirmed -and $parentExited)",
            committer,
        )
        unsafe_restart = "if ($Action -eq 'Commit' -and $rollbackConfirmed) {"
        self.assertNotIn(unsafe_restart, committer)

    def test_power_loss_recovery_contract_is_durable(self) -> None:
        bat = (ROOT / "Diagnostico_CityPC.bat").read_text(encoding="ascii")
        committer = (ROOT / "usbdiag-bundle-commit.ps1").read_text(encoding="utf-8")
        self.assertIn(":CLEAN_STATELESS_UPDATE_TRANSACTION", bat)
        self.assertIn('if exist "%UPDATE_TRANSACTION_DIR%\\state.json"', bat)
        self.assertIn("owner.pid", bat)
        self.assertIn("owner.pid", committer)
        for phase in ("prepared", "installing", "installed-awaiting-confirmation", "confirmed"):
            self.assertIn(phase, committer)
        self.assertLess(
            committer.index("phase = 'prepared'"),
            committer.index("Wait-ForParentExit", committer.index("phase = 'prepared'")),
        )

    def test_private_configuration_is_never_in_bundle_map(self) -> None:
        manifest = json.loads((ROOT / "update-manifest.json").read_text(encoding="utf-8"))
        self.assertNotIn("usbdiag.shared.local.cmd", manifest["files"])
        self.assertNotIn("wifi.local.cmd", manifest["files"])
        committer = (ROOT / "usbdiag-bundle-commit.ps1").read_text(encoding="utf-8")
        self.assertIn("$privateNames = @('usbdiag.shared.local.cmd', 'wifi.local.cmd')", committer)

    def test_tampered_helper_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            candidate = Path(temp) / "candidate"
            shutil.copytree(ROOT, candidate)
            with (candidate / "usbdiag-wifi-readiness.ps1").open("a", encoding="ascii") as handle:
                handle.write("\n# tampered\n")
            result = run_validator(candidate)
            self.assertNotEqual(result.returncode, 0, result.stdout)
            self.assertIn("manifest hashes do not match common files", result.stdout)

    def test_tampered_committer_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            candidate = Path(temp) / "candidate"
            shutil.copytree(ROOT, candidate)
            with (candidate / "usbdiag-bundle-commit.ps1").open("a", encoding="ascii") as handle:
                handle.write("\n# tampered\n")
            result = run_validator(candidate)
            self.assertNotEqual(result.returncode, 0, result.stdout)
            self.assertIn("manifest hashes do not match common files", result.stdout)

    def test_pair_allows_only_wifi_difference(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            temp_root = Path(temp)
            urano = provision_pair(temp_root, "urano", "TEST-URANO")
            jp = provision_pair(temp_root, "jp", "TEST-JP")
            result = run_validator(ROOT, "--pair", str(urano), str(jp))
            self.assertEqual(result.returncode, 0, result.stdout)

            with (jp / "usbdiag-bundle-commit.ps1").open("a", encoding="ascii") as handle:
                handle.write("\n# forbidden drift\n")
            drift = run_validator(ROOT, "--pair", str(urano), str(jp))
            self.assertNotEqual(drift.returncode, 0, drift.stdout)
            self.assertIn("forbidden branch drift", drift.stdout)


if __name__ == "__main__":
    unittest.main(verbosity=2)
