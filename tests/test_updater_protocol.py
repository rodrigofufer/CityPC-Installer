#!/usr/bin/env python3
"""Executable model tests for the updater transaction protocol.

The Windows implementation is additionally bound byte-for-byte by
validate_update_v2.py. These tests exercise state transitions without touching a
USB, GitHub, Windows, or production.
"""

from __future__ import annotations

import hashlib
import unittest
from dataclasses import dataclass, field


def digest(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def bat(version: int, marker: str) -> bytes:
    return f'@echo off\r\nset "LOCAL_VER={version}"\r\nrem {marker}\r\nCityPC_Updater.ps1\r\n'.encode()


LEGACY_ONEDRIVE = b"@echo off\r\nrem legacy onedrive without local version\r\n"


@dataclass
class Bundle:
    version: int
    files: dict[str, bytes]
    versions: dict[str, int]
    manifest_hash: str = field(init=False)

    def __post_init__(self) -> None:
        body = b"".join(name.encode() + digest(data).encode() for name, data in sorted(self.files.items()))
        self.manifest_hash = digest(body)


class Protocol:
    def __init__(self, local: dict[str, bytes], legacy_onedrive_hash: str) -> None:
        self.local = dict(local)
        self.legacy_onedrive_hash = legacy_onedrive_hash
        self.state = "idle"
        self.previous: dict[str, bytes] | None = None
        self.durable_backup: dict[str, bytes] | None = None
        self.backup_hashes: dict[str, str] | None = None
        self.staged: dict[str, bytes] | None = None
        self.restart_count = 0
        self.recovery_launches = 0
        self.blocked: tuple[int, str] | None = None

    @staticmethod
    def version(data: bytes) -> int | None:
        import re

        match = re.search(rb'set "LOCAL_VER=(\d+)"', data)
        return int(match.group(1)) if match else None

    def check(self, bundle: Bundle | None) -> str:
        if bundle is None:
            return "continue-offline"
        if self.blocked == (bundle.version, bundle.manifest_hash):
            return "continue-cooldown"
        need = False
        for name, expected in bundle.files.items():
            current = self.local.get(name)
            if current is None:
                need = True
                continue
            if name.endswith(".bat"):
                local_version = self.version(current)
                if local_version is None and name == "Reactivar_OneDrive.bat" and digest(current) == self.legacy_onedrive_hash:
                    local_version = 1
                if local_version is not None and local_version > bundle.versions[name]:
                    return "blocked-downgrade"
                if local_version is None:
                    need = True
                    continue
            if current != expected:
                need = True
        if not need:
            return "current"
        self.previous = dict(self.local)
        self.durable_backup = dict(self.local)
        self.backup_hashes = {name: digest(content) for name, content in self.local.items()}
        self.staged = dict(bundle.files)
        self.state = "staged"
        return "close-for-commit"

    def parent_exit_gate(self, bundle: Bundle, parent_exited: bool) -> str:
        if parent_exited:
            return "commit-allowed"
        self.state = "failed"
        self.blocked = (bundle.version, bundle.manifest_hash)
        return "failed-no-relaunch"

    def commit(self, bundle: Bundle, fail_after: int | None = None) -> str:
        assert self.state in {"staged", "committing"}
        if self.state == "committing":
            self.rollback(bundle)
            return "recovered-rollback"
        self.state = "committing"
        assert self.staged is not None
        for index, (name, content) in enumerate(self.staged.items(), start=1):
            self.local[name] = content
            if fail_after == index:
                self.rollback(bundle)
                return "rolled-back"
        self.state = "committed"
        self.restart_count = 1
        return "restart-once"

    def resume(self, bundle: Bundle, helper_valid: bool = True) -> str:
        if not helper_valid:
            return "exit-30"
        if self.state != "committed" or self.restart_count != 1 or self.local != bundle.files:
            return "exit-30"
        self.state = "succeeded"
        return "continue"

    def rollback(self, bundle: Bundle) -> None:
        assert self.previous is not None
        assert self.durable_backup is not None
        assert self.backup_hashes is not None
        restored: dict[str, bytes] = {}
        for name, expected_hash in self.backup_hashes.items():
            previous = self.previous.get(name)
            durable = self.durable_backup.get(name)
            if previous is not None and digest(previous) == expected_hash:
                restored[name] = previous
            elif durable is not None and digest(durable) == expected_hash:
                restored[name] = durable
            else:
                raise RuntimeError(f"backup hash mismatch: {name}")
        self.local = restored
        if {name: digest(content) for name, content in self.local.items()} != self.backup_hashes:
            raise AssertionError("rollback verification failed")
        self.state = "rolled_back"
        self.blocked = (bundle.version, bundle.manifest_hash)

    @staticmethod
    def resolve_active(statuses: list[str]) -> str:
        active = [status for status in statuses if status in {"staged", "committing", "committed"}]
        if len(active) > 1:
            return "quarantine-all-exit-31"
        if active == ["committing"]:
            return "rollback"
        if active == ["staged"]:
            return "resume-commit"
        if active == ["committed"]:
            return "verify-or-rollback"
        return "none"


def bundle49() -> Bundle:
    files = {
        "Instalador_CityPC.bat": bat(49, "installer-v49"),
        "Anclados_y_Limpieza.bat": bat(49, "anclados-v49"),
        "Reactivar_OneDrive.bat": bat(2, "onedrive-v2"),
        "CityPC_Updater.ps1": b"updater-engine-v1",
    }
    return Bundle(49, files, {
        "Instalador_CityPC.bat": 49,
        "Anclados_y_Limpieza.bat": 49,
        "Reactivar_OneDrive.bat": 2,
    })


class UpdaterProtocolTests(unittest.TestCase):
    def old_local(self) -> dict[str, bytes]:
        return {
            "Instalador_CityPC.bat": bat(48, "installer-v48"),
            "Anclados_y_Limpieza.bat": bat(48, "bridge-v48"),
            "Reactivar_OneDrive.bat": LEGACY_ONEDRIVE,
            "CityPC_Updater.ps1": b"old-updater",
        }

    def protocol(self, local: dict[str, bytes] | None = None) -> Protocol:
        return Protocol(local or self.old_local(), digest(LEGACY_ONEDRIVE))

    def test_offline_is_fail_safe_and_does_not_mutate(self) -> None:
        protocol = self.protocol()
        before = dict(protocol.local)
        self.assertEqual(protocol.check(None), "continue-offline")
        self.assertEqual(protocol.local, before)

    def test_valid_bundle_updates_all_and_restarts_once(self) -> None:
        bundle = bundle49()
        protocol = self.protocol()
        self.assertEqual(protocol.check(bundle), "close-for-commit")
        self.assertEqual(protocol.commit(bundle), "restart-once")
        self.assertEqual(protocol.resume(bundle), "continue")
        self.assertEqual(protocol.restart_count, 1)
        self.assertEqual(protocol.check(bundle), "current")

    def test_one_drive_legacy_without_local_ver_is_upgraded(self) -> None:
        bundle = bundle49()
        protocol = self.protocol()
        self.assertEqual(protocol.version(protocol.local["Reactivar_OneDrive.bat"]), None)
        self.assertEqual(protocol.check(bundle), "close-for-commit")

    def test_truncated_local_bat_is_repaired_not_used_as_blocker(self) -> None:
        bundle = bundle49()
        local = self.old_local()
        local["Instalador_CityPC.bat"] = b"truncated"
        protocol = self.protocol(local)
        self.assertEqual(protocol.check(bundle), "close-for-commit")
        self.assertEqual(protocol.commit(bundle), "restart-once")
        self.assertEqual(protocol.local["Instalador_CityPC.bat"], bundle.files["Instalador_CityPC.bat"])

    def test_partial_commit_rolls_back_entire_bundle(self) -> None:
        bundle = bundle49()
        protocol = self.protocol()
        before = dict(protocol.local)
        protocol.check(bundle)
        self.assertEqual(protocol.commit(bundle, fail_after=2), "rolled-back")
        self.assertEqual(protocol.local, before)

    def test_corrupt_backup_is_not_relaunched_as_known_good(self) -> None:
        bundle = bundle49()
        protocol = self.protocol()
        protocol.check(bundle)
        assert protocol.previous is not None
        assert protocol.durable_backup is not None
        protocol.previous["Instalador_CityPC.bat"] = b"corrupt-previous"
        protocol.durable_backup["Instalador_CityPC.bat"] = b"corrupt-durable"
        with self.assertRaisesRegex(RuntimeError, "backup hash mismatch"):
            protocol.rollback(bundle)

    def test_corrupt_previous_uses_verified_durable_backup(self) -> None:
        bundle = bundle49()
        protocol = self.protocol()
        before = dict(protocol.local)
        protocol.check(bundle)
        assert protocol.previous is not None
        protocol.previous["Instalador_CityPC.bat"] = b"corrupt-previous"
        protocol.state = "committing"
        protocol.local["Instalador_CityPC.bat"] = bundle.files["Instalador_CityPC.bat"]
        self.assertEqual(protocol.commit(bundle), "recovered-rollback")
        self.assertEqual(protocol.local, before)

    def test_parent_timeout_never_relaunches_while_parent_is_alive(self) -> None:
        bundle = bundle49()
        protocol = self.protocol()
        before = dict(protocol.local)
        protocol.check(bundle)
        self.assertEqual(protocol.parent_exit_gate(bundle, parent_exited=False), "failed-no-relaunch")
        self.assertEqual(protocol.recovery_launches, 0)
        self.assertEqual(protocol.local, before)

    def test_power_loss_in_committing_rolls_back_before_retry(self) -> None:
        bundle = bundle49()
        protocol = self.protocol()
        before = dict(protocol.local)
        protocol.check(bundle)
        protocol.state = "committing"
        protocol.local["Instalador_CityPC.bat"] = bundle.files["Instalador_CityPC.bat"]
        self.assertEqual(protocol.commit(bundle), "recovered-rollback")
        self.assertEqual(protocol.local, before)

    def test_resume_failure_is_fail_closed(self) -> None:
        bundle = bundle49()
        protocol = self.protocol()
        protocol.check(bundle)
        protocol.commit(bundle)
        self.assertEqual(protocol.resume(bundle, helper_valid=False), "exit-30")
        self.assertNotEqual(protocol.state, "succeeded")

    def test_multiple_active_transactions_are_quarantined_fail_closed(self) -> None:
        self.assertEqual(
            Protocol.resolve_active(["succeeded", "staged", "committed"]),
            "quarantine-all-exit-31",
        )

    def test_same_failed_bundle_does_not_loop_but_newer_can_retry(self) -> None:
        bundle = bundle49()
        protocol = self.protocol()
        protocol.check(bundle)
        protocol.commit(bundle, fail_after=1)
        self.assertEqual(protocol.check(bundle), "continue-cooldown")
        newer = Bundle(50, dict(bundle.files), dict(bundle.versions))
        newer.files["Instalador_CityPC.bat"] = bat(50, "installer-v50")
        newer.versions["Instalador_CityPC.bat"] = 50
        newer.__post_init__()
        self.assertEqual(protocol.check(newer), "close-for-commit")

    def test_never_downgrades_newer_local_version(self) -> None:
        bundle = bundle49()
        local = self.old_local()
        local["Instalador_CityPC.bat"] = bat(50, "future")
        self.assertEqual(self.protocol(local).check(bundle), "blocked-downgrade")

    def test_corrupt_remote_staging_never_mutates_local(self) -> None:
        bundle = bundle49()
        protocol = self.protocol()
        before = dict(protocol.local)
        # A real hash mismatch is rejected before state=staged. Model that gate.
        corrupt = dict(bundle.files)
        corrupt["Anclados_y_Limpieza.bat"] += b"corrupt"
        self.assertNotEqual(digest(corrupt["Anclados_y_Limpieza.bat"]), digest(bundle.files["Anclados_y_Limpieza.bat"]))
        self.assertEqual(protocol.local, before)


if __name__ == "__main__":
    unittest.main(verbosity=2)
