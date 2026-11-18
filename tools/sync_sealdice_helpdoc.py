#!/usr/bin/env python3
"""Install a sequence-library helpdoc export into a SealDice instance."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import stat
import subprocess
import sys
import tarfile
import tempfile
import time
from contextlib import contextmanager
from http.client import HTTPException
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urlsplit
from urllib.request import HTTPRedirectHandler, ProxyHandler, Request, build_opener

GROUPS = ("公共建筑", "技能表", "能量池", "特质改造", "职业")
MANIFEST = ".seqlib-helpdoc-manifest.json"
LEGACY_ZIP = "宏观界域.zip"
EXCLUDED = {"新过审序列", "荣誉室", "desktop.ini", "thumbs.db"}
MAX_MEMBER_BYTES = 64 * 1024 * 1024
MAX_BUNDLE_BYTES = 512 * 1024 * 1024
MAX_BUNDLE_MEMBERS = 20_000
API_REPLY_LIMIT = 16 * 1024 * 1024
RELOAD_WAIT_SECONDS = 180
POLL_INTERVAL = 0.5


class SyncError(Exception):
    """An invalid export or an unsuccessful installation."""


def mode(path: Path) -> int:
    try:
        return path.lstat().st_mode
    except FileNotFoundError:
        return 0


def json_object(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise ValueError("duplicate JSON key")
        result[key] = value
    return result


def reject_constant(_value):
    raise ValueError("non-standard JSON constant")


def read_json(data: bytes, where: str) -> dict:
    try:
        value = json.loads(data.decode("utf-8"), object_pairs_hook=json_object,
                           parse_constant=reject_constant)
    except (ValueError, RecursionError) as exc:
        raise SyncError(f"invalid UTF-8 JSON: {where}") from exc
    if not isinstance(value, dict):
        raise SyncError(f"expected a JSON object: {where}")
    return value


def path_parts(name: str) -> tuple[str, ...]:
    parts = tuple(name.split("/"))
    if any(not p or p in (".", "..") or p.casefold() in EXCLUDED
           or (p.startswith(".") and (parts != (MANIFEST,)))
           or any(ord(c) < 32 or c in '\\:*?"<>|' for c in p) for p in parts):
        raise SyncError(f"unsafe export path: {name!r}")
    if parts != (MANIFEST,) and parts[0] not in GROUPS:
        raise SyncError(f"unexpected top-level export path: {name!r}")
    return parts


def validate_export(root: Path, strict: bool = False) -> dict:
    """Validate an export's managed groups, UTF-8 JSON, hashes and totals.

    Strict mode also rejects unmanaged siblings, as required for staged bundles.
    """
    root = Path(root)
    if not stat.S_ISDIR(mode(root)):
        raise SyncError(f"export is not a real directory: {root}")
    if strict and {p.name for p in root.iterdir()} != {*GROUPS, MANIFEST}:
        raise SyncError("export must contain exactly the five managed groups and the manifest")
    manifest_path = root / MANIFEST
    if not stat.S_ISREG(mode(manifest_path)):
        raise SyncError(f"missing or invalid manifest: {manifest_path}")
    manifest = read_json(manifest_path.read_bytes(), MANIFEST)
    if type(manifest.get("schema_version")) is not int or manifest["schema_version"] != 1:
        raise SyncError("unsupported manifest schema version")
    commit = manifest.get("commit")
    if not isinstance(commit, str) or not commit.strip() or any(ord(c) < 32 for c in commit):
        raise SyncError("manifest needs a nonempty commit identifier")
    expected_files = manifest.get("files")
    expected_groups = manifest.get("groups")
    if not isinstance(expected_files, dict) or not isinstance(expected_groups, dict):
        raise SyncError("manifest needs file and group inventories")
    if set(expected_groups) != set(GROUPS):
        raise SyncError("manifest group inventory does not match managed groups")
    for name in expected_files:
        parts = path_parts(name)
        if len(parts) < 2 or not name.endswith(".json"):
            raise SyncError(f"invalid manifest file path: {name!r}")

    actual_paths = set()
    group_counts = {}
    for group in GROUPS:
        top = root / group
        if not stat.S_ISDIR(mode(top)):
            raise SyncError(f"missing or invalid group directory: {group}")
        paths = [top]
        files = entries = 0
        while paths:
            for path in paths.pop().iterdir():
                name = path.relative_to(root).as_posix()
                path_parts(name)
                kind = mode(path)
                if stat.S_ISDIR(kind):
                    paths.append(path)
                    continue
                if not stat.S_ISREG(kind) or path.suffix != ".json":
                    raise SyncError(f"expected an ordinary JSON file: {name!r}")
                record = expected_files.get(name)
                if not isinstance(record, dict):
                    raise SyncError(f"file not listed in manifest: {name!r}")
                data = path.read_bytes()
                if hashlib.sha256(data).hexdigest() != record.get("sha256"):
                    raise SyncError(f"manifest SHA-256 mismatch: {name!r}")
                doc = read_json(data, name)
                metadata = {"mod", "author", "brief", "comment"}
                if set(doc) != metadata | {"helpdoc"} or not all(isinstance(doc[k], str) for k in metadata):
                    raise SyncError(f"invalid SealDice helpdoc metadata: {name!r}")
                rows = doc["helpdoc"]
                if not isinstance(rows, dict) or not rows or any(
                    not title.strip() or not isinstance(body, str) or not body.strip()
                    for title, body in rows.items()
                ):
                    raise SyncError(f"invalid SealDice helpdoc entries: {name!r}")
                check_count(record, "entries", len(rows), name)
                actual_paths.add(name)
                files += 1
                entries += len(rows)
        group_counts[group] = (files, entries)
        record = expected_groups[group]
        if not isinstance(record, dict):
            raise SyncError(f"invalid manifest group inventory: {group}")
        check_count(record, "files", files, group)
        check_count(record, "entries", entries, group)
    if set(expected_files) != actual_paths or not actual_paths:
        raise SyncError("manifest file inventory does not match export or export is empty")
    check_count(manifest, "file_count", len(actual_paths), MANIFEST)
    check_count(manifest, "entry_count", sum(entries for _, entries in group_counts.values()), MANIFEST)
    return manifest


def check_count(record: dict, key: str, actual: int, where: str) -> None:
    if type(record.get(key)) is not int or record[key] != actual:
        raise SyncError(f"manifest {key} mismatch: {where}")


def extract_bundle(bundle: Path, stage: Path) -> None:
    """Extract only approved regular files/directories, never archive links or modes."""
    if not stat.S_ISREG(mode(bundle)):
        raise SyncError(f"bundle is not a regular tar.gz file: {bundle}")
    archive_paths = set()
    total_size = 0
    with tarfile.open(bundle, "r|gz") as archive:
        for count, member in enumerate(archive, start=1):
            if count > MAX_BUNDLE_MEMBERS:
                raise SyncError("bundle contains too many members")
            name = member.name
            if name in (".", "./") and member.isdir():
                if "." in archive_paths:
                    raise SyncError("duplicate bundle root member")
                archive_paths.add(".")
                continue
            if name.startswith("./"):
                name = name[2:]
            if member.isdir() and name.endswith("/"):
                name = name[:-1]
            parts = path_parts(name)
            if name in archive_paths:
                raise SyncError(f"duplicate bundle member: {name!r}")
            archive_paths.add(name)
            target = stage.joinpath(*parts)
            if member.isdir() and parts != (MANIFEST,):
                target.mkdir(parents=True, exist_ok=True)
                continue
            if (member.type not in (tarfile.REGTYPE, tarfile.AREGTYPE)
                    or not (parts == (MANIFEST,) or (len(parts) > 1 and name.endswith(".json")))):
                raise SyncError(f"invalid bundle file type or path: {name!r}")
            total_size += member.size
            if not 0 <= member.size <= MAX_MEMBER_BYTES or total_size > MAX_BUNDLE_BYTES:
                raise SyncError("bundle exceeds extraction size limit")
            target.parent.mkdir(parents=True, exist_ok=True)
            source = archive.extractfile(member)
            if source is None:
                raise SyncError(f"could not read bundle member: {name!r}")
            with source, target.open("xb") as output:
                shutil.copyfileobj(source, output)


def yaml_token(path: Path) -> str:
    """Read the first token from standard SealDice-generated accessTokens YAML."""
    try:
        text = path.read_text(encoding="utf-8-sig")
    except (OSError, UnicodeError):
        return ""
    try:
        import yaml  # Optional; the bundle installer requires only the standard library.
    except ImportError:
        pass
    else:
        try:
            config = yaml.safe_load(text)
        except yaml.YAMLError:
            return ""
        tokens = config.get("accessTokens") if isinstance(config, dict) else None
        return tokens[0] if isinstance(tokens, list) and tokens and isinstance(tokens[0], str) else ""
    lines = text.splitlines()
    for index, line in enumerate(lines):
        match = re.fullmatch(r"accessTokens\s*:\s*(.*)", line)
        if not match:
            continue
        value = match[1]
        if value.startswith("["):
            # SealDice emits a block list; also support a simple flow list.
            token = re.match(r"\[\s*('(?:[^']|'')*'|\"(?:[^\"\\]|\\.)*\"|[^,\]#]+)\s*[,\]]", value)
            return token_scalar(token[1]) if token else ""
        if value and not value.startswith("#"):
            return ""
        for item in lines[index + 1:]:
            if not item.strip() or item.lstrip().startswith("#"):
                continue
            first = re.fullmatch(r"[ \t]*-[ \t]+(.*)", item)
            return token_scalar(first[1]) if first else ""
    return ""


def token_scalar(value: str) -> str:
    value = value.strip()
    if value.startswith("'"):
        match = re.fullmatch(r"'((?:[^']|'')*)'\s*(?:#.*)?", value)
        return match[1].replace("''", "'") if match else ""
    if value.startswith('"'):
        match = re.fullmatch(r'("(?:[^"\\]|\\.)*")\s*(?:#.*)?', value)
        try:
            return json.loads(match[1]) if match else ""
        except ValueError:
            return ""
    return re.split(r"\s+#", value, maxsplit=1)[0].strip()


class NoRedirect(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


class SealDice:
    def __init__(self, url: str, token: str):
        try:
            parts = urlsplit(url)
            parts.port
        except ValueError:
            raise SyncError("invalid --api-url") from None
        if (parts.scheme not in {"http", "https"} or not parts.hostname or parts.username is not None
                or parts.query or parts.fragment or any(ord(c) <= 32 for c in url)):
            raise SyncError("invalid --api-url; expected an HTTP(S) server URL without credentials")
        if not re.fullmatch(r"[!-~]+", token):
            raise SyncError("no valid SealDice token; use --token, SEALDICE_API_TOKEN or data/dice.yaml")
        self.url = url.rstrip("/")
        if not self.url.endswith("/sd-api"):
            self.url += "/sd-api"
        self.token = token
        self.opener = build_opener(ProxyHandler({}), NoRedirect())

    def call(self, method: str, endpoint: str, payload: dict | None = None, *,
             allow_busy: bool = False, timeout: float = 15) -> dict:
        path = f"/helpdoc/{endpoint}"
        data = json.dumps(payload).encode("utf-8") if payload is not None else None
        req = Request(self.url + path, method=method, data=data,
                      headers={"token": self.token, "Content-Type": "application/json"})
        try:
            with self.opener.open(req, timeout=timeout) as response:
                data = response.read(API_REPLY_LIMIT + 1)
        except HTTPError as exc:
            raise SyncError(f"SealDice {method} {path}: HTTP {exc.code}") from None
        except (URLError, HTTPException, OSError, ValueError):
            raise SyncError(f"SealDice {method} {path}: request failed") from None
        if len(data) > API_REPLY_LIMIT:
            raise SyncError(f"SealDice {method} {path}: response too large")
        result = read_json(data, path)
        if result.get("result") is not True and not (allow_busy and self.busy(result)):
            raise SyncError(f"SealDice {method} {path} did not report success")
        return result

    @staticmethod
    def busy(response: dict) -> bool:
        error = response.get("err")
        return response.get("loading") is True or (
            response.get("result") is False and isinstance(error, str)
            and any(marker in error for marker in ("正在重新装载", "正在加载"))
        )

    @staticmethod
    def time_left(deadline: float) -> float:
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            raise SyncError("timed out waiting for SealDice helpdoc reload")
        return remaining

    def wait_for_tree(self, deadline: float) -> dict:
        while True:
            response = self.call("GET", "tree", allow_busy=True, timeout=min(15, self.time_left(deadline)))
            if not self.busy(response):
                return response
            time.sleep(min(POLL_INTERVAL, self.time_left(deadline)))

    def reload(self) -> dict:
        deadline = time.monotonic() + RELOAD_WAIT_SECONDS
        while True:
            response = self.call("POST", "reload", {}, allow_busy=True,
                                 timeout=min(120, self.time_left(deadline)))
            tree = self.wait_for_tree(deadline)
            if not self.busy(response):
                return tree
            time.sleep(min(POLL_INTERVAL, self.time_left(deadline)))

    def verify(self, response: dict) -> None:
        tree = response.get("data")
        if not isinstance(tree, list):
            raise SyncError("SealDice tree API did not return a list")
        for group in GROUPS:
            groups = [node for node in tree if isinstance(node, dict)
                      and node.get("name") == group and node.get("isDir") is True]
            if len(groups) != 1:
                raise SyncError(f"SealDice tree does not contain the managed group: {group}")
            self.verify_tree(groups[0], group)
        for group in GROUPS:
            response = self.call("POST", "textitem/get_page", {"pageNum": 1, "pageSize": 5, "group": group})
            if type(response.get("total")) is not int or response["total"] <= 0:
                raise SyncError(f"SealDice search has no indexed entries for group: {group}")

    @staticmethod
    def verify_tree(root: dict, group: str) -> None:
        nodes = [root]
        files = 0
        while nodes:
            node = nodes.pop()
            if not isinstance(node, dict) or node.get("deleted") is True:
                raise SyncError(f"SealDice tree contains an invalid or deleted node in group: {group}")
            if node.get("isDir") is True:
                children = node.get("children")
                if not isinstance(children, list):
                    raise SyncError(f"SealDice tree did not return group children: {group}")
                nodes.extend(children)
            elif node.get("isDir") is False and type(node.get("loadStatus")) is int and node["loadStatus"] == 1:
                files += 1
            else:
                raise SyncError(f"SealDice tree contains an unloaded or failed file in group: {group}")
        if not files:
            raise SyncError(f"SealDice tree has no loaded files in group: {group}")


@contextmanager
def sync_lock(path: Path):
    try:
        path.mkdir()
    except FileExistsError:
        raise SyncError(f"another sync is running (or left a lock at {path})") from None
    try:
        yield
    finally:
        path.rmdir()


def same_commit(target: Path, staged: dict) -> bool:
    try:
        return validate_export(target)["commit"] == staged["commit"]
    except (SyncError, OSError):
        return False


def install(stage: Path, target: Path, backup: Path, api: SealDice | None, clean_zip: bool) -> None:
    names = (*GROUPS, MANIFEST)
    old_names = (*names, LEGACY_ZIP) if clean_zip else names
    present = []
    for name in old_names:
        kind = mode(target / name)
        if kind:
            expected = stat.S_ISDIR if name in GROUPS else stat.S_ISREG
            if not expected(kind):
                raise SyncError(f"refusing to replace non-{('directory' if name in GROUPS else 'regular file')}: {target / name}")
            present.append(name)
    if mode(backup) and not stat.S_ISDIR(mode(backup)):
        raise SyncError(f"backup path is not a real directory: {backup}")
    if target.stat().st_dev != backup.parent.stat().st_dev:
        raise SyncError("helpdoc and backup must be on the same filesystem for rename-based rollback")

    work = Path(tempfile.mkdtemp(prefix=".helpdoc-seqlib-txn-", dir=backup.parent))
    original = work / "original"
    previous = work / "previous"
    original.mkdir()
    removed = []
    installed = []
    moved_previous = published = attempted_reload = False
    try:
        for name in present:
            (target / name).rename(original / name)
            removed.append(name)
        for name in names:
            (stage / name).rename(target / name)
            installed.append(name)
        if api:
            attempted_reload = True
            api.verify(api.reload())
        if mode(backup):
            backup.rename(previous)
            moved_previous = True
        original.rename(backup)
        published = True
    except (Exception, KeyboardInterrupt) as exc:
        try:
            if published:
                backup.rename(original)
            if moved_previous:
                previous.rename(backup)
            for name in reversed(installed):
                (target / name).rename(stage / name)
            for name in reversed(removed):
                (original / name).rename(target / name)
        except (Exception, KeyboardInterrupt) as rollback:
            raise SyncError(f"deployment failed ({exc}); rollback failed ({rollback}); recover original files from {work}") from None
        reload_error = None
        if attempted_reload:
            try:
                api.reload()
            except (SyncError, KeyboardInterrupt) as rollback:
                reload_error = rollback
        try:
            shutil.rmtree(work)
        except OSError:
            print(f"warning: could not clean rollback scratch: {work}", file=sys.stderr)
        if reload_error is not None:
            raise SyncError(f"deployment failed ({exc}); disk restored, but SealDice rollback reload failed ({reload_error})") from None
        raise SyncError(f"deployment failed ({exc}); previous helpdocs restored") from None
    try:
        shutil.rmtree(work)
    except OSError as exc:
        print(f"warning: installation succeeded but could not remove old backup scratch {work}: {exc}", file=sys.stderr)


def build_source(source: Path, stage: Path) -> None:
    try:
        if __package__:
            from .build_sealdice_helpdoc import build
        else:
            from build_sealdice_helpdoc import build
    except ImportError:
        raise SyncError("--source-dir needs the repository's build_sealdice_helpdoc.py and text_encoding.py; --bundle is standalone") from None
    build(source, stage, commit=None, check=False)


def run(args: argparse.Namespace) -> None:
    base = args.sealdice_dir.expanduser().resolve()
    target = args.helpdoc_dir.expanduser() if args.helpdoc_dir else base / "data" / "helpdoc"
    target = target.parent.resolve() / target.name
    backup_parent = base / "backups"
    backup = backup_parent / "helpdoc-seqlib-prev"
    if (target == target.parent or target == backup_parent or target in backup_parent.parents
            or backup_parent in target.parents):
        raise SyncError("helpdoc directory must not overlap the backup directory")
    if mode(target) and not stat.S_ISDIR(mode(target)):
        raise SyncError(f"helpdoc path is not a real directory: {target}")
    target.parent.mkdir(parents=True, exist_ok=True)
    lock = target.parent / f".{target.name}.seqlib-sync-lock"
    with sync_lock(lock), tempfile.TemporaryDirectory(prefix=".seqlib-helpdoc-stage-", dir=target.parent) as scratch:
        stage = Path(scratch) / "export"
        if args.bundle:
            stage.mkdir()
            extract_bundle(args.bundle.expanduser(), stage)
        else:
            build_source(args.source_dir.expanduser(), stage)
        manifest = validate_export(stage, strict=True)
        if not args.force and same_commit(target, manifest):
            print(f"Already installed: {manifest['commit']}; no changes.")
            return
        api = None
        if not args.skip_reload:
            token = args.token or os.environ.get("SEALDICE_API_TOKEN") or yaml_token(base / "data" / "dice.yaml")
            api = SealDice(args.api_url, token)
        if mode(backup_parent) and not stat.S_ISDIR(mode(backup_parent)):
            raise SyncError(f"backup parent is not a real directory: {backup_parent}")
        backup_parent.mkdir(parents=True, exist_ok=True)
        target.mkdir(exist_ok=True)
        install(stage, target, backup, api, args.clean_legacy_zip)
        print(f"Installed {manifest['commit']}: {manifest['file_count']} files, {manifest['entry_count']} entries.")
        print(f"Previous managed helpdocs: {backup}")
        if args.skip_reload:
            print("SealDice reload and API verification skipped.")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    inputs = parser.add_mutually_exclusive_group(required=True)
    inputs.add_argument("--bundle", type=Path, help="prepared tar.gz export (no repository required)")
    inputs.add_argument("--source-dir", type=Path, help="repository source for the helpdoc builder")
    parser.add_argument("--sealdice-dir", type=Path, default=Path("/root/sealdice/15"))
    parser.add_argument("--helpdoc-dir", type=Path, help="default: <sealdice-dir>/data/helpdoc")
    parser.add_argument("--api-url", default="http://127.0.0.1:3211")
    parser.add_argument("--token", help="override SEALDICE_API_TOKEN and data/dice.yaml; prefer the environment")
    parser.add_argument("--clean-legacy-zip", action="store_true", help=f"move the top-level {LEGACY_ZIP} into the backup")
    parser.add_argument("--force", action="store_true", help="reinstall even when the validated commit matches")
    parser.add_argument("--skip-reload", action="store_true", help="install files without calling SealDice")
    args = parser.parse_args(argv)
    try:
        run(args)
    except (SyncError, OSError, tarfile.TarError, ValueError, subprocess.CalledProcessError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        print("error: interrupted", file=sys.stderr)
        return 130
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
