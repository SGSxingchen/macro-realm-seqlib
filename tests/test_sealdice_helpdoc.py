import io
import json
import os
import re
import subprocess
import sys
import tarfile
import tempfile
import threading
import unittest
from collections import Counter, defaultdict
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from text_encoding import read_text
from tools import build_sealdice_helpdoc as builder

ROOT = Path(__file__).resolve().parents[1]
TOOLS = ROOT / "tools"
GROUPS = ("公共建筑", "技能表", "能量池", "特质改造", "职业")
MANIFEST = ".seqlib-helpdoc-manifest.json"
HEAD = re.compile(r"^\s*(?:\[[^\]]*名称\]|【[^】]*名称】|［[^］]*名称］|〖[^〗]*名称〗|〔[^〕]*名称〕)\s*[:：](.*)$")


def run(script, *args, no_yaml=False):
    env = {**os.environ, "PYTHONDONTWRITEBYTECODE": "1"}
    env.pop("SEALDICE_API_TOKEN", None)
    return subprocess.run([sys.executable, *(["-S"] if no_yaml else []), str(TOOLS / script), *map(str, args)],
                          cwd=ROOT, env=env, text=True, capture_output=True)


def load(path):
    return json.loads(path.read_bytes(), object_pairs_hook=builder.unique_object)


def temporary_directory():
    return tempfile.TemporaryDirectory(prefix=".test-sealdice-", dir=ROOT)


class RealBuildTest(unittest.TestCase):
    def test_full_export_and_lossless_cards(self):
        with temporary_directory() as workspace:
            output = Path(workspace) / "out"
            archive = Path(workspace) / "bundle.tar.gz"
            result = run("build_sealdice_helpdoc.py", "--check", "--output-dir", output, "--archive", archive)
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertEqual({p.name for p in output.iterdir()}, {*GROUPS, MANIFEST})
            manifest = load(output / MANIFEST)
            self.assertEqual(manifest["file_count"], 691)
            self.assertGreaterEqual(manifest["entry_count"], 5000)
            sources = {p.relative_to(ROOT / "序列库") for group in GROUPS
                       for p in (ROOT / "序列库" / group).rglob("*.txt")
                       if not any(part.startswith(".") or part in {"新过审序列", "荣誉室"}
                                  for part in p.relative_to(ROOT / "序列库").parts)}
            expected = {p.with_suffix(".json").as_posix() for p in sources}
            self.assertEqual(len(expected), 691)
            self.assertEqual(set(manifest["files"]), expected)
            self.assertEqual({p.relative_to(output).as_posix() for p in output.rglob("*.json")}, {*expected, MANIFEST})
            raw_owners = defaultdict(set)
            cards = []
            all_titles = []
            for rel in sorted(sources):
                raw, _ = read_text(ROOT / "序列库" / rel)
                doc = load(output / rel.with_suffix(".json"))
                overview, *entries = doc["helpdoc"].items()
                heads = [m[1].strip() or f"{doc['mod']}·未命名条目" for line in raw.splitlines()
                         if (m := HEAD.fullmatch(line))]
                self.assertEqual(len(entries), len(heads), rel)
                self.assertTrue(all(k.strip() and v.strip() for k, v in doc["helpdoc"].items()), rel)
                self.assertEqual(manifest["files"][rel.with_suffix(".json").as_posix()]["entries"], len(entries) + 1)
                all_titles.extend(doc["helpdoc"])
                raw_owners[doc["mod"]].add(rel)
                if not heads:
                    self.assertEqual(overview[1], raw.strip(), rel)
                    continue
                pos = 0
                leftovers = []
                for raw_title, (title, body) in zip(heads, entries):
                    at = raw.find(body, pos)
                    self.assertGreaterEqual(at, pos, (rel, title))
                    self.assertTrue(at == 0 or raw[at - 1] in "\r\n", (rel, title))
                    self.assertEqual(sum(HEAD.fullmatch(line) is not None for line in body.splitlines()), 1)
                    leftovers.append(raw[pos:at])
                    pos = at + len(body)
                    raw_owners[raw_title].add(rel)
                    cards.append((rel, doc["mod"], raw_title, title))
                leftovers.append(raw[pos:])
                self.assertIn("【收录词条目录】", overview[1])
                self.assertEqual(Counter(line for block in leftovers for line in block.splitlines() if line.strip()),
                                 Counter(line for line in overview[1].rsplit("\n\n【收录词条目录】", 1)[0].splitlines() if line.strip()), rel)
            self.assertEqual(len(all_titles), manifest["entry_count"])
            self.assertEqual(len(all_titles), len(set(all_titles)))
            cross_file = 0
            for rel, resource, raw, final in cards:
                if len(raw_owners[raw]) > 1:
                    self.assertTrue(final.startswith(f"{raw}〔{resource}〕"), (rel, raw, final))
                    cross_file += 1
            self.assertGreater(cross_file, 0)
            with tarfile.open(archive) as bundle:
                self.assertEqual({item.name for item in bundle.getmembers() if item.isfile()}, {*expected, MANIFEST})


class BuildFixturesTest(unittest.TestCase):
    def setUp(self):
        ctx = temporary_directory()
        self.addCleanup(ctx.cleanup)
        self.root = Path(ctx.name)
        self.source = self.root / "source"
        for group in GROUPS:
            (self.source / group).mkdir(parents=True)

    def put(self, rel, text, encoding="utf-8"):
        path = self.source / rel
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(text.encode(encoding))

    def test_cards_notes_prose_and_name_collisions(self):
        self.put("公共建筑/01】规则.txt", "\n 制作人：甲（123）\n审核人：乙\n自由规则\n", "gbk")
        self.put("特质改造/05】旧特质.txt", "旧特质\r\n不带 BOM 的正文\r\n", "utf-16-be")
        self.put("能量池/02】M能量池.txt", "M能量池\r\n\r\n[能量池名称]：M能量池\r\n1D6 + 10％\r\n", "utf-16")
        self.put("职业/神秘侧/03】天师.txt", "天师\n（制作人：丙，审核人：丁）（修改人：戊）\n\n【能力修改：改为 2D6】\n【说明：旧版兼容】\n  [能力名称]：重复\n效果  \n\n【称号】\n只有天师可以获得\n\n【称号名称】：重复\n等级 C\n［道具名称］：\n保留正文\n〖单位名称〗：白鸽\n〔技能名称〕：清风\n")
        self.put("技能表/04】风.txt", "风\n[能力名称]：重复\n用 2AP\n[能力名称]：重复〔天师〕\n保留天然标题\n")
        for excluded in ("新过审序列/临时.txt", "荣誉室/下架.txt", "技能表/.obsidian/内部.txt", "职业/荣誉室/下架.txt", "特质改造/.隐藏.txt"):
            self.put(excluded, "不能导出")
        output = self.root / "out"
        manifest = builder.build(self.source, output, "test")
        self.assertEqual(manifest["file_count"], 5)
        self.assertEqual(load(output / "特质改造/05】旧特质.json")["helpdoc"], {"旧特质": "旧特质\r\n不带 BOM 的正文"})
        rule = load(output / "公共建筑/01】规则.json")
        self.assertEqual(rule["author"], "甲（123）")
        self.assertEqual(rule["helpdoc"], {"规则": "制作人：甲（123）\n审核人：乙\n自由规则"})
        pool = load(output / "能量池/02】M能量池.json")
        self.assertEqual(pool["helpdoc"]["M能量池"], "[能量池名称]：M能量池\r\n1D6 + 10％")
        self.assertIn("M能量池（总览）", pool["helpdoc"])
        job = load(output / "职业/神秘侧/03】天师.json")
        self.assertEqual(job["author"], "丙、戊")
        self.assertIn("【称号】\n只有天师可以获得", job["helpdoc"]["天师"])
        self.assertIn("【能力修改：改为 2D6】\n【说明：旧版兼容】\n  [能力名称]：重复\n效果  ", job["helpdoc"].values())
        self.assertIn("重复〔天师〕（#2）", job["helpdoc"])
        self.assertIn("天师·未命名条目", job["helpdoc"])
        self.assertIn("白鸽", job["helpdoc"])
        self.assertIn("清风", job["helpdoc"])
        wind = load(output / "技能表/04】风.json")
        self.assertIn("重复〔风〕", wind["helpdoc"])
        self.assertIn("重复〔天师〕", wind["helpdoc"])
        self.assertFalse(job["helpdoc"].keys() & wind["helpdoc"].keys())

    def test_gate_and_archive_destination_fail_without_export(self):
        for group in GROUPS:
            self.put(f"{group}/001】{group}.txt", "测试")
        output = self.root / "out"
        with self.assertRaisesRegex(ValueError, "低于门禁"):
            builder.build(self.source, output, "test", check=True)
        result = run("build_sealdice_helpdoc.py", "--source-dir", self.source, "--output-dir", output,
                     "--archive", output / "bad.tar.gz")
        self.assertNotEqual(result.returncode, 0)
        self.assertFalse(output.exists())


class SealDiceMock(ThreadingHTTPServer):
    def __init__(self, token):
        super().__init__(("127.0.0.1", 0), SealDiceHandler)
        self.token = token
        self.reloads = 0
        self.pages = []
        self.tree_calls = 0
        self.loading = 0
        self.busy = False
        self.failure = None
        self.authorized = []


class SealDiceHandler(BaseHTTPRequestHandler):
    def log_message(self, *args):
        pass

    def do_GET(self):
        self.respond()

    def do_POST(self):
        self.respond()

    def respond(self):
        s = self.server
        body = self.rfile.read(int(self.headers.get("Content-Length", 0)))
        authorized = self.headers.get("token") == s.token
        s.authorized.append(authorized)
        if not authorized:
            answer = {"result": False, "err": "token rejected"}
        elif self.command == "POST" and self.path == "/sd-api/helpdoc/reload":
            s.reloads += 1
            if s.busy and s.reloads == 1:
                s.loading = 1
                answer = {"result": False, "err": "正在重新装载"}
            elif s.failure == "reload" and s.reloads == 1:
                answer = {"result": False, "err": "do not log this response: " + s.token}
            else:
                s.loading = 1
                answer = {"result": True}
        elif self.command == "GET" and self.path == "/sd-api/helpdoc/tree":
            s.tree_calls += 1
            if s.loading:
                s.loading -= 1
                answer = {"result": True, "loading": True, "data": []}
            else:
                answer = {"result": True, "loading": False, "data": [
                    {"name": group, "group": group, "isDir": True, "children": [
                        {"name": f"001】{group}.json", "group": group, "isDir": False,
                         "loadStatus": 2 if s.failure == "tree" and s.reloads == 1 and group == "技能表" else 1}
                    ]} for group in GROUPS]}
        elif self.command == "POST" and self.path == "/sd-api/helpdoc/text_item_page":
            request = json.loads(body)
            s.pages.append(request)
            answer = {"result": True, "total": 0 if s.failure == "page" and s.reloads == 1 and request["group"] == "技能表" else 1, "data": []}
        else:
            self.send_error(404)
            return
        data = json.dumps(answer).encode()
        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.end_headers()
        self.wfile.write(data)


class SyncTest(unittest.TestCase):
    TOKEN = "sample+#token-NOT-FOR-LOGS"

    def setUp(self):
        tmp = temporary_directory()
        self.addCleanup(tmp.cleanup)
        self.root = Path(tmp.name)
        self.sealdice = self.root / "sealdice"
        self.helpdoc = self.sealdice / "data/helpdoc"
        self.helpdoc.mkdir(parents=True)
        (self.helpdoc / "自定义").mkdir()
        (self.helpdoc / "自定义/my.json").write_text("unmanaged")
        (self.helpdoc / "help_config.yaml").write_text("keep settings")
        (self.helpdoc / "宏观界域.zip").write_bytes(b"old ZIP")
        (self.helpdoc / "职业").mkdir()
        (self.helpdoc / "职业/old.json").write_text("old file")
        (self.helpdoc / MANIFEST).write_text('{"commit":"old"}')
        (self.sealdice / "data/dice.yaml").write_text(f"unrelated: ignored\naccessTokens:\n- '{self.TOKEN}' # comment\n- other\nlast: value\n")
        self.bundle = self.root / "bundle.tar.gz"
        self.source = self.root / "source"
        for group in GROUPS:
            (self.source / group).mkdir(parents=True)
            (self.source / group / f"001】{group}.txt").write_text(f"{group}\n[能力名称]：{group}测试\n正文", encoding="utf-8")
        staging = self.root / "out"
        builder.build(self.source, staging, "mock-commit")
        builder.archive_export(staging, self.bundle)
        self.server = SealDiceMock(self.TOKEN)
        thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        thread.start()
        self.addCleanup(self.server.server_close)
        self.addCleanup(thread.join)
        self.addCleanup(self.server.shutdown)

    def deploy(self, *flags, bundle=None, source=False):
        origin = ("--source-dir", self.source) if source else ("--bundle", bundle or self.bundle)
        result = run("sync_sealdice_helpdoc.py", *origin, "--sealdice-dir", self.sealdice,
                     "--api-url", f"http://127.0.0.1:{self.server.server_port}", "--clean-legacy-zip", *flags, no_yaml=True)
        self.assertNotIn(self.TOKEN, result.stdout + result.stderr)
        return result

    def assert_unmanaged(self):
        self.assertEqual((self.helpdoc / "自定义/my.json").read_text(), "unmanaged")
        self.assertEqual((self.helpdoc / "help_config.yaml").read_text(), "keep settings")

    def assert_rolled_back(self):
        self.assertEqual((self.helpdoc / "职业/old.json").read_text(), "old file")
        self.assertEqual(load(self.helpdoc / MANIFEST)["commit"], "old")
        self.assertEqual((self.helpdoc / "宏观界域.zip").read_bytes(), b"old ZIP")
        self.assertFalse((self.helpdoc / "技能表").exists())
        self.assert_unmanaged()

    def test_success_yaml_fallback_idempotency_and_force(self):
        result = self.deploy()
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertGreaterEqual(self.server.tree_calls, 2)
        self.assertEqual(self.server.reloads, 1)
        self.assertGreater(len(self.server.authorized), 0)
        self.assertTrue(all(self.server.authorized))
        self.assertEqual({item["group"] for item in self.server.pages}, set(GROUPS))
        self.assertTrue(all(p == {"group": p["group"], "pageNum": 1, "pageSize": 5} for p in self.server.pages))
        self.assertEqual(load(self.helpdoc / MANIFEST)["commit"], "mock-commit")
        self.assertFalse((self.helpdoc / "宏观界域.zip").exists())
        backup = self.sealdice / "backups/helpdoc-seqlib-prev"
        self.assertEqual((backup / "宏观界域.zip").read_bytes(), b"old ZIP")
        self.assertEqual((backup / "职业/old.json").read_text(), "old file")
        self.assert_unmanaged()
        skip = self.deploy()
        self.assertEqual(skip.returncode, 0, skip.stderr)
        self.assertEqual(self.server.reloads, 1)
        force = self.deploy("--force")
        self.assertEqual(force.returncode, 0, force.stderr)
        self.assertEqual(self.server.reloads, 2)
        self.assertEqual(load(backup / MANIFEST)["commit"], "mock-commit")

    def test_search_failure_rolls_back(self):
        self.server.failure = "page"
        self.assertNotEqual(self.deploy().returncode, 0)
        self.assertEqual(self.server.reloads, 2)
        self.assert_rolled_back()

    def test_reload_failure_rolls_back(self):
        self.server.failure = "reload"
        self.assertNotEqual(self.deploy().returncode, 0)
        self.assertEqual(self.server.reloads, 2)
        self.assert_rolled_back()

    def test_tree_failure_rolls_back(self):
        self.server.failure = "tree"
        self.assertNotEqual(self.deploy().returncode, 0)
        self.assertEqual(self.server.reloads, 2)
        self.assert_rolled_back()

    def test_busy_reload_waits_and_retries(self):
        self.server.busy = True
        result = self.deploy()
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(self.server.reloads, 2)
        self.assertEqual(len(self.server.pages), 5)

    def test_source_dir_skip_reload(self):
        result = self.deploy("--skip-reload", source=True)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(self.server.reloads, 0)
        self.assertEqual(load(self.helpdoc / MANIFEST)["file_count"], 5)
        self.assert_unmanaged()

    def test_bundle_digest_mismatch_rejected_before_touching_helpdoc(self):
        bad = self.root / "tampered.tar.gz"
        with tarfile.open(self.bundle) as original, tarfile.open(bad, "w:gz") as archive:
            for member in original:
                if not member.isfile():
                    archive.addfile(member)
                    continue
                with original.extractfile(member) as source:
                    data = source.read()
                if member.name.startswith("技能表/"):
                    data += b" "  # Still valid JSON, but no longer what the runner built.
                member.size = len(data)
                archive.addfile(member, io.BytesIO(data))
        result = self.deploy(bundle=bad)
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("SHA-256", result.stderr)
        self.assertEqual(self.server.reloads, 0)
        self.assert_rolled_back()

    def test_bundle_traversal_rejected_before_touching_helpdoc(self):
        bad = self.root / "bad.tar.gz"
        with tarfile.open(bad, "w:gz") as archive:
            member = tarfile.TarInfo("../unwanted")
            member.size = 1
            archive.addfile(member, io.BytesIO(b"x"))
        self.assertNotEqual(self.deploy(bundle=bad).returncode, 0)
        self.assertEqual(self.server.reloads, 0)
        self.assert_rolled_back()


class WorkflowTest(unittest.TestCase):
    def test_deployment_guardrails(self):
        workflow = (ROOT / ".github/workflows/sync-sealdice-helpdoc.yml").read_text()
        for marker in ("main", '"v*"', "workflow_dispatch", "!序列库/新过审序列/**", "cancel-in-progress: false",
                       "group: sync-sealdice-helpdoc", "force_reload", "actions/upload-artifact@v4",
                       "lint_resources.py", "--all --strict", "--check", "--clean-legacy-zip",
                       "GITHUB_STEP_SUMMARY", "DEPLOY_SSH_KEY"):
            self.assertIn(marker, workflow)
        self.assertRegex(workflow, r'(?m)^\s+scp [^\n]*-P "\$PORT"')


if __name__ == "__main__":
    unittest.main()
