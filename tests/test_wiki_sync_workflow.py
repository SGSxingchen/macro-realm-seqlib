from pathlib import Path
import unittest


REPO_ROOT = Path(__file__).resolve().parents[1]
WORKFLOW = REPO_ROOT / ".github" / "workflows" / "wiki-sync.yml"


class WikiSyncWorkflowTest(unittest.TestCase):
    def test_workflow_exposes_manual_dry_run_and_sync_modes(self):
        text = WORKFLOW.read_text(encoding="utf-8")

        self.assertIn("workflow_dispatch:", text)
        self.assertNotIn("push:", text)
        self.assertIn("description: \"同步模式\"", text)
        self.assertIn("- dry-run", text)
        self.assertIn("- sync", text)
        self.assertIn("default: dry-run", text)

    def test_workflow_supports_scope_inputs(self):
        text = WORKFLOW.read_text(encoding="utf-8")

        self.assertIn("skip_honor:", text)
        self.assertIn("filter:", text)
        self.assertIn("delay:", text)
        self.assertIn('default: "5"', text)

    def test_workflow_guards_real_sync_with_secrets(self):
        text = WORKFLOW.read_text(encoding="utf-8")

        self.assertIn("secrets.WIKI_USER", text)
        self.assertIn("secrets.WIKI_PASSWORD", text)
        self.assertIn("真实同步需要配置 GitHub Secrets", text)
        self.assertIn("DRY_RUN_USER", text)
        self.assertIn("DRY_RUN_PASSWORD", text)

    def test_workflow_invokes_existing_sync_script_with_expected_flags(self):
        text = WORKFLOW.read_text(encoding="utf-8")

        self.assertIn("wiki/sync_to_wiki.py", text)
        self.assertIn("--dry-run", text)
        self.assertIn("--skip-honor", text)
        self.assertIn("--filter", text)
        self.assertIn("--delay", text)
        self.assertIn("--source-dir", text)


if __name__ == "__main__":
    unittest.main()
