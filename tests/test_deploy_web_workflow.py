from pathlib import Path
import unittest


REPO_ROOT = Path(__file__).resolve().parents[1]
WORKFLOW = REPO_ROOT / ".github" / "workflows" / "deploy-web.yml"


class DeployWebWorkflowTest(unittest.TestCase):
    def test_main_and_version_tags_trigger_deploy(self):
        text = WORKFLOW.read_text(encoding="utf-8")

        self.assertIn("branches:", text)
        self.assertIn("- main", text)
        self.assertIn("tags:", text)
        self.assertIn('- "v*"', text)

    def test_remote_checkout_fetches_tags_before_pulling_main(self):
        text = WORKFLOW.read_text(encoding="utf-8")

        fetch = "git fetch --force --tags origin"
        pull = "git pull --ff-only origin main"
        self.assertIn(fetch, text)
        self.assertIn(pull, text)
        self.assertLess(text.index(fetch), text.index(pull))


if __name__ == "__main__":
    unittest.main()
