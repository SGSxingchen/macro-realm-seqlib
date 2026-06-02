from pathlib import Path
import tempfile
import unittest

from tools.renumber_resources import plan_resource_updates, apply_resource_updates


def write_resource(path: Path, text: str = "旧标题\n\n正文") -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


class RenumberResourcesTest(unittest.TestCase):
    def test_plan_renumbers_each_directory_independently(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "序列库"
            bio = source / "特质改造" / "生化改造类"
            spec = source / "特质改造" / "特化改造类"
            write_resource(bio / "001】保留.txt")
            write_resource(bio / "002】子目录" / "001】内层.txt")
            write_resource(bio / "003】后续.txt")
            write_resource(spec / "004】另一目录.txt")

            plan = plan_resource_updates(tmp_path, [source])

            renames = {(item.old_path.name, item.new_path.name) for item in plan.renames}
            self.assertNotIn(("003】后续.txt", "002】后续.txt"), renames)
            self.assertIn(("004】另一目录.txt", "001】另一目录.txt"), renames)

    def test_numbered_directories_participate_in_sibling_order(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "序列库"
            parent = source / "技能表"
            write_resource(parent / "001】成组技能" / "001】内层.txt")
            write_resource(parent / "003】单文件.txt")

            plan = plan_resource_updates(tmp_path, [source])

            renames = {(item.old_path.name, item.new_path.name) for item in plan.renames}
            self.assertIn(("003】单文件.txt", "002】单文件.txt"), renames)

    def test_dry_run_plan_does_not_change_files(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "序列库"
            resource = source / "职业" / "003】剑士.txt"
            write_resource(resource)

            plan_resource_updates(tmp_path, [source])

            self.assertTrue(resource.exists())
            self.assertEqual(resource.read_text(encoding="utf-8").splitlines()[0], "旧标题")

    def test_apply_updates_renamed_txt_first_line_to_filename(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "序列库"
            first = source / "职业" / "001】保留.txt"
            second = source / "职业" / "003】剑士.txt"
            write_resource(first, "旧标题1\n\n正文1")
            write_resource(second, "旧标题2\n\n正文2")
            plan = plan_resource_updates(tmp_path, [source])

            apply_resource_updates(plan)

            renamed = source / "职业" / "002】剑士.txt"
            self.assertFalse(second.exists())
            self.assertTrue(renamed.exists())
            self.assertEqual(first.read_text(encoding="utf-8").splitlines()[0], "001】保留")
            self.assertEqual(renamed.read_text(encoding="utf-8").splitlines()[0], "002】剑士")


if __name__ == "__main__":
    unittest.main()
