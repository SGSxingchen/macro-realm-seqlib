from web.backend.app.search import SearchIndex


def test_creator_facets_and_filters_exclude_review_only_credits(tmp_path):
    resources = tmp_path / "序列库" / "职业"
    resources.mkdir(parents=True)
    manuscripts = {
        "001】甲.txt": "001】甲\n\n（制作人：甲）（审核人：乙）\n（复查人：丙）\n正文",
        "002】乙.txt": "002】乙\n\n（审核人：甲）（制作人：乙）\n（修改人：乙）\n正文",
        "003】丙.txt": "003】丙\n\n（审核人：丙）\n正文",
    }
    for filename, content in manuscripts.items():
        (resources / filename).write_text(content, encoding="utf-8")

    index = SearchIndex(tmp_path, ("序列库",))
    result = index.search("")
    assert result["facets"]["authors"] == [
        {"name": "乙", "count": 1}, {"name": "甲", "count": 1},
    ]
    for creator, filename in (("甲", "001】甲.txt"), ("乙", "002】乙.txt")):
        filtered = index.search("", authors=[creator], include_content=True)
        assert filtered["count"] == 1
        assert filtered["items"][0]["filename"] == filename
        assert filtered["items"][0]["content"] == manuscripts[filename]
    assert index.search("", authors=["丙"])["count"] == 0


def test_creator_aliases_deduplicate_without_merging_other_people(tmp_path):
    # 同一资源多次署名只计一次；旧链接仍可筛选；审核与帝国不得串入龙娘。
    resources = tmp_path / "序列库" / "职业"
    resources.mkdir(parents=True)
    manuscripts = {
        "001】甲.txt": "001】甲\n\n（制作人：龙、1828女士、审判与正义之龙）\n（调整人：糕、糕子、糕手）\n原文",
        "002】乙.txt": "002】乙\n\n（制作人：帝国）\n（审核人：龙娘）\n原文",
        "003】丙.txt": "003】丙\n\n（制作人：奥莱特希雅·坦格利安·乔斯琳[1828449440]）\n原文",
        "004】丁.txt": "004】丁\n\n（制作人：沧羽（853304398））\n原文",
    }
    for filename, content in manuscripts.items():
        (resources / filename).write_text(content, encoding="utf-8")
    index = SearchIndex(tmp_path, ("序列库",))
    counts = {item["name"]: item["count"] for item in index.search("")["facets"]["authors"]}
    assert counts == {"龙娘": 2, "糕手": 1, "帝国": 1, "沧羽": 1}
    for alias in ("龙娘", "龙", "1828女士", "审判与正义之龙"):
        result = index.search("", authors=[alias], include_content=True)
        assert result["total"] == 2
        assert all(item["content"] == manuscripts[item["filename"]] for item in result["items"])
    assert index.search("", authors=["帝国"])["total"] == 1
    assert index.search("", authors=["未知作者"])["total"] == 0
