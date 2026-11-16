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
        "001】甲.txt": "001】甲\n\n（制作人：龙、1828女士、审判与正义之龙、很大一只的龙、大只龙、一只很大的龙、很大只的龙）\n（调整人：糕、糕子、糕手）\n原文",
        "002】乙.txt": "002】乙\n\n（制作人：帝国）\n（审核人：龙娘）\n原文",
        "003】丙.txt": "003】丙\n\n（制作人：奥莱特希雅·坦格利安·乔斯琳[1828449440]）\n原文",
        "004】丁.txt": "004】丁\n\n（制作人：沧羽（853304398））\n原文",
    }
    for filename, content in manuscripts.items():
        (resources / filename).write_text(content, encoding="utf-8")
    index = SearchIndex(tmp_path, ("序列库",))
    counts = {item["name"]: item["count"] for item in index.search("")["facets"]["authors"]}
    assert counts == {"龙娘": 2, "糕手": 1, "帝国": 1, "沧羽": 1}
    for alias in ("龙娘", "龙", "1828女士", "审判与正义之龙", "很大一只的龙", "大只龙", "一只很大的龙", "很大只的龙"):
        result = index.search("", authors=[alias], include_content=True)
        assert result["total"] == 2
        assert all(item["content"] == manuscripts[item["filename"]] for item in result["items"])
    assert index.search("", authors=["帝国"])["total"] == 1
    assert index.search("", authors=["未知作者"])["total"] == 0


def test_expanded_aliases_preserve_compound_names_and_review_boundaries(tmp_path):
    resources = tmp_path / "序列库" / "职业"
    resources.mkdir(parents=True)
    manuscript = (
        "001】别名核验\n\n（制作人：玛丽亚、怀念浮士德的玛利亚）\n"
        "（调整人：夏弥/帝国、帝国）\n（修改人：审核小凯、小凯ovo）\n"
        "（制作人：Hello, World（陈旭缘 老陈））\n"
        "（修改人：TTTT443939661、准备收拾行李的TTTT）\n"
        "（制作人：Resh 希儿·芙乐艾、亚波伦（希儿·芙乐艾））\n"
        "（审核人：龙娘、糕手）\n原文"
    )
    (resources / "001】别名核验.txt").write_text(manuscript, encoding="utf-8")
    index = SearchIndex(tmp_path, ("序列库",))
    result = index.search("", include_content=True)
    assert set(result["items"][0]["authors"]) == {
        "玛利亚", "帝国", "小凯", "Hello, World", "TTTT", "希儿·芙乐艾",
    }
    assert result["items"][0]["content"] == manuscript
    for alias in ("玛丽亚", "夏弥/帝国", "审核小凯", "Hello, World（陈旭缘 老陈）", "亚波伦"):
        assert index.search("", authors=[alias])["total"] == 1
    assert index.search("", authors=["龙娘"])["total"] == 0


def test_split_credits_keep_compound_names_and_separate_people(tmp_path):
    resources = tmp_path / "序列库" / "职业"
    resources.mkdir(parents=True)
    content = (
        "001】署名核验\n（制作人：二十三夜杀（十六夜红月、七夜绍））\n"
        "（修改人：从者化进行全部重修、所以也过来修一下蓝宝石的红月……能到T1吗？不太可能、但是T2上游大概有吧？有吗？）\n"
        "（调整人：红月；审核人：糕，龙）\n（制作人：未知、审核组、我不道啊）\n"
        "（制作人：伪龙、大凯、小凯、半条咸鱼干、咸鱼、一条很咸的鱼）\n"
        "（修改人：一般路过北京热心群众、小小的凯？、阿虎/阿龙）\n"
        "（审核人：龙娘；制作人：独立创作者）\n"
    )
    (resources / "001】署名核验.txt").write_text(content, encoding="utf-8")
    index = SearchIndex(tmp_path, ("序列库",))
    result = index.search("", include_content=True)
    assert set(result["items"][0]["authors"]) == {
        "二十三夜杀", "红月", "伪龙", "大凯", "小凯", "半条咸鱼干", "咸鱼",
        "一条很咸的鱼", "沧羽", "阿虎", "独立创作者",
    }
    assert result["items"][0]["content"] == content
    assert index.search("", authors=["龙娘"])["total"] == 0
    assert index.search("", authors=["糕手"])["total"] == 0
