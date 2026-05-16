#!/usr/bin/env python3
from fastapi import HTTPException
from app.main import REPO_ROOT, SEARCH_INDEX, git_changes, get_resource, get_raw, list_resources, safe_resource_path, tree

print('repo =', REPO_ROOT)
SEARCH_INDEX.refresh()
resources = list_resources(include_content=False)
assert resources['count'] > 0
assert all(not item['path'].startswith('荣誉室/') for item in resources['items'])
print('resources =', resources['count'], '(public only)')
node_items = tree()['items']
assert node_items
assert all(node['path'] != '荣誉室' for node in node_items)
print('tree ok (honor hidden)')
sample = next(item for item in resources['items'] if item['path'].endswith('.txt'))
detail = get_resource(sample['path'])
assert detail['content']
print('sample =', sample['path'], 'encoding =', detail['encoding'])

# 智能搜索基本盘：标题精确、子序列、多 token、空查询
q1 = list_resources(q=sample['title'][:4], include_content=False, limit=10)
assert q1['count'] >= 1, '标题前缀搜索应当有命中'
print('search title-prefix ok, hits =', q1['count'])

# facet 至少有 kinds
assert 'facets' in q1 and 'kinds' in q1['facets']
print('facets ok, kinds =', len(q1['facets']['kinds']))

try:
    safe_resource_path('../README.md')
    raise AssertionError('path traversal not rejected')
except HTTPException:
    print('path guard ok')
try:
    get_resource('荣誉室/_should_not_be_public.txt')
    raise AssertionError('honor resource not rejected')
except HTTPException:
    print('honor public guard ok')
try:
    get_raw('荣誉室/_should_not_be_public.txt')
    raise AssertionError('honor raw not rejected')
except HTTPException:
    print('honor raw guard ok')
changes = git_changes()
assert 'summary' in changes
print('git changes from', changes['from_ref'])
print('SMOKE OK')
