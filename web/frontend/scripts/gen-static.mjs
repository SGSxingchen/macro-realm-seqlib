#!/usr/bin/env node
/**
 * 构建期静态页生成器（无第三方依赖，只用 Node 内置模块）。
 *
 * 在 `vite build` 之后运行，把 `序列库/` 下的全部 .txt 条目生成可直接被抓取的
 * 静态 HTML，另外生成分类目录页、sitemap.xml 与 robots.txt。
 *
 * 产物：
 *   dist/entry/<路径分段用 __ 连接>.html
 *   dist/catalog.html
 *   dist/sitemap.xml
 *   dist/robots.txt
 *
 * 说明：站点使用 Vite 多入口 SPA，运行时没有 URL 路由；因此这里把条目的
 * `分类/文件名` 扁平化为单个文件名（各路径分段用 __ 连接），这样：
 *   1. `dist/entry/` 下的文件数就等于条目数，便于核对；
 *   2. 不需要在 URL 里出现 %2F（nginx 默认会把 %2F 当路径分隔符，容易 404）。
 * 中文文件名由服务器按 UTF-8 提供，页面里的链接 / canonical / sitemap 统一做
 * 逐段 URL 编码。
 */

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const SCRIPT_DIR = path.dirname(fileURLToPath(import.meta.url));
const FRONTEND_DIR = path.resolve(SCRIPT_DIR, '..');
const REPO_ROOT = path.resolve(FRONTEND_DIR, '..', '..');
const LIBRARY_DIR = path.join(REPO_ROOT, '序列库');
const DIST_DIR = path.join(FRONTEND_DIR, 'dist');
const ENTRY_DIR = path.join(DIST_DIR, 'entry');

const SITE_ORIGIN = 'https://trpg.chordvers.org';
const SITE_NAME = '宏观界域强化序列库';

function fail(message) {
  console.error(`[gen-static] ${message}`);
  process.exit(1);
}

/** 递归查找 .txt（跳过 .obsidian 等隐藏目录），返回相对 序列库 的 posix 路径。 */
function collectEntries(dir) {
  const out = [];
  const visit = (current, relParts) => {
    const dirents = fs.readdirSync(current, { withFileTypes: true });
    dirents.sort((a, b) => a.name.localeCompare(b.name, 'zh-Hans-CN'));
    for (const dirent of dirents) {
      if (dirent.name.startsWith('.')) continue;
      const abs = path.join(current, dirent.name);
      const rel = [...relParts, dirent.name];
      if (dirent.isDirectory()) {
        visit(abs, rel);
      } else if (dirent.isFile() && dirent.name.endsWith('.txt')) {
        out.push(rel.join('/'));
      }
    }
  };
  visit(dir, []);
  return out;
}

function decodeText(buffer) {
  if (buffer.length >= 3 && buffer[0] === 0xef && buffer[1] === 0xbb && buffer[2] === 0xbf) {
    return buffer.subarray(3).toString('utf8');
  }
  if (buffer.length >= 2 && buffer[0] === 0xff && buffer[1] === 0xfe) {
    return buffer.subarray(2).toString('utf16le');
  }
  if (buffer.length >= 2 && buffer[0] === 0xfe && buffer[1] === 0xff) {
    const swapped = Buffer.from(buffer.subarray(2));
    for (let i = 0; i + 1 < swapped.length; i += 2) {
      const tmp = swapped[i];
      swapped[i] = swapped[i + 1];
      swapped[i + 1] = tmp;
    }
    return swapped.toString('utf16le');
  }
  return buffer.toString('utf8');
}

function escapeHtml(value) {
  return String(value)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function escapeXml(value) {
  return String(value)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&apos;');
}

/** 每个路径分段单独做 URL 编码；本项目的条目文件是扁平名，故只有一段。 */
function encodeUrlPath(relativePath) {
  return relativePath
    .split('/')
    .map((segment) => encodeURIComponent(segment))
    .join('/');
}

function entryFileName(relPath) {
  return `${relPath.slice(0, -'.txt'.length).split('/').join('__')}.html`;
}

function entryUrlPath(relPath) {
  return `/entry/${encodeUrlPath(entryFileName(relPath))}`;
}

function entryCanonical(relPath) {
  return `${SITE_ORIGIN}${entryUrlPath(relPath)}`;
}

function cleanTitle(raw, fallback) {
  const title = raw.replace(/^\uFEFF/, '').trim();
  if (!title) return fallback;
  if (/^(投稿人|审核人|制作人|作者|更新|修订|版本)\s*[:：]/.test(title)) return fallback;
  return title;
}

function descriptionOf(text, title) {
  const lines = text.replace(/\r\n?/g, '\n').split('\n');
  const body = lines.filter((line) => line.trim() && line.trim() !== title).join(' ');
  const compact = body.replace(/\s+/g, ' ').trim();
  if (compact.length <= 160) return compact;
  return `${Array.from(compact).slice(0, 157).join('')}…`;
}

function renderBody(text, title, skipFirst) {
  const normalized = text.replace(/\r\n?/g, '\n');
  const lines = normalized.split('\n');
  const start = skipFirst ? 1 : 0;
  const body = lines.slice(start).join('\n').replace(/^\n+/, '');
  if (!body.trim()) return '<p class="entry-empty">（本条目暂无正文。）</p>';
  const blocks = body
    .split(/\n[ \t]*\n+/)
    .map((block) => block.replace(/[ \t]+$/gm, '').trim())
    .filter((block) => block.length > 0);
  return blocks
    .map((block) => `<p>${escapeHtml(block).replaceAll('\n', '<br />\n')}</p>`)
    .join('\n');
}

function pageShell({ title, description, canonical, bodyClass, body }) {
  return `<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>${escapeHtml(title)}</title>
<meta name="description" content="${escapeHtml(description)}" />
<link rel="canonical" href="${escapeHtml(canonical)}" />
<meta name="robots" content="index,follow" />
<link rel="alternate" type="application/xml" title="站点地图" href="/sitemap.xml" />
<style>
:root { color-scheme: light dark; }
* { box-sizing: border-box; }
body { margin: 0; background: #f7f4ee; color: #25211c; font: 16px/1.85 -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Hiragino Sans GB", "Microsoft YaHei", sans-serif; }
a { color: #7a3e1d; text-underline-offset: 3px; }
a:hover { color: #a85a2a; }
.site-header { border-bottom: 1px solid #e2d9cb; background: #fffdf9; }
.site-header-inner { max-width: 900px; margin: 0 auto; padding: 14px 20px; display: flex; flex-wrap: wrap; gap: 8px 16px; align-items: baseline; }
.site-name { font-weight: 700; color: #3b2c20; text-decoration: none; }
.site-nav { font-size: 14px; }
main { max-width: 900px; margin: 0 auto; padding: 28px 20px 72px; }
.breadcrumb { color: #776a5c; font-size: 14px; margin-bottom: 12px; }
h1 { font-size: clamp(24px, 4vw, 34px); line-height: 1.35; margin: 0 0 12px; color: #3b2c20; }
.entry-meta { color: #776a5c; font-size: 14px; margin: 0 0 26px; }
.entry-content p { margin: 0 0 1.05em; overflow-wrap: anywhere; }
.entry-empty { color: #776a5c; }
.catalog-intro { color: #5b4b3c; margin: 0 0 24px; }
.category-index { columns: 2 260px; gap: 32px; margin: 0 0 36px; padding: 0; list-style: none; }
.category-index li { break-inside: avoid; margin: 0 0 8px; }
.category-index .count { color: #8a7a68; font-size: 13px; }
.category { margin: 0 0 36px; scroll-margin-top: 16px; }
.category h2 { font-size: 20px; line-height: 1.5; margin: 0 0 4px; color: #3b2c20; }
.category .category-path { color: #8a7a68; font-size: 14px; margin: 0 0 10px; }
.entry-list { margin: 0; padding: 0; list-style: none; border-top: 1px solid #e2d9cb; }
.entry-list li { border-bottom: 1px solid #e2d9cb; }
.entry-list a { display: block; padding: 8px 2px; text-decoration: none; }
.entry-list a:hover { background: #fffdf9; }
.site-footer { max-width: 900px; margin: 0 auto; padding: 18px 20px 40px; color: #8a7a68; font-size: 13px; border-top: 1px solid #e2d9cb; }
@media (prefers-color-scheme: dark) {
  body { background: #17130f; color: #e8e0d6; }
  .site-header { background: #1d1813; border-bottom-color: #3a2f24; }
  .site-name, h1, .category h2 { color: #f0e6d9; }
  .breadcrumb, .entry-meta, .catalog-intro, .site-footer, .category-path, .category-index .count { color: #b4a694; }
  a { color: #e2a36f; }
  a:hover { color: #f3c294; }
  .entry-list, .entry-list li { border-color: #3a2f24; }
  .entry-list a:hover { background: #221b15; }
}
</style>
</head>
<body class="${escapeHtml(bodyClass)}">
<header class="site-header">
  <div class="site-header-inner">
    <a class="site-name" href="/">${escapeHtml(SITE_NAME)}</a>
    <nav class="site-nav"><a href="/catalog.html">全部条目</a></nav>
  </div>
</header>
<main>
${body}
</main>
<footer class="site-footer">本页由构建期静态生成，正文来自序列库原文。</footer>
</body>
</html>
`;
}

function renderEntryPage(entry, catalogAnchor) {
  const { relPath, title, text, category, firstLineIsTitle } = entry;
  const description = descriptionOf(text, title);
  const body = `<div class="breadcrumb"><a href="/">首页</a> / <a href="/catalog.html">全部条目</a>${category ? ` / <a href="/catalog.html#${catalogAnchor}">${escapeHtml(category)}</a>` : ''}</div>
<h1>${escapeHtml(title)}</h1>
<p class="entry-meta">分类：${escapeHtml(category || '未分类')} · <a href="/">返回宏观界域强化序列库</a></p>
<article class="entry-content">
${renderBody(text, title, firstLineIsTitle)}
</article>`;
  return pageShell({
    title: `${title} ｜ ${SITE_NAME}`,
    description,
    canonical: entryCanonical(relPath),
    bodyClass: 'entry-page',
    body,
  });
}

function renderCatalogPage(groups, total) {
  const indexItems = groups
    .map(
      (group) =>
        `<li><a href="#${group.anchor}">${escapeHtml(group.category || '未分类')}</a> <span class="count">${group.entries.length}</span></li>`,
    )
    .join('\n');
  const sections = groups
    .map(
      (group) => `<section class="category" id="${group.anchor}">
<h2>${escapeHtml(group.category || '未分类')}</h2>
<p class="category-path">${group.entries.length} 条</p>
<ul class="entry-list">
${group.entries
  .map((entry) => `<li><a href="${entryUrlPath(entry.relPath)}">${escapeHtml(entry.title)}</a></li>`)
  .join('\n')}
</ul>
</section>`,
    )
    .join('\n');
  const body = `<div class="breadcrumb"><a href="/">首页</a> / 全部条目</div>
<h1>全部条目</h1>
<p class="catalog-intro">共 ${total} 条，按序列库分类顺序列出。点击条目可查看构建期生成的静态全文页。</p>
<nav aria-label="分类索引"><ul class="category-index">
${indexItems}
</ul></nav>
${sections}`;
  return pageShell({
    title: `全部条目 ｜ ${SITE_NAME}`,
    description: `${SITE_NAME}全部条目目录，共 ${total} 条，按分类列出可被抓取的静态全文页。`,
    canonical: `${SITE_ORIGIN}/catalog.html`,
    bodyClass: 'catalog-page',
    body,
  });
}

function renderSitemap(entries) {
  const fixed = [
    { url: `${SITE_ORIGIN}/`, priority: '1.0' },
    { url: `${SITE_ORIGIN}/catalog.html`, priority: '0.9' },
    { url: `${SITE_ORIGIN}/tabletop.html`, priority: '0.7' },
  ];
  const fixedUrls = fixed.map(
    (item) => `  <url>\n    <loc>${escapeXml(item.url)}</loc>\n    <priority>${item.priority}</priority>\n  </url>`,
  );
  const entryUrls = entries.map(
    (entry) => `  <url>\n    <loc>${escapeXml(entryCanonical(entry.relPath))}</loc>\n    <lastmod>${entry.lastmod}</lastmod>\n  </url>`,
  );
  return `<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
${[...fixedUrls, ...entryUrls].join('\n')}
</urlset>
`;
}

function main() {
  if (!fs.existsSync(LIBRARY_DIR)) fail(`找不到序列库目录：${LIBRARY_DIR}`);
  if (!fs.existsSync(DIST_DIR)) fail(`找不到 dist 目录（请先运行 vite build）：${DIST_DIR}`);

  const relPaths = collectEntries(LIBRARY_DIR);
  if (relPaths.length === 0) fail('序列库下没有找到任何 .txt 条目');

  const collator = new Intl.Collator('zh-Hans-CN', { numeric: true, sensitivity: 'base' });
  const entries = relPaths.map((relPath) => {
    const abs = path.join(LIBRARY_DIR, relPath);
    const buffer = fs.readFileSync(abs);
    const text = decodeText(buffer);
    const normalized = text.replace(/\r\n?/g, '\n').replace(/^\uFEFF/, '');
    const lines = normalized.split('\n');
    const firstMeaningful = lines.findIndex((line) => line.trim().length > 0);
    const stem = path.posix.basename(relPath, '.txt');
    const rawFirstLine = firstMeaningful >= 0 ? lines[firstMeaningful] : '';
    const title = cleanTitle(rawFirstLine, stem);
    const firstLineIsTitle = firstMeaningful >= 0 && rawFirstLine.trim() === title;
    const dir = path.posix.dirname(relPath);
    const category = dir === '.' ? '' : dir;
    const stat = fs.statSync(abs);
    const lastmod = new Date(stat.mtimeMs).toISOString();
    return { relPath, text, title, category, firstLineIsTitle, lastmod };
  });

  // 分类顺序：先按分类路径排序，再按标题排序。
  const groupsByCategory = new Map();
  for (const entry of entries) {
    if (!groupsByCategory.has(entry.category)) groupsByCategory.set(entry.category, []);
    groupsByCategory.get(entry.category).push(entry);
  }
  const groups = [...groupsByCategory.entries()]
    .sort(([a], [b]) => collator.compare(a, b))
    .map(([category, groupEntries], index) => ({
      category,
      anchor: `cat-${index}`,
      entries: groupEntries.sort((a, b) => collator.compare(a.title, b.title) || collator.compare(a.relPath, b.relPath)),
    }));
  const anchorByCategory = new Map(groups.map((group) => [group.category, group.anchor]));

  fs.rmSync(ENTRY_DIR, { recursive: true, force: true });
  fs.mkdirSync(ENTRY_DIR, { recursive: true });
  for (const entry of entries) {
    const fileName = entryFileName(entry.relPath);
    const html = renderEntryPage(entry, anchorByCategory.get(entry.category) || '');
    fs.writeFileSync(path.join(ENTRY_DIR, fileName), html, 'utf8');
  }

  fs.writeFileSync(path.join(DIST_DIR, 'catalog.html'), renderCatalogPage(groups, entries.length), 'utf8');
  fs.writeFileSync(path.join(DIST_DIR, 'sitemap.xml'), renderSitemap(entries), 'utf8');
  fs.writeFileSync(
    path.join(DIST_DIR, 'robots.txt'),
    `User-agent: *\nAllow: /\n\nSitemap: ${SITE_ORIGIN}/sitemap.xml\n`,
    'utf8',
  );

  const generated = fs.readdirSync(ENTRY_DIR).filter((name) => name.endsWith('.html')).length;
  if (generated !== entries.length) fail(`条目页数量不一致：源 ${entries.length}，生成 ${generated}`);
  console.log(`[gen-static] 条目 ${entries.length} 条，分类 ${groups.length} 个`);
  console.log(`[gen-static] 已生成 ${generated} 个 entry 页、catalog.html、sitemap.xml、robots.txt`);
}

main();
