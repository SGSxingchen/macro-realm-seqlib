import { Block } from './parser';

const FIELD_PALETTE: Record<string, string> = {
  '能力效果': 'effect',
  '初级效果': 'effect',
  '中级效果': 'effect',
  '高级效果': 'effect',
  '能力简介': 'intro',
  '能力形容': 'flavor',
  '释放类型': 'kind',
  '打击类型': 'kind',
  '伤害类型': 'kind',
  '消耗能量': 'cost',
  '冷却时间': 'cooldown',
  '技能冷却': 'cooldown',
  '持续时间': 'duration',
  '补充说明': 'note',
};

function levelClass(level?: string) {
  switch ((level || '').toUpperCase()) {
    case 'EX': return 'lvl-ex';
    case 'S': return 'lvl-s';
    case 'A': return 'lvl-a';
    case 'B': return 'lvl-b';
    case 'C': return 'lvl-c';
    case 'D': return 'lvl-d';
    case 'E': return 'lvl-e';
    case 'F': return 'lvl-f';
    default: return '';
  }
}

export function Structured({ blocks }: { blocks: Block[] }) {
  return (
    <section className="document doc-structured">
      {blocks.map((b, i) => <BlockNode key={i} block={b} />)}
    </section>
  );
}

function BlockNode({ block }: { block: Block }) {
  switch (block.kind) {
    case 'title':
      return <h1 className="doc-title">{block.text}</h1>;
    case 'meta':
      return <p className="doc-meta">{block.text}</p>;
    case 'banner':
      return <div className="doc-banner">{block.text}</div>;
    case 'paragraph':
      return <p className="doc-para">{block.text}</p>;
    case 'ability':
      return <Ability block={block} />;
    default:
      return null;
  }
}

function Ability({ block }: { block: Extract<Block, { kind: 'ability' }> }) {
  return (
    <div className={`ability-card ${levelClass(block.level)}`}>
      <header className="ability-header">
        <div className="ability-titles">
          <h3>{block.name}</h3>
          {block.tags.length > 0 && (
            <div className="ability-tags">{block.tags.map(t => <span key={t} className="ability-tag">{t}</span>)}</div>
          )}
        </div>
        {block.level && <span className={`ability-level ${levelClass(block.level)}`}>{block.level}级</span>}
      </header>
      <dl className="ability-fields">
        {block.fields.map((f, i) => (
          <div key={i} className={`ability-field f-${FIELD_PALETTE[f.key] || 'default'}`}>
            <dt>{f.key}</dt>
            <dd>{f.value}</dd>
          </div>
        ))}
      </dl>
      {block.tail && <p className="ability-tail">{block.tail}</p>}
    </div>
  );
}
