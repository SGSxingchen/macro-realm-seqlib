import { Block } from './parser';

const FIELD_PALETTE: Record<string, string> = {
  '能力效果': 'effect',
  '称号效果': 'effect',
  '道具效果': 'effect',
  '能量池效果': 'effect',
  '技艺效果': 'effect',
  '模块效果': 'effect',
  '奖励效果': 'effect',
  '魔药效果': 'effect',
  '称号解锁条件': 'requirement',
  '解锁条件': 'requirement',
  '开放条件': 'requirement',
  '兑换/消耗条件': 'requirement',
  '获取资格': 'requirement',
  '所需仪式': 'requirement',
  '所需材料': 'requirement',
  '晋升条件': 'requirement',
  '初级效果': 'effect',
  '中级效果': 'effect',
  '高级效果': 'effect',
  '能力简介': 'intro',
  '称号简介': 'intro',
  '道具简介': 'intro',
  '能量池简介': 'intro',
  '技艺简介': 'intro',
  '模块简介': 'intro',
  '能力形容': 'flavor',
  '释放类型': 'kind',
  '打击类型': 'kind',
  '伤害类型': 'kind',
  '段位/等级': 'kind',
  '魔药等级': 'kind',
  '所属序列': 'kind',
  '消耗能量': 'cost',
  '消耗/耐久': 'cost',
  '消耗规则': 'cost',
  '冷却时间': 'cooldown',
  '技能冷却': 'cooldown',
  '持续时间': 'duration',
  '恢复方式': 'duration',
  '补充说明': 'note',
  '基础限制': 'note',
  '冥想规则': 'note',
  '技艺栏限制': 'note',
  '风险/副作用': 'note',
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

function cardTypeClass(cardType: string) {
  const map: Record<string, string> = {
    '称号': 'title',
    '技艺': 'art',
    '道具': 'item',
    '能量池': 'pool',
    '公共建筑': 'building',
    '模块': 'module',
    '奖励': 'reward',
    '魔药': 'potion',
    '序列': 'sequence',
  };
  return map[cardType] || 'generic';
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
    case 'ability-note':
      return <div className="ability-note">{block.text}</div>;
    case 'level-tree':
      return <LevelTree block={block} />;
    case 'rule-section':
      return <RuleSection block={block} />;
    case 'paragraph':
      return <p className="doc-para">{block.text}</p>;
    case 'ability':
      return <Ability block={block} />;
    case 'special-card':
      return <SpecialCard block={block} />;
    default:
      return null;
  }
}

function LevelTree({ block }: { block: Extract<Block, { kind: 'level-tree' }> }) {
  return (
    <div className="level-tree" aria-label="升级树">
      {block.items.map((item) => (
        <div key={`${item.level}-${item.text}`} className={`level-tree-item ${levelClass(item.level)}`}>
          <span className="level-tree-badge">{item.level}</span>
          <p>{item.text}</p>
        </div>
      ))}
    </div>
  );
}

function RuleSection({ block }: { block: Extract<Block, { kind: 'rule-section' }> }) {
  return (
    <section className="rule-card">
      <header className="rule-card-head">
        <span>规则</span>
        <h2>{block.title}</h2>
      </header>
      <div className="rule-card-body">
        {block.lines.map((line, i) => <p key={i}>{line}</p>)}
      </div>
    </section>
  );
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

function SpecialCard({ block }: { block: Extract<Block, { kind: 'special-card' }> }) {
  return (
    <div className={`ability-card special-card special-${cardTypeClass(block.cardType)}`}>
      <header className="ability-header">
        <div className="ability-titles">
          <span className="special-card-kind">{block.cardType}</span>
          <h3>{block.name}</h3>
        </div>
      </header>
      <dl className="ability-fields">
        {block.fields.map((f, i) => (
          <div key={i} className={`ability-field f-${FIELD_PALETTE[f.key] || 'default'}`}>
            <dt>{f.key}</dt>
            <dd>{f.value}</dd>
          </div>
        ))}
      </dl>
    </div>
  );
}
