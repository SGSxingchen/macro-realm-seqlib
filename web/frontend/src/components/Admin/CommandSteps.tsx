import { CmdResult } from '../../types';

export function CommandSteps({ steps }: { steps?: CmdResult[] }) {
  if (!steps?.length) return null;
  const names = ['检查工作区', '加入发布范围', '提交 commit', '创建 tag', '推送 GitHub'];
  return (
    <div className="step-list">
      {steps.map((step, i) => (
        <details className={step.returncode === 0 ? 'step ok' : 'step bad'} key={`${i}-${step.cmd.join(' ')}`}>
          <summary><b>{names[i] || `步骤 ${i + 1}`}</b><span>{step.returncode === 0 ? '成功' : `失败 ${step.returncode}`}</span></summary>
          <small>{step.cmd.join(' ')}</small>
          {step.stdout && <pre>{step.stdout}</pre>}
          {step.stderr && <pre className="stderr">{step.stderr}</pre>}
        </details>
      ))}
    </div>
  );
}
