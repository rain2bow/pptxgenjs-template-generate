function sampleSpec(style = 'swiss') {
  if (style === 'cmb') {
    return {
      title: '招商银行经营汇报',
      subtitle: 'Business Review / Customer Operation / Risk Control',
      author: 'China Merchants Bank',
      style: 'cmb',
      theme: 'classic',
      logoHeader: 'logos/cmb-logo-lockup.png',
      logoMark: 'logos/cmb-logo-mark.svg',
      headY: 1.06,
      slides: [
        { layout: 'cover', kicker: 'CHINA MERCHANTS BANK / 2026', title: '招商银行业务增长与数字化经营汇报', subtitle: '围绕客户经营、风险控制与效率提升的阶段性复盘' },
        { layout: 'statement', allowEmptyMediaSlots: true, kicker: 'Executive Summary', title: '以客户价值为核心，形成增长、风控与效率的闭环。', body: '本页用于放置全篇核心判断。建议用一句明确结论加一行依据，不要写成普通目录。', callout: '稳健经营\n价值增长\n风险可控' },
        { layout: 'kpiTower', kicker: 'Performance Snapshot', title: '关键经营指标保持稳健改善', items: [
          { label: '零售客户增长', value: '+18%', valueNum: 18, note: '同比保持双位数增长。' },
          { label: '活跃客户提升', value: '+32%', valueNum: 32, note: '数字渠道贡献主要增量。' },
          { label: '风险预警覆盖', value: '96%', valueNum: 96, note: '重点客群实现全面监测。' },
          { label: '流程时效缩短', value: '-24%', valueNum: 24, note: '自动化处置压缩等待时间。' },
        ] },
        { layout: 'dashboard', kicker: 'Risk & Efficiency Dashboard', title: '风险与效率指标联动监控', metrics: [
          { label: '风险命中', value: '87%' },
          { label: '处置时效', value: '2.4h' },
          { label: '自动化率', value: '68%' },
          { label: '复核准确', value: '95%' },

        ], charts: [
          { chartType: 'column', title: '风险事件处置量', labels: ['1月', '2月', '3月', '4月'], values: [120, 148, 176, 214] },
          { chartType: 'doughnut', title: '渠道结构', labels: ['App', '网点', '远程', '客户经理'], values: [46, 18, 21, 15], showLegend: true },
        ] },
        { layout: 'briefing', kicker: 'Operating Brief', title: 'Readable operating brief.', summary: 'Dense material is grouped into summary, actions and takeaway for review.', sections: [
          { title: 'Segments', body: 'Segment by assets, behavior and lifecycle. Match each group with an offer and owner.' },
          { title: 'Signal capture', body: 'Capture high-value moments, churn signals and service friction across key channels.' },
          { title: 'Closed-loop action', body: 'Assign owners, trigger actions and review conversion weekly.' },
          { title: 'Risk coordination', body: 'Combine behavior alerts and review thresholds. Escalate only material exceptions.' },
        ], conclusion: 'Use when ordinary columns are too fragmented.' },
        { layout: 'media', kicker: 'Customer Operation', title: '客户经营从单点触达转向分层运营', body: '通过客群分层、权益匹配与渠道协同，提升客户转化与长期价值。', chart: { chartType: 'line', title: '客户活跃趋势', labels: ['Q1', 'Q2', 'Q3', 'Q4'], values: [42, 51, 63, 78], showValue: true }, items: [
          { icon: 'users', title: '客群分层', body: '按资产、行为与生命周期拆分运营策略。' },
          { icon: 'scan-search', title: '信号识别', body: '捕捉高价值触点与潜在流失风险。' },
          { icon: 'workflow', title: '渠道协同', body: '联动 App、网点、远程服务与客户经理。' },
        ] },
        { layout: 'roadmap', kicker: 'Implementation Roadmap', title: '下一阶段推进路径', steps: [
          { label: '01', title: '统一指标', body: '明确经营、风险和体验的核心口径。' },
          { label: '02', title: '试点验证', body: '选择重点客群和重点分行进行闭环验证。' },
          { label: '03', title: '规模推广', body: '沉淀可复制流程并接入自动化运营。' },
          { label: '04', title: '持续评估', body: '按月复盘指标并优化模型和策略。' },
        ] },
        { layout: 'closing', kicker: 'THANK YOU', title: '稳健经营，持续创造客户价值。', subtitle: 'CHINA MERCHANTS BANK' },
      ],
    };
  }
  if (style === 'magazine') {
    return {
      title: '杂志风 PPTX 样例',
      subtitle: '电子杂志 × 电子墨水',
      author: 'Presentation Team',
      style: 'magazine',
      theme: 'ink',
      slides: [
        {
          layout: 'cover',
          kicker: 'A Talk · 2026',
          title: '一人公司',
          subtitle: '被 AI 折叠的组织',
        },
        {
          layout: 'bigNumbers',
          kicker: 'Proof',
          title: '过去 64 天',
          subtitle: '先看数字，再谈方法。',
          items: [
            { label: 'Duration', value: '64', unit: '天', note: '从 0 到现在' },
            { label: 'Lines', value: '110K+', note: '代码规模' },
            { label: 'Commits', value: '608', note: '持续迭代' },
            { label: 'Platforms', value: '9', note: '多平台分发' },
            { label: 'Providers', value: '19', note: '跨模型接入' },
            { label: 'Stars', value: '5K+', note: '开源反馈' },
          ],
        },
        {
          layout: 'quoteImage',
          allowEmptyMediaSlots: true,
          kicker: 'But',
          title: '我不是程序员。',
          body: '过去十年做的是 UI 设计和 AI 特效。新的工具链让能力边界被重新折叠。',
          callout: '这东西在三年前，需要一个十人团队做一年。',
          image: { caption: 'Image slot · 16:10' },
        },
        {
          layout: 'pipeline',
          kicker: 'Workflow',
          title: '两条流水线',
          steps: [
            { title: 'Draft', desc: '起草初稿' },
            { title: 'Polish', desc: '润色去 AI 味' },
            { title: 'Morph', desc: '改写成多平台版本' },
            { title: 'Illustrate', desc: '生成配图' },
            { title: 'Distribute', desc: '分发与复盘' },
          ],
        },
        {
          layout: 'bigQuote',
          kicker: 'Takeaway',
          quote: '技能变厚，组织变薄。',
          body: 'Thin harness, fat skills.',
        },
      ],
    };
  }

  return {
    title: '一人公司：被 AI 折叠的组织',
    subtitle: '一个关于 AI、组织和个体的新叙事',
    author: 'Presentation Team',
    style: 'swiss',
    theme: 'ikb',
    slides: [
      {
        layout: 'cover',
        kicker: 'Swiss Field Note · AI Organization',
        title: '一人公司',
        subtitle: '被 AI 折叠的组织',
      },
      {
        layout: 'statement',
        allowEmptyMediaSlots: true,
        kicker: 'Thesis · 01',
        title: '组织正在从人数问题，变成接口问题。',
        body: '当个体能调用足够厚的技能层，公司边界会重新定义。',
      },
      {
        layout: 'kpiTower',
        kicker: 'Proof · Numbers',
        title: '64 天里的可见结果',
        items: [
          { label: 'Duration', value: '64d', valueNum: 64 },
          { label: 'Lines', value: '110K+', valueNum: 110 },
          { label: 'Platforms', value: '9', valueNum: 9 },
          { label: 'Commits', value: '608', valueNum: 608 },
        ],
      },
      {
        layout: 'fourCards',
        kicker: 'Mechanism · 04',
        title: '四个被压缩的组织环节',
        items: [
          { title: '研究', desc: '从资料检索到结构化判断。' },
          { title: '设计', desc: '从审美草图到可执行规范。' },
          { title: '开发', desc: '从想法描述到代码实现。' },
          { title: '分发', desc: '从单篇内容到多平台适配。' },
        ],
      },
      {
        layout: 'duoCompare',
        kicker: 'Before / After',
        title: '从交接到共建',
        before: {
          label: 'Before',
          title: '设计 → 开发 → 交接',
          items: ['文件传递', '像素还原', '反复解释'],
        },
        after: {
          label: 'After',
          title: '同工具 · 并行 · 共建',
          items: ['共享上下文', '代理执行', '快速收敛'],
        },
      },
      {
        layout: 'closing',
        kicker: 'Takeaway',
        title: '技能变厚，组织变薄。',
        subtitle: 'Thin harness, fat skills.',
      },
    ],
  };
}

module.exports = { sampleSpec };

