/**
 * Demo deck used by the end-to-end test and available as a ready-made example.
 *
 * It is a hand-authored CardDeck in the 1280x720 DSL that exercises every
 * element type (text, shape, svg, chart, image) against a real corporate
 * template's tokens. Content mirrors the design doc's own scenario.
 */
export const HW_DEMO_DECK = {
  deckId: 'hw-demo',
  title: '品牌智能演示：AI 驱动的企业内容生产',
  templateId: 'ppt-hw',
  slides: [
    {
      slideId: 'cover',
      background: '$colors.background',
      speakerNotes: '开场：一句话说明这是一份由 BrandDeck 自动生成、可在 PowerPoint 中直接编辑的品牌演示。',
      elements: [
        {
          type: 'shape',
          id: 'cover-accent',
          shapeType: 'rect',
          x: 80,
          y: 236,
          w: 8,
          h: 132,
          style: { fill: '$colors.primary' },
        },
        {
          type: 'text',
          id: 'cover-kicker',
          role: 'caption',
          content: 'BRANDDECK · 品牌演示引擎',
          x: 112,
          y: 220,
          w: 640,
          h: 32,
          style: { fontSize: 18, color: '$colors.primary', bold: true },
        },
        {
          type: 'text',
          id: 'cover-title',
          role: 'title',
          content: 'AI 驱动的企业内容生产',
          x: 112,
          y: 258,
          w: 860,
          h: 96,
          style: { fontSize: 60, bold: true, color: '$colors.textPrimary' },
        },
        {
          type: 'text',
          id: 'cover-subtitle',
          role: 'subtitle',
          content: '从企业模板到原生可编辑演示文稿，一条链路打通设计与内容',
          x: 112,
          y: 372,
          w: 780,
          h: 40,
          style: { fontSize: 22, color: '$colors.textSecondary' },
        },
        {
          type: 'shape',
          id: 'cover-rule',
          shapeType: 'line',
          x: 112,
          y: 440,
          w: 1040,
          h: 2,
          style: { stroke: '$colors.surface', strokeWidth: 2 },
        },
        {
          type: 'text',
          id: 'cover-meta',
          role: 'caption',
          content: '2026 年度品牌内容规划 · 内部评审稿',
          x: 112,
          y: 456,
          w: 620,
          h: 30,
          style: { fontSize: 16, color: '$colors.textSecondary' },
        },
      ],
    },
    {
      slideId: 'agenda',
      title: '目录',
      background: '$colors.background',
      elements: [
        {
          type: 'text',
          id: 'agenda-title',
          role: 'title',
          content: '目录',
          x: 80,
          y: 72,
          w: 420,
          h: 64,
          style: { fontSize: 40, bold: true, color: '$colors.textPrimary' },
        },
        {
          type: 'shape',
          id: 'agenda-bar',
          shapeType: 'rect',
          x: 80,
          y: 146,
          w: 96,
          h: 6,
          style: { fill: '$colors.primary' },
        },
        ...['01  业务挑战与机会', '02  增长数据与趋势', '03  落地路径与节奏', '04  投入产出与保障'].flatMap(
          (label, index) => {
            const top = 210 + index * 108
            return [
              {
                type: 'shape',
                id: `agenda-num-bg-${index}`,
                shapeType: 'roundRect',
                x: 80,
                y: top,
                w: 72,
                h: 72,
                style: { fill: '$colors.surface', radius: 16 },
              },
              {
                type: 'text',
                id: `agenda-num-${index}`,
                role: 'body',
                content: String(index + 1).padStart(2, '0'),
                x: 80,
                y: top + 18,
                w: 72,
                h: 40,
                style: { fontSize: 26, bold: true, color: '$colors.primary', align: 'center' },
              },
              {
                type: 'text',
                id: `agenda-label-${index}`,
                role: 'subtitle',
                content: label,
                x: 176,
                y: top + 18,
                w: 760,
                h: 40,
                style: { fontSize: 26, color: '$colors.textPrimary' },
              },
            ]
          },
        ),
      ],
    },
    {
      slideId: 'challenge',
      title: '业务挑战与机会',
      background: '$colors.background',
      elements: [
        {
          type: 'text',
          id: 'challenge-title',
          role: 'title',
          content: '业务挑战与机会',
          x: 80,
          y: 72,
          w: 720,
          h: 64,
          style: { fontSize: 40, bold: true, color: '$colors.textPrimary' },
        },
        {
          type: 'text',
          id: 'challenge-lead',
          role: 'body',
          content: '内容需求增长 3 倍，但排版与合规返工消耗了 60% 的交付时间。',
          x: 80,
          y: 150,
          w: 980,
          h: 36,
          style: { fontSize: 22, color: '$colors.textSecondary' },
        },
        ...[
          ['模板割裂', '各团队自建模板，品牌色与字体漂移'],
          ['返工频繁', '版式溢出与字号错位需要人工逐页修'],
          ['无法复用', '每次交付都从零开始，资产难以沉淀'],
        ].flatMap(([heading, body], index) => {
          const left = 80 + index * 380
          return [
            {
              type: 'shape',
              id: `challenge-card-${index}`,
              shapeType: 'roundRect',
              x: left,
              y: 220,
              w: 340,
              h: 280,
              style: { fill: '$colors.surface', radius: 20 },
            },
            {
              type: 'shape',
              id: `challenge-card-bar-${index}`,
              shapeType: 'rect',
              x: left + 32,
              y: 252,
              w: 56,
              h: 6,
              style: { fill: '$colors.primary' },
            },
            {
              type: 'text',
              id: `challenge-card-title-${index}`,
              role: 'subtitle',
              content: heading,
              x: left + 32,
              y: 276,
              w: 276,
              h: 44,
              style: { fontSize: 28, bold: true, color: '$colors.textPrimary' },
            },
            {
              type: 'text',
              id: `challenge-card-body-${index}`,
              role: 'body',
              content: body,
              x: left + 32,
              y: 330,
              w: 276,
              h: 96,
              style: { fontSize: 20, color: '$colors.textSecondary' },
            },
          ]
        }),
        {
          type: 'text',
          id: 'challenge-footnote',
          role: 'caption',
          content: '数据来源：2025 年内部内容运营复盘',
          x: 80,
          y: 560,
          w: 520,
          h: 28,
          style: { fontSize: 14, color: '$colors.textSecondary' },
        },
      ],
    },
    {
      slideId: 'growth',
      title: '增长数据与趋势',
      background: '$colors.background',
      elements: [
        {
          type: 'text',
          id: 'growth-title',
          role: 'title',
          content: '增长数据与趋势',
          x: 80,
          y: 72,
          w: 720,
          h: 64,
          style: { fontSize: 40, bold: true, color: '$colors.textPrimary' },
        },
        {
          type: 'chart',
          id: 'growth-chart',
          chartType: 'bar',
          x: 80,
          y: 176,
          w: 700,
          h: 400,
          data: {
            labels: ['Q1', 'Q2', 'Q3', 'Q4'],
            series: [
              { name: '内容交付量', values: [120, 168, 214, 286] },
              { name: '人工返工量', values: [86, 62, 40, 22] },
            ],
          },
          style: { legend: true, showValues: false },
        },
        ...[
          ['+138%', '内容交付量同比增长'],
          ['-74%', '返工工时下降'],
          ['3.2x', '单页产出效率提升'],
        ].flatMap(([value, label], index) => {
          const top = 190 + index * 128
          return [
            {
              type: 'text',
              id: `growth-kpi-value-${index}`,
              role: 'title',
              content: value,
              x: 840,
              y: top,
              w: 300,
              h: 64,
              style: { fontSize: 46, bold: true, color: '$colors.primary' },
            },
            {
              type: 'text',
              id: `growth-kpi-label-${index}`,
              role: 'body',
              content: label,
              x: 840,
              y: top + 62,
              w: 320,
              h: 32,
              style: { fontSize: 18, color: '$colors.textSecondary' },
            },
          ]
        }),
      ],
    },
    {
      slideId: 'roadmap',
      title: '落地路径与节奏',
      background: '$colors.background',
      elements: [
        {
          type: 'text',
          id: 'roadmap-title',
          role: 'title',
          content: '落地路径与节奏',
          x: 80,
          y: 72,
          w: 720,
          h: 64,
          style: { fontSize: 40, bold: true, color: '$colors.textPrimary' },
        },
        {
          type: 'svg',
          id: 'roadmap-svg',
          description: '四阶段流程：模板解析 → 卡组编排 → 预览微调 → 原生导出',
          x: 80,
          y: 190,
          w: 1100,
          h: 250,
          svgCode: `<svg xmlns="http://www.w3.org/2000/svg" width="1100" height="250" viewBox="0 0 1100 250">
            <defs>
              <linearGradient id="flowGrad" x1="0" y1="0" x2="1" y2="0">
                <stop offset="0%" stop-color="#C7000A"/>
                <stop offset="100%" stop-color="#E9002F"/>
              </linearGradient>
            </defs>
            <line x1="120" y1="126" x2="980" y2="126" stroke="#EBEBEB" stroke-width="4"/>
            <polygon points="984,126 966,117 966,135" fill="#C7000A"/>
            <circle cx="200" cy="126" r="46" fill="url(#flowGrad)"/>
            <text x="200" y="134" font-family="Microsoft YaHei" font-size="26" font-weight="700" fill="#FFFFFF" text-anchor="middle">01</text>
            <text x="200" y="204" font-family="Microsoft YaHei" font-size="22" fill="#1D1D1A" text-anchor="middle">模板解析</text>
            <circle cx="466" cy="126" r="46" fill="#EBEBEB"/>
            <text x="466" y="134" font-family="Microsoft YaHei" font-size="26" font-weight="700" fill="#C7000A" text-anchor="middle">02</text>
            <text x="466" y="204" font-family="Microsoft YaHei" font-size="22" fill="#1D1D1A" text-anchor="middle">卡组编排</text>
            <circle cx="732" cy="126" r="46" fill="#EBEBEB"/>
            <text x="732" y="134" font-family="Microsoft YaHei" font-size="26" font-weight="700" fill="#C7000A" text-anchor="middle">03</text>
            <text x="732" y="204" font-family="Microsoft YaHei" font-size="22" fill="#1D1D1A" text-anchor="middle">预览微调</text>
            <circle cx="998" cy="126" r="46" fill="#1D1D1A"/>
            <text x="998" y="134" font-family="Microsoft YaHei" font-size="26" font-weight="700" fill="#FFFFFF" text-anchor="middle">04</text>
            <text x="998" y="204" font-family="Microsoft YaHei" font-size="22" fill="#1D1D1A" text-anchor="middle">原生导出</text>
          </svg>`,
        },
        {
          type: 'text',
          id: 'roadmap-note',
          role: 'body',
          content: '四步全部自动化：模板只需解析一次，卡组可复用到下一次汇报。',
          x: 80,
          y: 468,
          w: 900,
          h: 36,
          style: { fontSize: 20, color: '$colors.textSecondary' },
        },
      ],
    },
    {
      slideId: 'end',
      title: '谢谢',
      background: '$colors.background',
      speakerNotes: '收尾页：留白为主，只放一句致谢与联系方式。',
      elements: [
        {
          type: 'text',
          id: 'end-title',
          role: 'title',
          content: '谢谢',
          x: 80,
          y: 288,
          w: 1120,
          h: 92,
          style: { fontSize: 64, bold: true, color: '$colors.textPrimary', align: 'center' },
        },
        {
          type: 'shape',
          id: 'end-bar',
          shapeType: 'rect',
          x: 608,
          y: 392,
          w: 64,
          h: 6,
          style: { fill: '$colors.primary' },
        },
        {
          type: 'text',
          id: 'end-contact',
          role: 'body',
          content: 'brand-content@example.com',
          x: 80,
          y: 424,
          w: 1120,
          h: 32,
          style: { fontSize: 20, color: '$colors.textSecondary', align: 'center' },
        },
      ],
    },
  ],
}

export default HW_DEMO_DECK
