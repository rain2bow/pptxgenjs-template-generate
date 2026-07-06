function speakerNotesText(slide, context = {}) {
  const explicit = explicitSpeakerNotes(slide);
  const normalized = normalizeSpeakerNotes(explicit);
  if (normalized) return normalized;
  const shouldGenerate = slide?.generateSpeakerNotes === true || context.spec?.generateSpeakerNotes === true;
  return shouldGenerate ? generateSpeakerNotes(slide, context) : '';
}

function explicitSpeakerNotes(slide) {
  if (!slide || typeof slide !== 'object') return undefined;
  if (slide.speakerNotes !== undefined) return slide.speakerNotes;
  if (slide.speaker_notes !== undefined) return slide.speaker_notes;
  if (slide.presenterNotes !== undefined) return slide.presenterNotes;
  if (slide.presenter_notes !== undefined) return slide.presenter_notes;
  return undefined;
}

function normalizeSpeakerNotes(value) {
  if (value === null || value === undefined) return '';
  if (typeof value === 'string' || typeof value === 'number') return String(value).trim();
  if (Array.isArray(value)) return value.map(normalizeNoteItem).filter(Boolean).join('\n');
  if (typeof value === 'object') {
    return Object.entries(value)
      .map(([key, item]) => normalizeNoteEntry(key, item))
      .filter(Boolean)
      .join('\n');
  }
  return '';
}

function normalizeNoteEntry(key, value) {
  const label = noteLabel(key);
  if (Array.isArray(value)) {
    const items = value.map(normalizeNoteItem).filter(Boolean);
    return items.length ? `${label}：\n${items.map((item) => `- ${item}`).join('\n')}` : '';
  }
  if (value && typeof value === 'object') {
    const text = normalizeSpeakerNotes(value);
    return text ? `${label}：${text}` : '';
  }
  const text = normalizeSpeakerNotes(value);
  return text ? `${label}：${text}` : '';
}

function normalizeNoteItem(item) {
  if (item === null || item === undefined) return '';
  if (typeof item === 'string' || typeof item === 'number') return String(item).trim();
  if (typeof item !== 'object') return '';
  const title = scalarText(item.title || item.label || item.name || item.metric || item.value);
  const body = scalarText(item.body || item.desc || item.note || item.summary || item.detail || item.story || item.text);
  if (title && body) return `${title}：${body}`;
  return title || body || normalizeSpeakerNotes(item);
}

function generateSpeakerNotes(slide, context = {}) {
  const parts = [];
  const page = Number(context.index) >= 0 ? `第 ${context.index + 1} 页` : '本页';
  const title = scalarText(slide?.title || slide?.quote || slide?.caseTitle || slide?.kicker);
  if (title) parts.push(`${page}讲解主题是“${title}”。`);
  const subtitle = scalarText(slide?.subtitle);
  if (subtitle) parts.push(`开场先说明背景：${subtitle}。`);
  const body = scalarText(slide?.body || slide?.story || slide?.summary || slide?.detail || slide?.note);
  if (body) parts.push(`核心信息：${body}`);
  const points = collectSlidePoints(slide).slice(0, 5);
  if (points.length) {
    parts.push(`讲解时按这几个重点展开：${points.map((item, i) => `${i + 1}. ${item}`).join('；')}。`);
  }
  const chart = slide?.chart || (Array.isArray(slide?.charts) ? slide.charts[0] : null);
  if (chart) {
    const chartTitle = scalarText(chart.title || chart.name || chart.type || chart.chartType);
    parts.push(`涉及图表时，重点解释${chartTitle ? `“${chartTitle}”` : '趋势、对比和异常值'}。`);
  }
  if (slide?.table || slide?.rows || slide?.headers) parts.push('涉及表格时，先讲口径，再讲关键行和结论。');
  if (slide?.image || slide?.images) parts.push('涉及图片时，先说明图片与本页结论的关系，再回到业务含义。');
  return parts.join('\n');
}

function collectSlidePoints(slide) {
  const values = [
    slide?.sections,
    slide?.items,
    slide?.columns,
    slide?.agenda,
    slide?.steps,
    slide?.milestones,
    slide?.nodes,
    slide?.layers,
    slide?.lanes,
    slide?.insights,
    slide?.points,
    slide?.metrics,
  ];
  const points = [];
  values.forEach((value) => normalizeCollection(value).forEach((item) => {
    const text = normalizeNoteItem(item);
    if (text) points.push(text);
  }));
  return points;
}

function normalizeCollection(value) {
  if (!value) return [];
  if (typeof value === 'string' || typeof value === 'number') return [value];
  if (Array.isArray(value)) return value;
  if (typeof value === 'object') return Object.entries(value).map(([title, body]) => ({ title, body }));
  return [];
}

function noteLabel(key) {
  const labels = {
    opening: '开场',
    intro: '开场',
    context: '背景',
    points: '讲解要点',
    keyPoints: '讲解要点',
    transition: '过渡',
    closing: '收束',
    reminder: '提醒',
  };
  return labels[key] || key;
}

function scalarText(value) {
  if (value === null || value === undefined) return '';
  if (typeof value === 'object') return '';
  return String(value).trim();
}

module.exports = {
  explicitSpeakerNotes,
  normalizeSpeakerNotes,
  speakerNotesText,
  generateSpeakerNotes,
};
