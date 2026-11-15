import { composeDeck } from '../src/generate.js'
import { CANVAS, CardDeckSchema, LOGO_RESERVE, SAFE_AREA, innerBox, repairDeck } from '../src/schema.js'
const HW = { templateId: 'ppt-hw', templateName: 't', colors: { background:'#FFFFFF', surface:'#EBEBEB', textPrimary:'#1D1D1A', textSecondary:'#666666', primary:'#C7000A', accent:'#E9002F', chartPalette:['#C7000A'] }, fonts:{title:'Microsoft YaHei',body:'Arial',latin:'Arial'}, brandAssets:{logoPath:'x',logoWidthPx:10,logoHeightPx:10,icons:{}} }
const intersects = (a,b) => a.x < b.x+b.w && a.x+a.w > b.x && a.y < b.y+b.h && a.y+a.h > b.y
const contains = (o,i) => i.x>=o.x && i.y>=o.y && i.x+i.w<=o.x+o.w && i.y+i.h<=o.y+o.h
const LOGO = { x: CANVAS.width-SAFE_AREA.right-LOGO_RESERVE.width, y: SAFE_AREA.top, w: LOGO_RESERVE.width, h: LOGO_RESERVE.height }
function errs(slide) {
  const box = innerBox(HW); const out = []
  const t = slide.elements.filter(e=>e.type==='text'), s = slide.elements.filter(e=>e.type==='shape')
  for (const e of slide.elements) if (!contains(box,e)) out.push(`outside:${e.id}`)
  for (let i=0;i<t.length;i++){ if(intersects(t[i],LOGO)) out.push(`logo:${t[i].id}`); for(let j=i+1;j<t.length;j++) if(intersects(t[i],t[j])) out.push(`toverlap:${t[i].id}/${t[j].id}`)}
  for(let i=0;i<s.length;i++) for(let j=i+1;j<s.length;j++){ if(!intersects(s[i],s[j]))continue; if(contains(s[i],s[j])||contains(s[j],s[i]))continue; out.push(`soverlap:${s[i].id}/${s[j].id}`)}
  for (const a of t) for (const b of s) { if(!intersects(a,b))continue; if(contains(b,a))continue; out.push(`straddle:${a.id}/${b.id}`) }
  return out
}
const huge = '内容中台把模板解析、卡组编排、内容生成与原生导出串成一条完整链路，评审阶段直接给出可编辑的 pptx，业务方可以在自己的电脑上改字改图改配色而不破坏版式，因此一次解析可以长期复用，边际成本接近零，团队不再重复解释口径，也不会因为字号错位而返工重排，这是整套流程真正的价值所在。'
console.log('huge len', Array.from(huge).length)
const cases = {
  'compare/huge4': { title:'对比', slides:[{ title:'极端对比页标题也很长很长很长很长很长', kind:'compare', bullets:[`旧：${huge}`, `旧：${huge}`, `新：${huge}`, `新：${huge}`] }] },
  'stats/huge4': { title:'指标', slides:[{ title:'指标', kind:'stats', bullets:[`交付 ${huge}`, `返工 ${huge}`, '效率 3.2 倍', `成本 ${huge}`] }] },
  'timeline/12': { title:'时间线', slides:[{ title:'十二步', kind:'timeline', bullets:Array.from({length:12},(_,i)=>`第${i+1}步：${huge}`) }] },
  'stats/40': { title:'指标', slides:[{ title:'很多指标', kind:'stats', bullets:Array.from({length:40},(_,i)=>`第 ${i+1} 项增长 ${i+1}%`) }] },
  'compare/9': { title:'对比', slides:[{ title:'很多对比', kind:'compare', bullets:['旧：a','旧：b','旧：c','旧：d','旧：e','新：f','新：g','新：h','新：i'] }] },
  'timeline/1': { title:'时间线', slides:[{ title:'一步', kind:'timeline', bullets:['第一步：唯一一步'] }] },
  'compare/1': { title:'对比', slides:[{ title:'一条', kind:'compare', bullets:['唯一一条要点'] }] },
}
for (const [name, outline] of Object.entries(cases)) {
  for (const maxSlides of [12, 60]) {
    const { deck, warnings } = composeDeck({ outline, theme: HW, style: { maxSlides } })
    const bad = []
    if (!CardDeckSchema.safeParse(deck).success) bad.push('SCHEMA')
    for (const w of repairDeck(deck, HW).warnings) {
      if (/characters of text|overflows its box|clamped|logo corner|duplicate element/.test(w)) bad.push(`REPAIR:${w}`)
    }
    deck.slides.forEach((s) => errs(s).forEach((e) => bad.push(`${s.slideId}:${e}`)))
    console.log(`${bad.length === 0 ? 'OK  ' : 'FAIL'} ${name} maxSlides=${maxSlides} slides=${deck.slides.length} kinds=${deck.slides.map(s=>s.kind).join(',')}${bad.length?' :: '+bad.slice(0,3).join(' | '):''}`)
    if (bad.length === 0 && maxSlides === 60 && warnings.length) console.log(`     warnings(${warnings.length}): ${warnings.slice(0,3).join(' | ')}`)
  }
}
