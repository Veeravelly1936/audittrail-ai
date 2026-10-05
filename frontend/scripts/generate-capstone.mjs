import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import pptxgen from 'pptxgenjs'

const scriptDir = path.dirname(fileURLToPath(import.meta.url))
const root = path.resolve(scriptDir, '../..')
const outputDir = path.join(root, 'deliverables')
const screenshotSource = path.join(root, 'audittrail-dashboard.png')
const screenshotTarget = path.join(outputDir, 'audittrail-dashboard.png')

fs.mkdirSync(outputDir, { recursive: true })
if (fs.existsSync(screenshotSource)) fs.copyFileSync(screenshotSource, screenshotTarget)

const pptx = new pptxgen()
pptx.layout = 'LAYOUT_WIDE'
pptx.author = '[Your name]'
pptx.subject = 'Capstone journey through Workshops 1–6'
pptx.title = 'Two Applications, Six Workshops'
pptx.company = 'Capstone presentation'
pptx.lang = 'en-US'
pptx.theme = {
  headFontFace: 'Aptos Display',
  bodyFontFace: 'Aptos',
  lang: 'en-US',
}

const W = 13.333
const H = 7.5
const C = {
  ink: '18362C',
  forest: '174536',
  forest2: '235C47',
  lime: 'C9EF7B',
  paper: 'F5F6F1',
  white: 'FFFFFF',
  text: '26352F',
  muted: '78847D',
  line: 'DFE5DD',
  pale: 'E8F0E5',
  orange: 'D8794C',
  paleOrange: 'FFF0E8',
  teal: '5F9590',
  paleTeal: 'EAF3F1',
  blue: '6682A0',
  paleBlue: 'EDF2F8',
  gold: 'A67828',
  paleGold: 'FFF6DF',
  red: 'B9453A',
  paleRed: 'FCEDEA',
}
const SH = pptx.ShapeType

function text(slide, value, x, y, w, h, options = {}) {
  slide.addText(value, {
    x, y, w, h,
    fontFace: options.fontFace || 'Aptos',
    fontSize: options.fontSize || 14,
    color: options.color || C.text,
    bold: options.bold || false,
    margin: options.margin ?? 0,
    breakLine: false,
    valign: options.valign || 'mid',
    align: options.align || 'left',
    paraSpaceAfterPt: options.paraSpaceAfterPt || 0,
    fit: 'shrink',
    transparency: options.transparency || 0,
    isTextBox: true,
  })
}

function box(slide, x, y, w, h, fill = C.white, line = C.line, radius = true) {
  slide.addShape(radius ? SH.roundRect : SH.rect, {
    x, y, w, h,
    rectRadius: 0.08,
    fill: { color: fill },
    line: { color: line, width: 0.8 },
    radius: 0.08,
  })
}

function pill(slide, label, x, y, w, fill = C.pale, color = C.forest2) {
  slide.addShape(SH.roundRect, {
    x, y, w, h: 0.27,
    rectRadius: 0.08,
    line: { color: fill, transparency: 100 },
    fill: { color: fill },
  })
  text(slide, label.toUpperCase(), x + 0.08, y + 0.025, w - 0.16, 0.19, {
    fontSize: 7.5, bold: true, color,
  })
}

function base(section, title, subtitle = '') {
  const slide = pptx.addSlide()
  slide.background = { color: C.paper }
  slide.addShape(SH.rect, { x: 0, y: 0, w: W, h: 0.12, line: { color: C.lime, transparency: 100 }, fill: { color: C.lime } })
  text(slide, section.toUpperCase(), 0.62, 0.27, 8, 0.22, { fontSize: 8, bold: true, color: C.forest2 })
  text(slide, title, 0.62, 0.58, 12.05, 0.52, { fontFace: 'Aptos Display', fontSize: 25, bold: true, color: C.ink })
  if (subtitle) text(slide, subtitle, 0.64, 1.14, 11.9, 0.34, { fontSize: 11, color: C.muted })
  slide.addShape(SH.line, { x: 0.62, y: 7.08, w: 12.05, h: 0, line: { color: C.line, width: 0.8 } })
  text(slide, 'AUDITTRAIL AI  ·  CAPSTONE JOURNEY', 0.64, 7.16, 5.2, 0.16, { fontSize: 7, bold: true, color: C.muted })
  text(slide, String(pptx._slides.length).padStart(2, '0'), 12.05, 7.13, 0.55, 0.22, { fontSize: 8, bold: true, color: C.forest2, align: 'right' })
  return slide
}

function notes(slide, content) {
  slide.addNotes(content)
}

function bullets(slide, items, x, y, w, options = {}) {
  const gap = options.gap ?? 0.52
  const fontSize = options.fontSize ?? 13
  items.forEach((item, index) => {
    const yy = y + index * gap
    slide.addShape(SH.ellipse, {
      x, y: yy + 0.105, w: 0.075, h: 0.075,
      line: { color: options.dotColor || C.orange, transparency: 100 },
      fill: { color: options.dotColor || C.orange },
    })
    text(slide, item, x + 0.2, yy, w - 0.2, gap - 0.04, { fontSize, color: options.color || C.text, valign: 'top' })
  })
}

function metric(slide, value, label, x, y, w, color = C.forest2) {
  box(slide, x, y, w, 0.88, C.white, C.line)
  text(slide, value, x + 0.16, y + 0.12, w - 0.32, 0.38, { fontFace: 'Aptos Display', fontSize: 25, bold: true, color })
  text(slide, label, x + 0.16, y + 0.55, w - 0.32, 0.18, { fontSize: 8.5, color: C.muted })
}

function node(slide, x, y, w, h, title, detail, fill = C.white, accent = C.forest2) {
  box(slide, x, y, w, h, fill, C.line)
  slide.addShape(SH.rect, { x, y, w: 0.07, h, line: { color: accent, transparency: 100 }, fill: { color: accent } })
  text(slide, title, x + 0.18, y + 0.12, w - 0.32, 0.32, { fontSize: 12.5, bold: true, color: C.ink })
  text(slide, detail, x + 0.18, y + 0.49, w - 0.33, h - 0.58, { fontSize: 9, color: C.muted, valign: 'top' })
}

function arrow(slide, x1, y1, x2, y2, color = C.teal) {
  slide.addShape(SH.line, { x: x1, y: y1, w: x2 - x1, h: y2 - y1, line: { color, width: 1.6, endArrowType: 'triangle' } })
}

function sourceNeeded(slide, x = 9.45, y = 1.72, w = 2.85) {
  pill(slide, 'App #1 source needed', x, y, w, C.paleOrange, C.orange)
}

// 1. Title
{
  const slide = pptx.addSlide()
  slide.background = { color: C.ink }
  slide.addShape(SH.rect, { x: 0, y: 0, w: 0.22, h: H, line: { color: C.lime, transparency: 100 }, fill: { color: C.lime } })
  slide.addShape(SH.ellipse, { x: 9.15, y: -1.25, w: 5.25, h: 5.25, line: { color: C.forest2, transparency: 100 }, fill: { color: C.forest2, transparency: 35 } })
  pill(slide, 'CAPSTONE · WORKSHOPS 1–6', 0.78, 0.72, 2.55, '315B49', C.lime)
  text(slide, 'Two applications,', 0.78, 1.5, 8.5, 0.74, { fontFace: 'Aptos Display', fontSize: 36, bold: true, color: C.white })
  text(slide, 'one evolving practice.', 0.78, 2.3, 9.3, 0.82, { fontFace: 'Aptos Display', fontSize: 36, bold: true, color: C.lime })
  text(slide, 'A journey from problem framing to a live AWS deployment', 0.82, 3.3, 8.3, 0.35, { fontSize: 15, color: 'D6E2DB' })
  box(slide, 0.82, 4.18, 4.15, 0.92, '214638', '456A56')
  text(slide, 'APPLICATION 01', 1.02, 4.32, 3.65, 0.2, { fontSize: 8, bold: true, color: 'BCD0C1' })
  text(slide, '[App #1 name · Workshop 1–4]', 1.02, 4.62, 3.65, 0.26, { fontSize: 13, bold: true, color: C.white })
  box(slide, 5.18, 4.18, 4.15, 0.92, '214638', '456A56')
  text(slide, 'APPLICATION 02', 5.38, 4.32, 3.65, 0.2, { fontSize: 8, bold: true, color: 'BCD0C1' })
  text(slide, 'AuditTrail AI · Workshops 5–6', 5.38, 4.62, 3.65, 0.26, { fontSize: 13, bold: true, color: C.white })
  text(slide, '[Your name]', 0.84, 5.72, 3.4, 0.32, { fontSize: 13, bold: true, color: C.white })
  text(slide, 'APP 01  [ADD VERIFIED LIVE URL]', 0.84, 6.23, 4.0, 0.22, { fontSize: 8.5, color: 'B5C7BB' })
  text(slide, 'APP 02  http://54.82.56.190:8000', 5.18, 6.23, 4.4, 0.22, { fontSize: 8.5, color: C.lime })
  text(slide, '01 / 23', 11.62, 6.77, 0.9, 0.18, { fontSize: 8, color: 'A9BBB0', align: 'right' })
  notes(slide, 'Introduce yourself and both projects. Replace [Your name], App #1 name, and the App #1 URL before presenting. AuditTrail AI is the second application; its public demo is served from an AWS Fargate task over HTTP.')
}

// 2. Executive summary
{
  const slide = base('01 · Overview', 'Two problems, two applied solutions', 'One capstone arc: frame the problem, prototype, test, deploy, and learn from feedback.')
  box(slide, 0.68, 1.72, 5.82, 3.28, C.white)
  pill(slide, 'APPLICATION 01', 0.94, 1.95, 1.65, C.paleOrange, C.orange)
  text(slide, '[Problem statement]', 0.96, 2.42, 4.95, 0.4, { fontSize: 20, bold: true, color: C.ink })
  bullets(slide, ['Who experiences the problem? [Add audience]', 'What makes it costly or frustrating? [Add evidence]', 'AI-assisted solution: [Name the workflow]'], 0.98, 3.1, 5.1, { gap: 0.5, fontSize: 11.5 })
  sourceNeeded(slide, 3.65, 4.55, 2.3)
  box(slide, 6.82, 1.72, 5.82, 3.28, C.white)
  pill(slide, 'APPLICATION 02', 7.08, 1.95, 1.65, C.pale, C.forest2)
  text(slide, 'Invoice fraud & anomaly review', 7.08, 2.42, 5.05, 0.4, { fontSize: 18, bold: true, color: C.ink })
  bullets(slide, ['Problem: manual rate-cap and duplicate checks', 'Solution: deterministic invoice rules + auditor workflow', 'Audience: audit and accounts-payable teams'], 7.1, 3.1, 5.0, { gap: 0.5, fontSize: 11.5, dotColor: C.forest2 })
  metric(slide, '4', 'seeded invoices', 0.72, 5.32, 2.8)
  metric(slide, '3 / 4', 'invoices with anomalies', 3.72, 5.32, 2.8, C.orange)
  metric(slide, '2 · 1 · 1', 'High · Medium · Low risk', 6.72, 5.32, 2.8, C.teal)
  metric(slide, '$150/hr', 'contract cap checked', 9.72, 5.32, 2.8)
  notes(slide, 'The App #2 counts come from its four seeded invoices: two High rate violations, one Medium duplicate, and one Low clean invoice. No measured business savings or real invoice volume are available; do not present the metrics as production outcomes. App #1 fields require your source material.')
}

// 3. App 1 problem
{
  const slide = base('02 · Application 01 · Workshop 1', 'Start with the user’s problem', 'Problem framing before solution design')
  sourceNeeded(slide)
  box(slide, 0.72, 1.8, 7.15, 3.95, C.white)
  text(slide, 'PROBLEM STATEMENT', 1.04, 2.12, 3.6, 0.22, { fontSize: 8, bold: true, color: C.orange })
  text(slide, '[When user group] needs to [complete task], they currently [pain point], which leads to [measurable impact].', 1.04, 2.62, 6.25, 1.4, { fontFace: 'Aptos Display', fontSize: 24, bold: true, color: C.ink, valign: 'top' })
  text(slide, 'Evidence to include', 1.04, 4.48, 2.3, 0.24, { fontSize: 10, bold: true, color: C.forest2 })
  bullets(slide, ['Interview or observation', 'Frequency / time / cost baseline', 'Current workaround'], 1.06, 4.82, 6.2, { gap: 0.32, fontSize: 10 })
  box(slide, 8.2, 1.8, 4.35, 3.95, C.paleOrange, 'F0D6C8')
  text(slide, 'WORKSHOP 1', 8.52, 2.15, 2.2, 0.22, { fontSize: 8, bold: true, color: C.orange })
  text(slide, 'The “why” behind the build', 8.52, 2.55, 3.55, 0.7, { fontSize: 20, bold: true, color: C.ink })
  text(slide, 'AI use case: [Explain what AI does, what stays human-led, and why AI is appropriate.]', 8.52, 3.55, 3.55, 1.12, { fontSize: 13, color: C.text, valign: 'top' })
  text(slide, 'Do not substitute a feature list for user evidence.', 8.52, 5.05, 3.45, 0.38, { fontSize: 9, bold: true, color: C.orange })
  notes(slide, 'Fill in the bracketed audience, task, pain, and impact from App #1 discovery. Add one real evidence point such as a quote, observation, or baseline metric. The assignment specifies an AI use case but provides no App #1 domain details in the available project materials.')
}

// 4. App 1 AI hypothesis
{
  const slide = base('02 · Application 01 · Workshop 1', 'Turn the problem into an AI hypothesis', 'Define the model’s job, human decision point, and success criteria.')
  sourceNeeded(slide)
  const xs = [0.78, 4.62, 8.46]
  const titles = ['INPUT', 'AI ASSIST', 'HUMAN OUTCOME']
  const descriptions = [
    '[User data / task / context]',
    '[Classification, generation, search, or recommendation]',
    '[Decision or action the user can take]',
  ]
  xs.forEach((x, i) => node(slide, x, 2.2, 3.25, 1.45, titles[i], descriptions[i], i === 1 ? C.paleTeal : C.white, i === 1 ? C.teal : C.forest2))
  arrow(slide, 4.12, 2.9, 4.48, 2.9)
  arrow(slide, 7.96, 2.9, 8.32, 2.9)
  box(slide, 0.8, 4.28, 11.5, 1.32, C.ink, C.ink)
  text(slide, 'SUCCESS MEASURE', 1.12, 4.54, 2.1, 0.2, { fontSize: 8, bold: true, color: C.lime })
  text(slide, '[e.g., time saved, task accuracy, adoption, or error reduction]', 3.15, 4.46, 8.55, 0.42, { fontSize: 16, bold: true, color: C.white })
  text(slide, 'Responsible use: [privacy boundary] · [known failure mode] · [human review]', 1.12, 5.08, 10.8, 0.22, { fontSize: 10, color: 'CEDDD2' })
  notes(slide, 'Describe the real data source, AI capability, human review step, and success measure. Make clear where automation stops. Keep this diagram as an editable template until App #1 evidence is supplied.')
}

// 5. App 1 architecture
{
  const slide = base('03 · Application 01 · Technical solution', 'Architecture decisions for App #1', 'Replace each technology placeholder with the stack actually used.')
  sourceNeeded(slide)
  node(slide, 0.7, 2.45, 2.75, 1.55, 'USER EXPERIENCE', '[Claude Artifact / frontend framework]', C.white, C.forest2)
  node(slide, 4.0, 2.45, 2.75, 1.55, 'APPLICATION LAYER', '[API / orchestration / validation]', C.white, C.teal)
  node(slide, 7.3, 2.45, 2.75, 1.55, 'AI CAPABILITY', '[Model / API / retrieval choice]', C.paleTeal, C.teal)
  node(slide, 10.35, 2.45, 2.25, 1.55, 'DATA', '[Storage / integrations]', C.white, C.orange)
  arrow(slide, 3.48, 3.22, 3.88, 3.22)
  arrow(slide, 6.78, 3.22, 7.18, 3.22)
  arrow(slide, 10.0, 3.22, 10.23, 3.22)
  box(slide, 1.0, 4.65, 11.25, 1.08, C.paleGold, 'EFE0B8')
  text(slide, 'DECISION LOG', 1.28, 4.88, 1.7, 0.2, { fontSize: 8, bold: true, color: C.gold })
  text(slide, '[Why this stack? What constraint shaped deployment, cost, latency, privacy, or maintainability?]', 3.02, 4.8, 8.65, 0.45, { fontSize: 13, color: C.text })
  notes(slide, 'This is the requested architecture diagram. Replace each bracket with App #1 implementation facts. Use the notes to explain at least one tradeoff rather than naming technologies without rationale.')
}

// 6. App 1 Claude Artifact
{
  const slide = base('04 · Application 01 · Phase 3 · Workshop 1', 'Prototype the experience with Claude Artifacts', 'Show the transition from prompt to a testable interaction.')
  sourceNeeded(slide)
  const cards = [
    ['01 · PROMPT', '[Paste the prompt that set the user, task, constraints, and success criterion.]'],
    ['02 · ARTIFACT', '[Add a screenshot or describe the key screen and interaction.]'],
    ['03 · ITERATION', '[What changed after testing or critique?]'],
  ]
  cards.forEach((item, index) => {
    const x = 0.82 + index * 4.05
    box(slide, x, 2.0, 3.52, 2.65, index === 1 ? C.ink : C.white, index === 1 ? C.ink : C.line)
    text(slide, item[0], x + 0.23, 2.28, 2.95, 0.22, { fontSize: 8, bold: true, color: index === 1 ? C.lime : C.orange })
    text(slide, item[1], x + 0.23, 2.78, 3.0, 1.25, { fontSize: 14, bold: index === 1, color: index === 1 ? C.white : C.text, valign: 'top' })
  })
  text(slide, 'Artifact URL or screenshot: [ADD SOURCE]', 0.9, 5.15, 5.8, 0.28, { fontSize: 11, bold: true, color: C.forest2 })
  text(slide, 'Evidence needed before submission', 0.9, 5.68, 3.5, 0.22, { fontSize: 9, bold: true, color: C.muted })
  notes(slide, 'The rubric explicitly requests a Claude Artifact prototype. Add the actual prompt, prototype image or URL, and one concrete iteration. No App #1 artifact is present in this workspace.')
}

// 7. App 1 IDE migration
{
  const slide = base('05 · Application 01 · Phase 3 · Workshop 2', 'Move from prototype to local development', 'Document the IDE migration and the first working local run.')
  sourceNeeded(slide)
  box(slide, 0.8, 1.85, 5.55, 3.86, C.white)
  text(slide, 'DEVELOPMENT TRANSITION', 1.1, 2.15, 3.4, 0.22, { fontSize: 8, bold: true, color: C.orange })
  bullets(slide, ['Project scaffold: [framework / command]', 'Prototype pieces reused: [list]', 'New modules and API boundaries: [list]', 'Local run + first verification: [command / result]'], 1.12, 2.72, 4.8, { gap: 0.61, fontSize: 12 })
  box(slide, 6.75, 1.85, 5.78, 3.86, C.pale)
  text(slide, 'LOCAL SYSTEM MAP', 7.05, 2.15, 3.4, 0.22, { fontSize: 8, bold: true, color: C.forest2 })
  node(slide, 7.12, 2.78, 2.12, 1.15, 'IDE', '[code + terminal]', C.white, C.forest2)
  node(slide, 10.02, 2.78, 2.12, 1.15, 'LOCAL APP', '[browser + API]', C.white, C.teal)
  arrow(slide, 9.28, 3.35, 9.9, 3.35)
  text(slide, 'Screenshot: [replace with local App #1 view]', 7.12, 4.45, 4.7, 0.3, { fontSize: 10, italic: true, color: C.muted })
  notes(slide, 'Describe the actual Workshop 2 migration: IDE, framework, files, setup commands, and first local result. The diagram is a placeholder, not a claim about App #1 implementation.')
}

// 8. App 1 evaluation
{
  const slide = base('06 · Application 01 · Phase 4', 'Evaluate quality, safety, and usefulness', 'Separate what was tested from what still needs evaluation.')
  sourceNeeded(slide)
  const cols = [
    ['FUNCTIONAL', '[Core user journey tests]', C.pale],
    ['RESPONSIBLE AI', '[Bias, privacy, hallucination, human oversight]', C.paleTeal],
    ['SECURITY', '[Input validation, secrets, access control]', C.paleOrange],
  ]
  cols.forEach((item, index) => {
    const x = 0.78 + index * 4.18
    box(slide, x, 2.05, 3.74, 2.7, item[2], item[2])
    text(slide, item[0], x + 0.22, 2.34, 3.2, 0.22, { fontSize: 8, bold: true, color: index === 1 ? C.teal : C.forest2 })
    text(slide, item[1], x + 0.22, 2.88, 3.12, 1.1, { fontSize: 14, bold: true, color: C.ink, valign: 'top' })
    text(slide, 'Evidence: [tests / observations]', x + 0.22, 4.23, 3.12, 0.22, { fontSize: 9, color: C.muted })
  })
  box(slide, 0.82, 5.18, 11.55, 0.74, C.ink, C.ink)
  text(slide, 'Claim only outcomes that were actually tested. Mark the remaining checks as planned.', 1.12, 5.4, 10.9, 0.24, { fontSize: 14, bold: true, color: C.white })
  notes(slide, 'Use this slide to report the test plan and measured findings for App #1. Responsible AI is not a checkbox: state which risks were evaluated and the evidence. The source record for those tests was not supplied.')
}

// 9. App 1 bugs
{
  const slide = base('06 · Application 01 · Phase 4', 'What failed, what changed, what passed', 'A useful test report connects each fix to a reproducible result.')
  sourceNeeded(slide)
  const headers = ['BUG / RISK', 'REPRODUCTION', 'FIX', 'RESULT']
  const xs = [0.86, 3.95, 6.72, 9.5]
  headers.forEach((h, i) => text(slide, h, xs[i], 2.0, i === 3 ? 2.5 : 2.6, 0.24, { fontSize: 8, bold: true, color: C.muted }))
  for (let r = 0; r < 3; r++) {
    const y = 2.42 + r * 0.92
    box(slide, 0.78, y, 11.75, 0.68, r === 1 ? C.white : 'EFF2EB', r === 1 ? C.line : 'EFF2EB', false)
    const labels = ['[Issue]', '[Steps]', '[Change]', '[Pass / fail]']
    labels.forEach((label, i) => text(slide, label, xs[i], y + 0.19, i === 3 ? 2.5 : 2.6, 0.24, { fontSize: 10.5, color: C.text }))
  }
  text(slide, 'Testing summary: [n passed] / [n run] · Environment: [browser / device / version]', 0.88, 5.5, 10.6, 0.35, { fontSize: 12, bold: true, color: C.forest2 })
  notes(slide, 'Replace the three example rows with actual bugs and reproduction steps. Add real pass/fail counts and environment. If there were no bugs, state the scope tested rather than inventing defects.')
}

// 10. App 1 App Runner architecture
{
  const slide = base('07 · Application 01 · Phase 5 · Workshop 3', 'Deploy App #1 with AWS App Runner', 'Show the actual source, build path, runtime, and public endpoint.')
  sourceNeeded(slide)
  node(slide, 0.9, 2.35, 2.45, 1.3, 'SOURCE', '[GitHub / image source]', C.white, C.forest2)
  node(slide, 4.05, 2.35, 2.45, 1.3, 'APP RUNNER', '[build + managed runtime]', C.paleTeal, C.teal)
  node(slide, 7.2, 2.35, 2.45, 1.3, 'APPLICATION', '[container / route / health]', C.white, C.orange)
  node(slide, 10.35, 2.35, 2.1, 1.3, 'USERS', '[live link]', C.white, C.forest2)
  arrow(slide, 3.4, 3.0, 3.95, 3.0)
  arrow(slide, 6.55, 3.0, 7.1, 3.0)
  arrow(slide, 9.7, 3.0, 10.25, 3.0)
  box(slide, 0.95, 4.38, 11.35, 1.05, C.white)
  text(slide, 'DEPLOYMENT RECORD', 1.24, 4.62, 1.9, 0.2, { fontSize: 8, bold: true, color: C.orange })
  text(slide, 'Region [ ] · Service [ ] · Build source [ ] · Health check [ ] · Runtime size [ ]', 3.2, 4.55, 8.55, 0.38, { fontSize: 12, color: C.text })
  notes(slide, 'The assignment specifies AWS App Runner for App #1, but account-level details and URL were not provided. Add exact service region, source mode, health check, and live link. The diagram is an architecture guide only.')
}

// 11. App 1 deployment learning
{
  const slide = base('07 · Application 01 · Phase 5 · Workshop 3', 'Deployment is a system, not a button', 'Capture one real challenge and the evidence that the service worked.')
  sourceNeeded(slide)
  box(slide, 0.82, 1.92, 5.45, 3.65, C.ink, C.ink)
  text(slide, 'THE CHALLENGE', 1.15, 2.27, 2.4, 0.22, { fontSize: 8, bold: true, color: C.lime })
  text(slide, '[Build, port, permission, dependency, or health-check issue]', 1.15, 2.72, 4.6, 1.0, { fontSize: 20, bold: true, color: C.white, valign: 'top' })
  text(slide, 'Fix: [what changed]  ·  Lesson: [what you will repeat]', 1.15, 4.35, 4.55, 0.58, { fontSize: 12, color: 'D7E5DC', valign: 'top' })
  box(slide, 6.68, 1.92, 5.65, 3.65, C.white)
  text(slide, 'LIVE DEMO', 7.0, 2.27, 2.0, 0.22, { fontSize: 8, bold: true, color: C.forest2 })
  text(slide, '[APP RUNNER URL]', 7.0, 2.78, 4.8, 0.42, { fontSize: 18, bold: true, color: C.teal })
  bullets(slide, ['Homepage returns 200: [verified]', 'API / workflow responds: [verified]', 'Logs and health status: [verified]'], 7.02, 3.55, 4.8, { gap: 0.52, fontSize: 11 })
  notes(slide, 'Do not paste a URL unless verified. Include one deployment challenge you actually encountered and how you resolved it. The App #1 URL and deployment facts need to be supplied.')
}

// 12. App 1 iteration
{
  const slide = base('08 · Application 01 · Phase 6 · Workshop 4', 'Close the loop with user feedback', 'Feedback should change a product decision, not just decorate a slide.')
  sourceNeeded(slide)
  const phases = [
    ['HEARD', '[User feedback / observation]'],
    ['DECIDED', '[Prioritized change + rationale]'],
    ['CHANGED', '[UI, prompt, logic, or workflow]'],
    ['RECHECKED', '[New result / follow-up feedback]'],
  ]
  phases.forEach((item, index) => {
    const x = 0.7 + index * 3.17
    box(slide, x, 2.15, 2.72, 1.7, index === 2 ? C.pale : C.white)
    text(slide, `0${index + 1} · ${item[0]}`, x + 0.2, 2.42, 2.25, 0.2, { fontSize: 8, bold: true, color: C.forest2 })
    text(slide, item[1], x + 0.2, 2.92, 2.23, 0.58, { fontSize: 13, bold: true, color: C.ink, valign: 'top' })
    if (index < phases.length - 1) arrow(slide, x + 2.78, 3.0, x + 3.08, 3.0)
  })
  box(slide, 0.82, 4.47, 11.45, 0.92, C.paleOrange, 'F0D6C8')
  text(slide, 'Feedback source: [who / how many]  ·  Implemented iteration: [change]  ·  Outcome: [evidence]', 1.12, 4.78, 10.8, 0.28, { fontSize: 12.5, bold: true, color: C.ink })
  notes(slide, 'The brief identifies Workshop 4 as user feedback and improvements, but does not state what feedback was collected. Name the participants or source, the change made, and the post-change evidence.')
}

// 13. App 2 problem
{
  const slide = base('09 · Application 02 · Workshop 5', 'AuditTrail AI: make invoice risk visible', 'An auditor dashboard for deterministic invoice-fraud and anomaly checks.')
  box(slide, 0.78, 1.85, 7.0, 3.83, C.white)
  text(slide, 'THE PROBLEM', 1.1, 2.18, 2.1, 0.22, { fontSize: 8, bold: true, color: C.orange })
  text(slide, 'Contract-rate violations and duplicate invoices can hide in routine review.', 1.1, 2.65, 5.95, 0.95, { fontSize: 23, bold: true, color: C.ink, valign: 'top' })
  bullets(slide, ['Manual comparisons are repetitive.', 'Risk signals compete with routine line-item details.', 'Auditors need explanations and final decision authority.'], 1.12, 4.0, 5.9, { gap: 0.43, fontSize: 11.5, dotColor: C.orange })
  box(slide, 8.12, 1.85, 4.4, 3.83, C.ink, C.ink)
  text(slide, 'RULES FIRST', 8.48, 2.22, 2.0, 0.22, { fontSize: 8, bold: true, color: C.lime })
  text(slide, 'AI-branded, rule-based decisions', 8.48, 2.66, 3.45, 0.75, { fontSize: 19, bold: true, color: C.white })
  text(slide, 'Rate > $150/hr → High risk\nDuplicate invoice number → Medium risk\nHuman auditor chooses the disposition', 8.48, 3.72, 3.42, 1.12, { fontSize: 12, color: 'D7E4DB', valign: 'top' })
  notes(slide, 'The current implementation uses deterministic rules rather than a generative AI model. This is intentional and testable for two explicit anomaly types. Be transparent about that distinction when presenting the product as AuditTrail AI.')
}

// 14. App 2 architecture
{
  const slide = base('09 · Application 02 · Technical architecture', 'One container serves the dashboard and API', 'A simple single-origin design keeps the demo deployment small.')
  node(slide, 0.72, 2.05, 2.48, 1.32, 'AUDITOR', 'Browser · React dashboard', C.white, C.forest2)
  node(slide, 3.68, 2.05, 2.48, 1.32, 'FASTAPI', 'Pydantic validation · rules', C.paleTeal, C.teal)
  node(slide, 6.64, 2.05, 2.48, 1.32, 'IN-MEMORY DATA', 'Four seeded invoice records', C.white, C.orange)
  node(slide, 9.6, 2.05, 2.85, 1.32, 'ECS FARGATE', 'Single task · static UI + API', C.white, C.forest2)
  arrow(slide, 3.24, 2.7, 3.58, 2.7)
  arrow(slide, 6.2, 2.7, 6.54, 2.7)
  arrow(slide, 9.16, 2.7, 9.5, 2.7)
  box(slide, 0.92, 4.15, 11.55, 1.02, C.white)
  text(slide, 'BUILD + OPERATE', 1.22, 4.42, 1.75, 0.2, { fontSize: 8, bold: true, color: C.forest2 })
  text(slide, 'GitHub → CodeBuild → ECR → ECS · CloudWatch logs · /healthz health check', 3.02, 4.34, 8.9, 0.35, { fontSize: 13, bold: true, color: C.ink })
  text(slide, 'No persistent database is connected.', 1.22, 4.84, 4.7, 0.18, { fontSize: 9, color: C.muted })
  notes(slide, 'Explain the complete deployed topology. The React/Vite app is compiled into static assets and served by FastAPI in the container. The API holds records and review status in memory; there is no database.')
}

// 15. App 2 screenshot
{
  const slide = base('10 · Application 02 · Product experience', 'A compact workspace for invoice triage', 'The selected invoice, contract comparison, and next decision stay together.')
  if (fs.existsSync(screenshotTarget)) {
    slide.addImage({ path: screenshotTarget, x: 0.68, y: 1.82, w: 8.05, h: 3.89 })
    slide.addShape(SH.rect, { x: 0.68, y: 1.82, w: 8.05, h: 3.89, line: { color: C.line, width: 1 }, fill: { color: C.white, transparency: 100 } })
  } else {
    box(slide, 0.68, 1.82, 8.05, 3.89, C.white)
    text(slide, '[Dashboard screenshot asset missing]', 1.0, 3.3, 7.4, 0.35, { fontSize: 16, bold: true, color: C.muted, align: 'center' })
  }
  text(slide, 'LIVE AWS DASHBOARD · CAPTURED OCT 5, 2026', 0.72, 5.87, 7.7, 0.2, { fontSize: 7.5, bold: true, color: C.muted })
  const points = [
    ['01', 'Scan', 'Search/filter invoice register'],
    ['02', 'Compare', 'Line items vs. $150/hr cap'],
    ['03', 'Decide', 'Approve · flag · reject'],
  ]
  points.forEach((item, index) => {
    const y = 2.02 + index * 1.05
    box(slide, 9.02, y, 3.48, 0.82, index === 1 ? C.pale : C.white)
    text(slide, item[0], 9.25, y + 0.17, 0.4, 0.25, { fontSize: 12, bold: true, color: C.orange })
    text(slide, item[1], 9.8, y + 0.12, 2.2, 0.23, { fontSize: 13, bold: true, color: C.ink })
    text(slide, item[2], 9.8, y + 0.43, 2.35, 0.19, { fontSize: 8.5, color: C.muted })
  })
  notes(slide, 'This is a screenshot captured from the live public app on October 5, 2026. Walk through the register, selected invoice, contract-rate comparison, explanation, and auditor actions. The app is a demonstration and contains seeded data.')
}

// 16. App 2 tests
{
  const slide = base('11 · Application 02 · Phase 4 · Testing & security', 'Test the rules at their boundaries', 'Observed test results from the seeded dataset and API validation.')
  metric(slide, '4', 'GET /api/invoices records', 0.78, 1.9, 2.75)
  metric(slide, '2 · 1 · 1', 'High · Medium · Low', 3.78, 1.9, 2.75, C.orange)
  metric(slide, '6 / 6', 'malformed payloads rejected', 6.78, 1.9, 2.75, C.teal)
  metric(slide, '200 / 404 / 422', 'health / missing ID / invalid input', 9.78, 1.9, 2.75, C.blue)
  box(slide, 0.82, 3.2, 5.52, 2.3, C.white)
  text(slide, 'SECURITY CHECKS', 1.1, 3.48, 2.2, 0.2, { fontSize: 8, bold: true, color: C.forest2 })
  bullets(slide, ['CORS: only http://localhost:5173', 'Strict status Literal + extra="forbid"', 'No secrets / .env / key files found'], 1.12, 3.9, 4.95, { gap: 0.42, fontSize: 10.5 })
  box(slide, 6.72, 3.2, 5.6, 2.3, C.paleOrange, 'F0D6C8')
  text(slide, 'RESPONSIBLE-AI LIMIT', 7.02, 3.48, 2.5, 0.2, { fontSize: 8, bold: true, color: C.orange })
  text(slide, 'Rules explain the flags; there is no trained model or confidence score. Auditors retain the decision.', 7.02, 3.94, 4.8, 0.9, { fontSize: 15, bold: true, color: C.ink, valign: 'top' })
  notes(slide, 'The two status errors and CORS behavior were verified against the HTTP API. Six malformed ReviewRequest payloads were verified by Pydantic model validation. Explain that CORS is not authentication: the deployed service has no login and the demo endpoint is public HTTP.')
}

// 17. App 2 deployment
{
  const slide = base('12 · Application 02 · Phase 5 · Workshop 6', 'Build in the cloud, run on Fargate', 'The deployment avoids local Docker and uses a single public task for the demo.')
  const steps = [
    ['GITHUB', 'Source: main'],
    ['CODEBUILD', 'Docker build · buildspec.yml'],
    ['ECR', 'audittrail-ai · scan on push'],
    ['ECS / FARGATE', '1 task · 256 CPU / 512 MiB'],
  ]
  steps.forEach((item, index) => {
    const x = 0.72 + index * 3.17
    node(slide, x, 2.05, 2.67, 1.37, item[0], item[1], index === 3 ? C.pale : C.white, index === 2 ? C.orange : C.forest2)
    if (index < steps.length - 1) arrow(slide, x + 2.72, 2.72, x + 3.06, 2.72)
  })
  box(slide, 0.82, 4.13, 7.25, 1.18, C.ink, C.ink)
  text(slide, 'LIVE URL', 1.12, 4.4, 1.1, 0.2, { fontSize: 8, bold: true, color: C.lime })
  text(slide, 'http://54.82.56.190:8000', 2.28, 4.32, 5.45, 0.36, { fontSize: 18, bold: true, color: C.white })
  box(slide, 8.45, 4.13, 3.85, 1.18, C.paleOrange, 'F0D6C8')
  text(slide, 'DEPLOYMENT LEARNING', 8.74, 4.38, 2.9, 0.2, { fontSize: 8, bold: true, color: C.orange })
  text(slide, 'Fixed unsupported CODEBUILD_ENV_FILE; rebuilt via buildspec override.', 8.74, 4.72, 3.0, 0.42, { fontSize: 10, color: C.ink })
  notes(slide, 'The first CodeBuild run failed because the buildspec assumed CODEBUILD_ENV_FILE existed. The corrected buildspec was passed with a one-run file:// override and succeeded, pushing tag 67a00c21413e. The GitHub main branch was later pushed with the corrected buildspec; do not imply automatic deploy-on-push without webhook/pipeline setup.')
}

// 18. App 2 known issues and showcase
{
  const slide = base('13 · Application 02 · Phase 6 · Showcase', 'Demo-ready is not production-ready', 'Use the showcase to explain current value and the next safe step.')
  box(slide, 0.82, 1.9, 5.42, 3.8, C.pale)
  text(slide, 'WORKING TODAY', 1.12, 2.2, 2.3, 0.22, { fontSize: 8, bold: true, color: C.forest2 })
  bullets(slide, ['Rule-based rate and duplicate flags', 'Search, filters, CSV, alerts, review actions', 'Live dashboard + API health endpoint', 'Keyboard focus and Escape behavior'], 1.13, 2.66, 4.65, { gap: 0.56, fontSize: 11 })
  box(slide, 6.72, 1.9, 5.42, 3.8, C.paleOrange, 'F0D6C8')
  text(slide, 'LIMITATIONS TO DISCLOSE', 7.02, 2.2, 2.8, 0.22, { fontSize: 8, bold: true, color: C.orange })
  bullets(slide, ['Public HTTP; no authentication / TLS', 'In-memory data resets on task restart', 'Single task; no load balancer or autoscaling', 'No external user feedback documented'], 7.03, 2.66, 4.65, { gap: 0.56, fontSize: 11, dotColor: C.orange })
  text(slide, 'Feedback status: [Add real auditor/classmate feedback and resulting iteration.]', 0.94, 6.0, 11.1, 0.3, { fontSize: 12, bold: true, color: C.ink })
  notes(slide, 'The app was functionally tested by its developer, but no external user feedback was recorded in the available project. Do not claim Workshop 6 feedback happened unless you have evidence. Public demo must not be used for real invoice information.')
}

// 19. AI tools
{
  const slide = base('14 · AI role', 'Prompts that produced useful engineering work', 'Specific behavior, constraints, and verification criteria improved the results.')
  const promptCards = [
    ['BUILD', '“Seed four invoices with one clean, two rate violations, and one duplicate.”', 'Outcome: deterministic rule cases with testable expected counts.'],
    ['HARDEN', '“Reject malformed review payloads; verify exact CORS behavior.”', 'Outcome: strict Pydantic model; extra fields rejected with 422.'],
    ['DEPLOY', '“Build in CodeBuild, push to ECR, deploy on Fargate; verify health.”', 'Outcome: cloud build and a checked live service.'],
  ]
  promptCards.forEach((card, index) => {
    const y = 1.88 + index * 1.4
    box(slide, 0.82, y, 11.6, 1.12, C.white)
    pill(slide, card[0], 1.06, y + 0.18, 1.08, index === 1 ? C.paleOrange : C.pale, index === 1 ? C.orange : C.forest2)
    text(slide, card[1], 2.35, y + 0.15, 9.6, 0.32, { fontSize: 13, bold: true, color: C.ink })
    text(slide, card[2], 2.35, y + 0.59, 9.55, 0.25, { fontSize: 10, color: C.muted })
  })
  text(slide, 'Tools: Copilot · Pylance/runtime checks · browser interaction tests · AWS CLI', 0.95, 6.28, 10.8, 0.28, { fontSize: 11, bold: true, color: C.forest2 })
  notes(slide, 'These prompt patterns reflect the documented App #2 work. The App #1 Claude Artifact prompt should be added when available. AI was used to draft and refine code, while runtime/API checks verified claims before deployment.')
}

// 20. Lessons and comparison
{
  const slide = base('14 · AI role', 'The process improved when verification became part of the prompt', 'From describing what to build to defining how to prove it works.')
  box(slide, 0.82, 1.95, 5.45, 3.78, C.white)
  pill(slide, 'APP 01 · WORKSHOPS 1–4', 1.1, 2.25, 2.45, C.paleOrange, C.orange)
  text(slide, '[Add your actual evolution]', 1.1, 2.78, 4.4, 0.42, { fontSize: 19, bold: true, color: C.ink })
  bullets(slide, ['Prompt / prototype approach: [ ]', 'Testing and feedback loop: [ ]', 'What changed by Workshop 4: [ ]'], 1.12, 3.52, 4.6, { gap: 0.52, fontSize: 11 })
  sourceNeeded(slide, 2.75, 5.1, 2.45)
  box(slide, 6.72, 1.95, 5.45, 3.78, C.ink, C.ink)
  pill(slide, 'APP 02 · WORKSHOPS 5–6', 7.0, 2.25, 2.45, '315B49', C.lime)
  text(slide, 'Specify outcomes; test boundaries.', 7.0, 2.78, 4.5, 0.58, { fontSize: 19, bold: true, color: C.white })
  bullets(slide, ['Expected seed/risk counts were explicit.', 'Negative payloads and CORS were tested.', 'Deployment failures informed a revised buildspec.'], 7.02, 3.62, 4.58, { gap: 0.5, fontSize: 11, color: 'DCE8DF', dotColor: C.lime })
  notes(slide, 'The evidence supports App #2 lessons. Do not claim a comparison to App #1 until you add your actual first-project process. Invite the audience to compare prompt specificity, test discipline, and deployment planning across the two journeys.')
}

// 21. Value and cost benefit
{
  const slide = base('15 · Value add & impact', 'Potential value: less repetitive checking, clearer decisions', 'ROI is a hypothesis until time and error-rate baselines are measured.')
  box(slide, 0.82, 1.9, 5.45, 3.55, C.white)
  pill(slide, 'APP 01', 1.1, 2.2, 1.0, C.paleOrange, C.orange)
  text(slide, '[Value proposition]', 1.1, 2.7, 4.35, 0.4, { fontSize: 19, bold: true, color: C.ink })
  bullets(slide, ['Target workflow: [ ]', 'Current baseline: [time / errors]', 'Improvement measured: [ ]'], 1.12, 3.38, 4.55, { gap: 0.48, fontSize: 11 })
  box(slide, 6.72, 1.9, 5.45, 3.55, C.pale)
  pill(slide, 'APP 02', 7.0, 2.2, 1.0, C.pale, C.forest2)
  text(slide, 'Make anomalies visible early', 7.0, 2.7, 4.48, 0.4, { fontSize: 19, bold: true, color: C.ink })
  bullets(slide, ['Rule-based checks are consistent.', 'Explanations support audit follow-up.', 'Human retains final disposition.'], 7.02, 3.38, 4.55, { gap: 0.48, fontSize: 11, dotColor: C.forest2 })
  box(slide, 0.88, 5.78, 11.25, 0.78, C.ink, C.ink)
  text(slide, 'Illustrative only: 100 invoices × 5 min saved × $45/hr = $375 monthly capacity · not measured', 1.15, 6.02, 10.6, 0.24, { fontSize: 13, bold: true, color: C.white })
  notes(slide, 'The ROI scenario is illustrative, not an achieved metric: 100 invoices per month, five minutes saved per invoice, and $45 hourly labor are assumptions. It represents $375 monthly labor capacity before infrastructure cost; validate all three inputs with users. App #1 value needs actual problem and outcome data.')
}

// 22. Celebration checklist
{
  const slide = base('16 · Capstone milestone', 'Celebrate the work—and verify the evidence', 'Mark each item complete only when it is true for your submission.')
  const items = [
    ['Built two complete applications', 'App #1 details needed to document'],
    ['Deployed both applications to AWS', 'App #2 verified · App #1 URL needed'],
    ['Implemented user feedback', 'Add the actual feedback + iteration'],
    ['Created comprehensive documentation', 'Deck + reflection draft created'],
    ['Developed AI partnership skills', 'Evidence: iterative prompting + tests'],
    ['Grown as a developer', 'Personal reflection: add your example'],
    ['Served others through technology', 'Add target-user or feedback evidence'],
    ['Completed the capstone course', 'Confirm course completion'],
  ]
  items.forEach((item, index) => {
    const col = index < 4 ? 0 : 1
    const row = index % 4
    const x = 0.82 + col * 6.0
    const y = 1.88 + row * 0.96
    box(slide, x, y, 5.55, 0.72, C.white)
    slide.addShape(SH.roundRect, { x: x + 0.2, y: y + 0.2, w: 0.23, h: 0.23, rectRadius: 0.03, line: { color: index === 1 && col === 0 ? C.orange : C.forest2, width: 1 }, fill: { color: C.white } })
    text(slide, item[0], x + 0.58, y + 0.1, 4.65, 0.25, { fontSize: 11.5, bold: true, color: C.ink })
    text(slide, item[1], x + 0.58, y + 0.41, 4.65, 0.18, { fontSize: 8.5, color: C.muted })
  })
  text(slide, 'This is a celebration and an evidence checklist—not a claim that every item was independently verified.', 0.92, 6.03, 11.1, 0.25, { fontSize: 10, italic: true, color: C.muted })
  notes(slide, 'Invite the audience to celebrate the work. The checklist mirrors the requested capstone items. App #1 deployment and feedback, user-service evidence, and course completion are not available in the project materials, so confirm those before checking them off.')
}

// 23. Conclusion
{
  const slide = pptx.addSlide()
  slide.background = { color: C.ink }
  slide.addShape(SH.rect, { x: 0, y: 0, w: W, h: 0.12, line: { color: C.lime, transparency: 100 }, fill: { color: C.lime } })
  pill(slide, '17 · CONCLUSION', 0.8, 0.62, 1.72, '315B49', C.lime)
  text(slide, 'Build. Verify. Learn.', 0.8, 1.25, 8.9, 0.78, { fontFace: 'Aptos Display', fontSize: 35, bold: true, color: C.white })
  text(slide, 'Next: persist data · add identity · protect the public endpoint · learn from real auditors', 0.84, 2.25, 10.7, 0.36, { fontSize: 15, color: 'D2E0D7' })
  box(slide, 0.84, 3.22, 5.2, 1.25, '214638', '456A56')
  text(slide, 'APPLICATION 01', 1.15, 3.5, 2.0, 0.2, { fontSize: 8, bold: true, color: 'BDD0C1' })
  text(slide, '[App #1 name] · [ADD VERIFIED APP RUNNER URL]', 1.15, 3.9, 4.45, 0.3, { fontSize: 12, bold: true, color: C.white })
  box(slide, 6.43, 3.22, 5.2, 1.25, '214638', '456A56')
  text(slide, 'APPLICATION 02', 6.74, 3.5, 2.0, 0.2, { fontSize: 8, bold: true, color: 'BDD0C1' })
  text(slide, 'AuditTrail AI · http://54.82.56.190:8000', 6.74, 3.9, 4.45, 0.3, { fontSize: 12, bold: true, color: C.lime })
  text(slide, 'Demo warning: App #2 is public HTTP, unauthenticated, and backed by in-memory demo data.', 0.88, 5.14, 10.9, 0.34, { fontSize: 11, color: 'F3CDBE' })
  text(slide, 'Replace App #1 placeholders and personalize the reflection before submission.', 0.88, 5.72, 10.8, 0.28, { fontSize: 10, color: 'B8CBC0' })
  text(slide, '23 / 23', 11.72, 6.8, 0.72, 0.16, { fontSize: 8, color: 'AABDB2', align: 'right' })
  notes(slide, 'Close by restating the capstone progression and the next steps. Read the App #2 limitation aloud so the demo is not mistaken for a production invoice system. Add the verified App #1 URL before presenting.')
}

const deckPath = path.join(outputDir, 'AuditTrail-AI-Capstone.pptx')
await pptx.writeFile({ fileName: deckPath })
console.log(`Generated ${pptx._slides.length} slides: ${deckPath}`)