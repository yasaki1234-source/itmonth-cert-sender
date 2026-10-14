// 比對邏輯單元測試：node build/test_logic.js
const fs = require('fs');
const vm = require('vm');
const assert = require('assert');
const src = fs.readFileSync(__dirname + '/../apps-script/Code.gs', 'utf8');
const ctx = {};
vm.createContext(ctx);
vm.runInContext(src + '\nthis.__x = {buildList, computeRow, applyPatch, stats, nameFontSize, nextCertNo, normPhone, normEmail, isMark, readInOut, fillText, DEFAULT_CONFIG};', ctx);
const L = ctx.__x;
const cfg = Object.assign({}, L.DEFAULT_CONFIG);
let pass = 0;
function t(name, fn) { fn(); pass++; console.log('✓', name); }
const byName = (rows, n) => rows.find(r => r.name === n);
const finalize = rows => rows.map(r => L.computeRow(r, cfg));

const reg = [
  { name: '王小明', email: 'Ming@Example.com ', school: '臺灣大學', phone: '0912-345-678', identity: '大專院校學生' },
  { name: '李大華', email: 'hua@ntu.edu.tw', school: '政治大學', phone: '0922111222', identity: '碩士生' },
  { name: '陳 美玲', email: '', school: '成功大學', phone: '+886933444555', identity: '博士班' },
  { name: '張老師', email: 'teacher@school.edu.tw', school: '某高中', phone: '', identity: '教師' },
  { name: '王小明', email: 'ming@example.com', school: '臺灣大學', phone: '', identity: '大專院校學生' }, // 重複
  { name: '林志豪', email: 'bad-email', school: '清華大學', phone: '0955666777', identity: '大學生' }
];
const survey = [
  { email: 'MING@example.com', name: '王小明', phone: '' },
  { email: 'other@gmail.com', name: '李大華', phone: '0922-111-222' }, // 用手機對到
  { email: '', name: '陳美玲', phone: '' },                           // 用姓名對到
  { email: 'x@y.com', name: '路人甲', phone: '' }
];

t('正規化', () => {
  assert.strictEqual(L.normPhone('+886 912-345-678'), '0912345678');
  assert.strictEqual(L.normPhone('912345678'), '0912345678');
  assert.strictEqual(L.normEmail(' Ming＠Example.COM '), 'ming@example.com');
  assert.ok(L.isMark('✓') && L.isMark('10:32') && !L.isMark('') && !L.isMark('否') && !L.isMark('x'));
});

let rows;
t('紙本模式：篩身分、去重、問卷比對', () => {
  rows = finalize(L.buildList({ mode: 'match', reg, filterIdentity: true, checkin: null, survey }, cfg, []));
  assert.strictEqual(rows.length, 4, '老師應被篩掉、王小明重複應合併');
  assert.deepStrictEqual(JSON.parse(JSON.stringify(rows.map(r => r.seq))), ['R001', 'R002', 'R003', 'R004']);
  assert.ok(byName(rows, '王小明').note.includes('報名重複'));
  assert.strictEqual(byName(rows, '王小明').survey, '✓');
  assert.strictEqual(byName(rows, '李大華').survey, '✓');
  assert.strictEqual(byName(rows, '陳 美玲').survey, '✓');
  assert.ok(byName(rows, '陳 美玲').note.includes('問卷以姓名比對'));
  assert.strictEqual(byName(rows, '林志豪').survey, '');
  rows.forEach(r => assert.strictEqual(r.result, '不符合', '尚未勾簽到 → 不符合'));
  rows.forEach(r => assert.strictEqual(r.approved, ''));
});

t('網頁勾選紙本簽到退', () => {
  const a = byName(rows, '王小明'); L.applyPatch(a, { inManual: 'Y', outManual: 'Y' }); L.computeRow(a, cfg);
  assert.strictEqual(a.result, '符合'); assert.strictEqual(a.approved, '✓');
  const b = byName(rows, '李大華'); L.applyPatch(b, { inManual: 'Y' }); L.computeRow(b, cfg);
  assert.strictEqual(b.result, '待確認'); assert.ok(b.note.includes('未簽退'));
  const c = byName(rows, '陳 美玲'); L.applyPatch(c, { inManual: 'Y', outManual: 'Y' }); L.computeRow(c, cfg);
  assert.strictEqual(c.result, '待確認', '缺 Email'); assert.ok(c.note.includes('缺 Email'));
  L.applyPatch(c, { email: 'mei@ncku.edu.tw' }); L.computeRow(c, cfg);
  assert.strictEqual(c.result, '符合');
  const d = byName(rows, '林志豪'); L.applyPatch(d, { inManual: 'Y', outManual: 'Y', approve: 'Y' }); L.computeRow(d, cfg);
  assert.strictEqual(d.result, '待確認'); assert.strictEqual(d.approved, '✓', '人工強制核發');
  assert.throws(() => L.applyPatch(d, { email: 'still-bad' }));
});

t('重新上傳：保留編號、勾選、寄送狀態；新人接續編號', () => {
  const a = byName(rows, '王小明'); a.status = '已寄出'; a.certNo = 'ITM2026-001'; a.sentAt = '2026/11/20';
  const reg2 = reg.concat([{ name: '黃新生', email: 'new@nccu.edu.tw', school: '政治大學', phone: '', identity: '大學部' }]);
  const rows2 = finalize(L.buildList({ mode: 'match', reg: reg2, filterIdentity: true, checkin: null, survey }, cfg, rows));
  assert.strictEqual(rows2.length, 5);
  const a2 = byName(rows2, '王小明');
  assert.strictEqual(a2.status, '已寄出'); assert.strictEqual(a2.certNo, 'ITM2026-001'); assert.strictEqual(a2.seq, 'R001');
  assert.strictEqual(byName(rows2, '李大華').inManual, 'Y');
  assert.strictEqual(byName(rows2, '陳 美玲').email, 'mei@ncku.edu.tw', '網頁改過的 Email 要保留');
  assert.strictEqual(byName(rows2, '黃新生').seq, 'R005');
  const s = L.stats(rows2);
  assert.strictEqual(s.sent, 1); assert.strictEqual(s.toSend, 2, '陳美玲 + 林志豪(強制)');
  rows = rows2;
});

t('上傳謄打檔：編號比對、類型欄、現場未報名', () => {
  const checkin = [
    { code: 'R002', name: '李大華', type: '簽到' },
    { code: 'R002', name: '李大華', type: '簽退' },
    { code: '', name: '黃新生', type: '簽到' },
    { code: '', name: '黃新生', type: '簽退' },
    { code: '', name: '路過生', email: 'walk@in.com', type: '簽到' }
  ];
  const fresh = finalize(L.buildList({ mode: 'match', reg: reg.concat([{ name: '黃新生', email: 'new@nccu.edu.tw', identity: '大學部' }]), filterIdentity: true, checkin, checkinMap: { type: true }, survey }, cfg, []));
  const b = byName(fresh, '李大華');
  assert.strictEqual(b.checkIn + b.checkOut, '✓✓'); assert.strictEqual(b.result, '符合');
  const h = byName(fresh, '黃新生');
  assert.strictEqual(h.checkIn, '✓'); assert.strictEqual(h.result, '待確認'); assert.ok(h.note.includes('姓名比對'));
  const w = byName(fresh, '路過生');
  assert.strictEqual(w.source, '現場（未報名）'); assert.strictEqual(w.result, '待確認'); assert.strictEqual(w.approved, '');
});

t('勾選欄格式（簽到核對/簽退核對）', () => {
  const io = L.readInOut({ in: 'V', out: '' }, { in: true, out: true });
  assert.deepStrictEqual(JSON.parse(JSON.stringify(io)), { in: true, out: false });
  const io2 = L.readInOut({}, {});
  assert.deepStrictEqual(JSON.parse(JSON.stringify(io2)), { in: true, out: true });
});

t('設定：不核對簽到 → 只看問卷', () => {
  const c2 = Object.assign({}, cfg, { requireCheckin: false });
  const r = L.buildList({ mode: 'match', reg, filterIdentity: true, checkin: null, survey }, c2, []).map(x => L.computeRow(x, c2));
  assert.strictEqual(byName(r, '王小明').result, '符合');
  assert.strictEqual(byName(r, '林志豪').result, '不符合');
});

t('直接上傳核發名單模式', () => {
  const r = finalize(L.buildList({ mode: 'direct', reg: [{ name: 'A', email: 'a@a.com' }, { name: 'B', email: 'b@b.com' }] }, cfg, []));
  assert.ok(r.every(x => x.result === '符合' && x.approved === '✓'));
});

t('證書編號、姓名字級、信件套字', () => {
  assert.strictEqual(L.nextCertNo([{ certNo: 'ITM2026-007' }, { certNo: '' }], 'ITM2026-'), 'ITM2026-008');
  assert.strictEqual(L.nameFontSize('王小明', 50), 50);
  assert.ok(L.nameFontSize('Alexander Christopher Johnson-Smith', 50) < 30);
  assert.strictEqual(L.fillText('{{姓名}} 同學｜{{證書編號}}', { name: '王小明', certNo: 'ITM2026-001' }, cfg), '王小明 同學｜ITM2026-001');
});

console.log(`\n全部通過：${pass} 項`);
