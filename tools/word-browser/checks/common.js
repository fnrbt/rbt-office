// Shared by the checks: a tiny assertion log and the module loaders.
export const log = [];
export const counts = { passed: 0, failed: 0 };
export const check = (label, got, want) => {
    const g = JSON.stringify(got), w = JSON.stringify(want);
    if (g === w) { counts.passed++; log.push('ok   ' + label); } else { counts.failed++; log.push(`FAIL ${label}: got ${g} want ${w}`); }
};
export const note = (text) => log.push('ok   (' + text + ')');
export const finish = () => { window.result = { log, passed: counts.passed, failed: counts.failed }; };
export const fail = (e) => { counts.failed++; log.push('threw ' + String(e && e.stack || e).split('\n').slice(0, 4).join(' | ')); };
export const loadWord = async () => (await import('/word/Word.js')).default;
export const loadFormats = async () => { const main = await import('/office-formats/src/Main.js'); return typeof main.default === 'function' ? main.default() : main.default; };
export const bytesOf = async (url) => new Uint8Array(await (await fetch(url)).arrayBuffer());
