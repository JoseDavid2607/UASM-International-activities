import test from 'node:test';
import assert from 'node:assert/strict';
import { parseExcelDate, safeURL, rowToEvent, normKey, listEvidence } from './build-data.mjs';

test('fechas: serial, texto y casos inválidos', () => {
  assert.deepEqual(parseExcelDate(46100), { y: 2026, m: 3, d: 19 });
  assert.deepEqual(parseExcelDate(46100.99), { y: 2026, m: 3, d: 19 });
  assert.deepEqual(parseExcelDate('2026-03-15'), { y: 2026, m: 3, d: 15 });
  assert.deepEqual(parseExcelDate('15/03/2026'), { y: 2026, m: 3, d: 15 });
  assert.equal(parseExcelDate('31/02/2026'), null);
  assert.equal(parseExcelDate('basura'), null);
  assert.equal(parseExcelDate(null), null);
});

test('enlaces: solo http(s)', () => {
  assert.equal(safeURL('javascript:alert(1)'), null);
  assert.equal(safeURL(' https://a.co/x '), 'https://a.co/x');
});

test('fila -> evento: fin anterior se ignora, cancelada se omite, inválida se omite', () => {
  const base = { 'Programa': 'Pregrado', 'Categoría': 'Evento', 'Título': 'X', 'Fecha inicio': 46100 };
  assert.equal(rowToEvent({ ...base, 'Fecha fin (opcional)': 46101 }).d2, 20);
  const w = [];
  assert.equal(rowToEvent({ ...base, 'Fecha fin (opcional)': 46000 }, (m) => w.push(m)).d2, undefined);
  assert.equal(w.length, 1);
  assert.equal(rowToEvent({ ...base, 'Estado': 'Cancelada' }), null);
  assert.equal(rowToEvent({ ...base, 'Categoría': 'Otra' }, () => {}), null);
});

test('evidencias: recorre subcarpetas y agrupa por título normalizado', async () => {
  const tree = {
    ROOT: [{ id: 'A', name: 'Sesiones EPOS', mimeType: 'application/vnd.google-apps.folder' }],
    A: [{ id: 'B', name: 'Sesión Oportunidades Internacionales', mimeType: 'application/vnd.google-apps.folder' }],
    B: [{ id: 'f1', name: 'x.pdf', mimeType: 'application/pdf' }],
  };
  const fake = async (u) => ({ files: tree[decodeURIComponent(u).match(/'([^']+)' in parents/)[1]] || [] });
  const r = await listEvidence('k', 'ROOT', fake);
  assert.deepEqual(Object.keys(r), [normKey('Sesion oportunidades internacionales')]);
  assert.match(r['sesion oportunidades internacionales'][0].thumb, /thumbnail\?id=f1/);
});
