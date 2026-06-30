/**
 * Regressioni di LAYOUT — esattamente la categoria che jsdom non vede.
 * Include il test che avrebbe intercettato il bug "prima cella una riga
 * più in basso" delle colonne bloccate.
 */
import { test, expect } from '@playwright/test';
import { openGrid, cell } from './helpers';

test('header bloccato allineato alla riga header (bug pinned, regressione)', async ({ page }) => {
  await openGrid(page);
  const headerRow = page.locator('.eg-header-row');
  const pinnedHeader = page.locator('.eg-hcell[data-colid="id"]');
  const hr = (await headerRow.boundingBox())!;
  const ph = (await pinnedHeader.boundingBox())!;
  // L'header pinnato deve stare DENTRO la riga header, non una riga sotto
  expect(Math.abs(ph.y - hr.y)).toBeLessThan(2);
  expect(ph.height).toBeLessThanOrEqual(hr.height + 1);
});

test('celle bloccate allineate alla propria riga dati', async ({ page }) => {
  await openGrid(page);
  for (const r of [0, 3, 7]) {
    const pinned = (await cell(page, r, 0).boundingBox())!;
    const center = (await cell(page, r, 1).boundingBox())!;
    expect(Math.abs(pinned.y - center.y)).toBeLessThan(2);
  }
});

test('scroll orizzontale: le colonne bloccate restano ferme', async ({ page }) => {
  await openGrid(page);
  const before = (await cell(page, 0, 0).boundingBox())!;
  await page.locator('.eg-viewport').evaluate((vp) => (vp.scrollLeft = 300));
  await page.waitForTimeout(80);
  const after = (await cell(page, 0, 0).boundingBox())!;
  expect(Math.abs(after.x - before.x)).toBeLessThan(2);
});

test('riga totali Σ fissa in basso e dentro il viewport', async ({ page }) => {
  await openGrid(page);
  const totals = page.locator('.eg-totals-row');
  await expect(totals).toBeVisible();
  const root = (await page.locator('.eg-root').boundingBox())!;
  const t = (await totals.boundingBox())!;
  expect(t.y + t.height).toBeLessThanOrEqual(root.y + root.height + 1);
});

test('pannello detail: dimensione adattiva e nessuna scrollbar superflua', async ({ page }) => {
  await openGrid(page, 'detail');
  await page.locator('.eg-expand').first().click();
  const nestedViewport = page.locator('.eg-detail .eg-viewport');
  await expect(nestedViewport).toBeVisible();
  const noScroll = await nestedViewport.evaluate(
    (vp) => vp.scrollHeight <= vp.clientHeight + 1 && vp.scrollWidth <= vp.clientWidth + 1,
  );
  expect(noScroll).toBe(true);
});
