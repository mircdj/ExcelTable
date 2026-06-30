import { test, expect } from '@playwright/test';
import { openGrid, cell } from './helpers';

test('drag di selezione con il mouse', async ({ page }) => {
  await openGrid(page);
  const a = (await cell(page, 1, 1).boundingBox())!;
  const b = (await cell(page, 4, 3).boundingBox())!;
  await page.mouse.move(a.x + 5, a.y + 5);
  await page.mouse.down();
  await page.mouse.move(b.x + 5, b.y + 5, { steps: 5 });
  await page.mouse.up();
  const range = await page.evaluate(() => (window as never as { grid: { ranges: { endRow: number; endCol: number }[] } }).grid.ranges[0]);
  expect(range.endRow).toBe(4);
  expect(range.endCol).toBe(3);
});

test('fill handle: trascina e genera la serie', async ({ page }) => {
  await openGrid(page);
  await cell(page, 0, 3).dblclick();
  await page.keyboard.press('Control+a');
  await page.keyboard.type('10');
  await page.keyboard.press('Escape'); // chiude senza muovere il focus
  await cell(page, 0, 3).click();
  await page.keyboard.type('10');
  await page.keyboard.press('Enter');
  await cell(page, 1, 3).click();
  await page.keyboard.type('20');
  await page.keyboard.press('Enter');
  await cell(page, 0, 3).click();
  await page.keyboard.press('Shift+ArrowDown'); // seleziona 10,20
  const handle = page.locator('.eg-fill-handle');
  const h = (await handle.boundingBox())!;
  const target = (await cell(page, 4, 3).boundingBox())!;
  await page.mouse.move(h.x + 3, h.y + 3);
  await page.mouse.down();
  await page.mouse.move(target.x + 10, target.y + 10, { steps: 6 });
  await page.mouse.up();
  await expect(cell(page, 2, 3)).toHaveText('30');
  await expect(cell(page, 4, 3)).toHaveText('50');
});

test('drag & drop header: la colonna si sposta', async ({ page }) => {
  await openGrid(page);
  const from = (await page.locator('.eg-hcell[data-colid="ore"]').boundingBox())!;
  const to = (await page.locator('.eg-hcell[data-colid="nome"]').boundingBox())!;
  await page.mouse.move(from.x + 30, from.y + 10);
  await page.mouse.down();
  await page.mouse.move(to.x + 10, to.y + 10, { steps: 8 });
  await expect(page.locator('.eg-drag-ghost')).toBeVisible();
  await page.mouse.up();
  const order = await page.evaluate(() =>
    (window as never as { grid: { visibleColumns(): { id: string }[] } }).grid.visibleColumns().map((c) => c.id),
  );
  expect(order.indexOf('ore')).toBeLessThan(order.indexOf('nome'));
});

test('context menu: tasto destro apre, Elimina riga funziona', async ({ page }) => {
  await openGrid(page);
  const before = await page.evaluate(() => (window as never as { grid: { totalRowCount: number } }).grid.totalRowCount);
  await cell(page, 0, 1).click({ button: 'right' });
  await expect(page.locator('.eg-menu')).toBeVisible();
  await page.locator('.eg-menu-item', { hasText: 'Elimina riga' }).click();
  const after = await page.evaluate(() => (window as never as { grid: { totalRowCount: number } }).grid.totalRowCount);
  expect(after).toBe(before - 1);
});
