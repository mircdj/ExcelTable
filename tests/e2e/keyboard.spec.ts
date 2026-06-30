import { test, expect } from '@playwright/test';
import { openGrid, cell, activeCell } from './helpers';

test('frecce, Ctrl+frecce, Tab/Invio, Ctrl+A', async ({ page }) => {
  await openGrid(page);
  await cell(page, 0, 1).click();
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowRight');
  expect(await activeCell(page)).toEqual({ rowIndex: 1, colId: 'zona' });
  await page.keyboard.press('Control+ArrowDown'); // salto al bordo dati
  expect((await activeCell(page))!.rowIndex).toBe(199);
  await page.keyboard.press('Control+Home');
  expect((await activeCell(page))!.rowIndex).toBe(0);
  await page.keyboard.press('Control+a');
  const ranges = await page.evaluate(() => (window as never as { grid: { ranges: unknown[] } }).grid.ranges);
  expect(ranges).toHaveLength(1);
});

test('digitare sovrascrive, Invio conferma e scende, Esc annulla', async ({ page }) => {
  await openGrid(page);
  await cell(page, 0, 1).click();
  await page.keyboard.type('Nuovo nome');
  await page.keyboard.press('Enter');
  await expect(cell(page, 0, 1)).toHaveText('Nuovo nome');
  expect((await activeCell(page))!.rowIndex).toBe(1);
  await page.keyboard.press('F2');
  await page.keyboard.type('xxx');
  await page.keyboard.press('Escape');
  await expect(cell(page, 1, 1)).not.toContainText('xxx');
});

test('Ctrl+Z annulla, Ctrl+Y ripristina', async ({ page }) => {
  await openGrid(page);
  await cell(page, 0, 3).dblclick();
  await page.keyboard.press('Control+a');
  await page.keyboard.type('42');
  await page.keyboard.press('Enter');
  await expect(cell(page, 0, 3)).toHaveText('42');
  await page.keyboard.press('Control+z');
  await expect(cell(page, 0, 3)).toHaveText('0');
  await page.keyboard.press('Control+y');
  await expect(cell(page, 0, 3)).toHaveText('42');
});

test('Ctrl+F apre trova, naviga i risultati', async ({ page }) => {
  await openGrid(page);
  await cell(page, 0, 1).click();
  await page.keyboard.press('Control+f');
  await expect(page.locator('.eg-find')).toBeVisible();
  await page.locator('.eg-find-q').fill('Verdi 2');
  await expect(page.locator('.eg-find-count')).toContainText('1 /');
  await page.keyboard.press('Escape');
  await expect(page.locator('.eg-find')).toBeHidden();
});
