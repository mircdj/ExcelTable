import { test, expect } from '@playwright/test';
import { openGrid, cell } from './helpers';

test.use({ permissions: ['clipboard-read', 'clipboard-write'] });

test('Ctrl+C scrive TSV e HTML; Ctrl+V incolla il blocco', async ({ page, context }) => {
  await openGrid(page);
  await cell(page, 0, 1).click();
  await page.keyboard.press('Shift+ArrowDown');
  await page.keyboard.press('Shift+ArrowRight');
  await page.keyboard.press('Control+c');
  await expect(page.locator('.eg-copy-t').first()).toBeVisible(); // marching ants
  // incolla 2x2 a partire da una cella più in basso
  await cell(page, 5, 1).click();
  await page.keyboard.press('Control+v');
  const v00 = await cell(page, 0, 1).textContent();
  await expect(cell(page, 5, 1)).toHaveText(v00!);
});

test('copia come HTML: gli appunti contengono una <table>', async ({ page }) => {
  await openGrid(page);
  await cell(page, 0, 1).click();
  await page.keyboard.press('Control+c');
  const html = await page.evaluate(async () => {
    const items = await navigator.clipboard.read();
    for (const it of items) {
      if (it.types.includes('text/html')) return await (await it.getType('text/html')).text();
    }
    return '';
  });
  expect(html).toContain('<table>');
});
