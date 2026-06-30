import { test, expect, devices } from '@playwright/test';
import { openGrid, cell, activeCell } from './helpers';

test.use({ ...devices['iPad (gen 7)'], browserName: 'chromium' });

test('tap seleziona; doppio tap modifica', async ({ page }) => {
  await openGrid(page);
  const c = (await cell(page, 1, 1).boundingBox())!;
  await page.touchscreen.tap(c.x + 10, c.y + 10);
  expect((await activeCell(page))!.rowIndex).toBe(1);
  await page.touchscreen.tap(c.x + 10, c.y + 10);
  await page.touchscreen.tap(c.x + 10, c.y + 10);
  await expect(page.locator('.eg-editor')).toBeVisible();
});

test('long-press apre il context menu', async ({ page }) => {
  await openGrid(page);
  const c = (await cell(page, 2, 1).boundingBox())!;
  const client = await page.context().newCDPSession(page);
  await client.send('Input.dispatchTouchEvent', {
    type: 'touchStart',
    touchPoints: [{ x: c.x + 10, y: c.y + 10 }],
  });
  await page.waitForTimeout(600);
  await client.send('Input.dispatchTouchEvent', { type: 'touchEnd', touchPoints: [] });
  await expect(page.locator('.eg-menu')).toBeVisible();
});
