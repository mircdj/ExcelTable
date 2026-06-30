/**
 * Regressioni visive. Prima esecuzione: `npm run e2e:update` genera le
 * baseline in tests/e2e/__screenshots__; le successive confrontano.
 */
import { test, expect } from '@playwright/test';
import { openGrid } from './helpers';

for (const theme of ['excel', 'dark', 'excel']) {
  test(`tema ${theme}: griglia con pinned + totali`, async ({ page }) => {
    await openGrid(page, 'base', { theme });
    await expect(page.locator('.eg-root')).toHaveScreenshot(`base-${theme}.png`);
  });
}

test('gruppi con aggregati', async ({ page }) => {
  await openGrid(page, 'groups');
  await expect(page.locator('.eg-root')).toHaveScreenshot('groups.png');
});

test('pannello detail aperto', async ({ page }) => {
  await openGrid(page, 'detail');
  await page.locator('.eg-expand').first().click();
  await expect(page.locator('.eg-root')).toHaveScreenshot('detail-open.png');
});
