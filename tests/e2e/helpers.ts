import { type Page, expect } from '@playwright/test';
import { pathToFileURL } from 'node:url';
import { resolve } from 'node:path';

export const HARNESS = pathToFileURL(resolve(__dirname, 'harness.html')).href;

export async function openGrid(page: Page, preset = 'base', extra = {}): Promise<void> {
  await page.goto(HARNESS);
  await page.waitForFunction(() => (window as never as { __themeReady: boolean }).__themeReady);
  await page.evaluate(
    ([p, x]) => (window as never as { makeGrid(p: string, x: unknown): boolean }).makeGrid(p as string, x),
    [preset, extra] as const,
  );
  await expect(page.locator('.eg-root')).toBeVisible();
}

export const cell = (page: Page, row: number, col: number) =>
  page.locator(`.eg-cell[data-row="${row}"][data-col="${col}"]`).first();

export async function activeCell(page: Page): Promise<{ rowIndex: number; colId: string } | null> {
  return page.evaluate(() => (window as never as { grid: { activeCell: never } }).grid.activeCell);
}
