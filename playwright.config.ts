import { defineConfig } from '@playwright/test';

/**
 * Suite E2E in browser reale — copre ciò che jsdom non può: layout
 * (colonne bloccate, gruppi, pannelli detail), tastiera vera, clipboard,
 * drag del mouse, touch e regressioni visive via screenshot.
 *
 *   npx playwright install chromium   # una volta
 *   npm run e2e                       # esegue
 *   npm run e2e:update                # rigenera le baseline screenshot
 */
export default defineConfig({
  testDir: './tests/e2e',
  fullyParallel: true,
  reporter: [['list']],
  use: {
    viewport: { width: 1200, height: 800 },
    // La harness è un file statico: nessun server necessario.
  },
  projects: [
    { name: 'chromium', use: { browserName: 'chromium' } },
  ],
  expect: { toHaveScreenshot: { maxDiffPixelRatio: 0.002 } },
});
