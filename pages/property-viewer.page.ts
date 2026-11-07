import { expect, Locator, Page } from '@playwright/test';

export interface ExtractedPropertyRow {
  section: string;
  field: string;
  value: string;
}

export class PropertyViewerPage {
  constructor(private readonly page: Page) { }

  private get searchBox(): Locator {
    return this.page.getByRole('textbox', { name: 'Search' });
  }

  async open(url: string): Promise<void> {
    await this.page.goto(url, { waitUntil: 'domcontentloaded' });
    await expect(this.page).toHaveTitle(/York County Property Viewer/i);
    await expect(this.page).toHaveTitle("York County Property Viewer 2.0");
    console.log("Page title is: " + await this.page.title());

    await this.page.locator('div').filter({ hasText: /^DISCLAIMER$/ }).nth(5).click();
    await this.page.getByRole('checkbox', { name: 'I agree to the above terms' }).check();
    await this.page.getByRole('button', { name: 'OK', exact: true }).click();
    await this.page.waitForTimeout(15000); // Wait for 20 seconds to ensure the page is fully loaded

    await expect(this.searchBox).toBeVisible();
  }

  async searchAndSelectPin(pin: string): Promise<void> {
    await expect(this.searchBox).toBeVisible();
    await this.searchBox.click();
    await this.searchBox.fill(pin);
    await this.searchBox.press('Enter');

    const suggestion = this.page
      .getByRole('group', { name: 'PIDN (13-Digit)' })
      .getByRole('option', { name: pin, exact: true });

    await expect(suggestion).toBeVisible({ timeout: 25_000 });
    // Selecting the suggestion is the only click required. Do not click the
    // same option again; ArcGIS may detach it during its re-render.
    await suggestion.click();

    await expect(
      this.page.getByRole('heading', { name: 'Property Information', exact: true })
    ).toBeVisible({ timeout: 30_000 });
  }

  async extractPropertyRows(): Promise<ExtractedPropertyRow[]> {
    const sections = this.page.locator('arcgis-feature-fields');
    await expect(sections.first()).toBeVisible({ timeout: 20_000 });

    const output: ExtractedPropertyRow[] = [];
    const sectionCount = await sections.count();

    for (let i = 0; i < sectionCount; i++) {
      const sectionLocator = sections.nth(i);
      const heading = sectionLocator.locator('h3[role="heading"], h3').first();
      const sectionName = (await heading.innerText().catch(() => `Section ${i + 1}`)).trim();

      const tables = sectionLocator.locator('table');
      for (let t = 0; t < await tables.count(); t++) {
        const rows = tables.nth(t).locator('tbody tr');
        for (let r = 0; r < await rows.count(); r++) {
          const row = rows.nth(r);
          const fieldLocator = row.locator('th').first();
          const valueLocator = row.locator('td').first();

          const field = (await fieldLocator.innerText().catch(() => '')).trim();
          const value = (await valueLocator.innerText().catch(() => '')).trim();
          if (field || value) {
            output.push({
              section: sectionName || `Section ${i + 1}`,
              field: field || `Row ${r + 1}`,
              value
            });
          }
        }
      }
    }

    if (output.length === 0) {
      throw new Error('No property table rows were found. The page structure may have changed.');
    }
    return output;
  }

  async getPageTitle(): Promise<string> {
    return this.page.title();
  }
}
