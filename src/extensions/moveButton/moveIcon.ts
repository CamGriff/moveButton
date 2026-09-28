/**
 * The move-page icon, as an inline SVG data URI whose fill color can be
 * swapped at runtime to match the current SharePoint theme.
 *
 * The command bar renders `iconImageUrl` as an <img>, so the SVG's `fill`
 * has to be baked into the markup (CSS like `currentColor` isn't inherited
 * across an <img> boundary) — hence building the string per theme instead
 * of setting a single static asset in the manifest.
 */
const SVG_PATH: string =
  'M6 2a2 2 0 0 0-2 2v5.2c.32-.08.66-.15 1-.18V4a1 1 0 0 1 1-1h4v3.5c0 .83.67 1.5 1.5 1.5H15v8a1 1 0 0 1-1 1h-3.6c-.18.36-.4.7-.66 1H14a2 2 0 0 0 2-2V7.41c0-.4-.16-.78-.44-1.06l-3.91-3.91A1.5 1.5 0 0 0 10.59 2H6Zm8.8 5h-3.3a.5.5 0 0 1-.5-.5V3.2L14.8 7ZM5.5 19a4.5 4.5 0 1 0 0-9 4.5 4.5 0 0 0 0 9Zm.35-2.15a.5.5 0 0 1-.7-.7L6.29 15H3.5a.5.5 0 0 1 0-1h2.8l-1.15-1.15a.5.5 0 0 1 .7-.7l2 2a.5.5 0 0 1 .15.35.5.5 0 0 1-.14.35l-2 2Z';

/** Neutral gray used when no theme is available yet (matches the manifest's static fallback icon). */
export const DEFAULT_ICON_FILL: string = '#323130';

export function getMoveIconDataUri(fillColor: string): string {
  const svg: string =
    `<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 20 20">` +
    `<path fill="${fillColor}" d="${SVG_PATH}"/></svg>`;
  return `data:image/svg+xml;base64,${window.btoa(svg)}`;
}
