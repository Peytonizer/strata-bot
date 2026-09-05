/**
 * The light/dark toggle, for sites in the family that don't already have one of their own.
 *
 * lodger and former each carry their own copy inside their app code and do not import this;
 * it exists for the hub and for future static sites, which have no bundler and no app to hang
 * it off. Load it as a module and it wires itself to any [data-theme-toggle] button:
 *
 *   <script type="module" src="/assets/theme-toggle.js"></script>
 *
 * The choice is remembered in localStorage under `strata-theme`. That is per-site storage, not
 * shared across the family's subdomains — a browser gives each origin its own — so choosing
 * dark on the hub does not carry into lodger. Living with that is deliberate: sharing it would
 * need a cookie on the parent domain or a cross-origin frame, and neither is worth doing to a
 * pair of sites whose whole claim is that they store nothing about you.
 */
const STORAGE_KEY = 'strata-theme';

function systemPrefersDark() {
  return globalThis.matchMedia?.('(prefers-color-scheme: dark)').matches ?? false;
}

function stored() {
  try {
    const value = localStorage.getItem(STORAGE_KEY);
    return value === 'dark' || value === 'light' ? value : null;
  } catch {
    // Private windows and blocked site data throw on access rather than returning null.
    return null;
  }
}

function apply(theme, buttons) {
  document.documentElement.dataset.theme = theme;
  for (const button of buttons) {
    button.setAttribute('aria-pressed', String(theme === 'dark'));
    const label = button.querySelector('[data-theme-label]');
    if (label) label.textContent = theme === 'dark' ? 'Dark' : 'Light';
  }
}

export function initThemeToggle() {
  const buttons = [...document.querySelectorAll('[data-theme-toggle]')];
  if (buttons.length === 0) return;

  let theme = stored() ?? (systemPrefersDark() ? 'dark' : 'light');
  apply(theme, buttons);

  for (const button of buttons) {
    button.addEventListener('click', () => {
      theme = theme === 'dark' ? 'light' : 'dark';
      apply(theme, buttons);
      try {
        localStorage.setItem(STORAGE_KEY, theme);
      } catch {
        // Nothing to do — the choice simply won't survive the tab.
      }
    });
  }
}

initThemeToggle();
