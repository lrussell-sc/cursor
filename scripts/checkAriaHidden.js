const { chromium } = require('playwright');

const url = process.argv[2];

if (!url) {
  console.error('Usage: node scripts/checkAriaHidden.js <url>');
  process.exit(1);
}

function trimHtml(html, limit = 300) {
  if (!html) return '';
  const clean = html.replace(/\s+/g, ' ').trim();
  if (clean.length <= limit) return clean;
  return `${clean.slice(0, limit)}…`;
}

async function run() {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1280, height: 720 } });

  await page.goto(url, { waitUntil: 'networkidle' });
  // Give client-side scripts a moment to finish rendering menus, etc.
  await page.waitForTimeout(4000);

  const violations = await page.evaluate(() => {
    const focusableSelector = [
      'a[href]',
      'area[href]',
      'button',
      'input',
      'select',
      'textarea',
      'iframe',
      'audio[controls]',
      'video[controls]',
      'summary',
      '[tabindex]',
      '[contenteditable="true"]'
    ].join(',');

    const isFocusable = (node) => {
      if (!(node instanceof HTMLElement)) return false;
      if (node.hasAttribute('disabled')) return false;
      if (node.matches('a[href],area[href],button,input,select,textarea,iframe,audio[controls],video[controls],summary')) {
        return true;
      }
      if (node.hasAttribute('tabindex')) return true;
      if (node.getAttribute('contenteditable') === 'true') return true;
      return false;
    };

    const getPath = (node) => {
      if (!(node instanceof Element)) return '';
      const parts = [];
      let current = node;
      while (current && parts.length < 8) {
        const tag = current.tagName.toLowerCase();
        if (current.id) {
          parts.unshift(`${tag}#${current.id}`);
          break;
        }
        let nth = 1;
        let sibling = current;
        while ((sibling = sibling.previousElementSibling)) {
          if (sibling.tagName === current.tagName) nth++;
        }
        parts.unshift(`${tag}:nth-of-type(${nth})`);
        current = current.parentElement;
      }
      return parts.join(' > ');
    };

    const toIssue = (type, hiddenNode, focusableNode) => ({
      type,
      ariaHiddenPath: getPath(hiddenNode),
      ariaHiddenHtml: hiddenNode.outerHTML,
      focusablePath: focusableNode ? getPath(focusableNode) : null,
      focusableHtml: focusableNode ? focusableNode.outerHTML : null
    });

    const results = [];

    document.querySelectorAll('[aria-hidden="true"]').forEach((hiddenNode) => {
      if (isFocusable(hiddenNode)) {
        results.push(toIssue('aria-hidden element is focusable', hiddenNode, hiddenNode));
      }

      hiddenNode.querySelectorAll(focusableSelector).forEach((focusableNode) => {
        if (isFocusable(focusableNode)) {
          results.push(toIssue('aria-hidden element contains focusable descendant', hiddenNode, focusableNode));
        }
      });
    });

    return results;
  });

  await browser.close();

  if (!violations.length) {
    console.log('No focusable elements found inside aria-hidden containers.');
    return;
  }

  console.log(`Found ${violations.length} potential issue(s):`);
  violations.forEach((violation, index) => {
    console.log(`\nIssue ${index + 1}: ${violation.type}`);
    console.log(`  aria-hidden node: ${trimHtml(violation.ariaHiddenHtml, 500)}`);
    console.log(`  DOM path: ${violation.ariaHiddenPath}`);
    if (violation.focusableHtml) {
      console.log(`  Focusable node: ${trimHtml(violation.focusableHtml)}`);
      console.log(`  Focusable DOM path: ${violation.focusablePath}`);
    }
  });
}

run().catch((error) => {
  console.error('Failed to complete scan:', error);
  process.exit(1);
});
