import * as React from 'react';
import styles from './Scriptinjector.module.scss';
import { DisplayMode } from '@microsoft/sp-core-library';
import type { IScriptinjectorProps } from './IScriptinjectorProps';

const Scriptinjector: React.FC<IScriptinjectorProps> = (props) => {
  const {
    htmlInput, htmlFileName,
    hasTeamsContext, instanceId, displayMode
  } = props;

  const isEditMode        = displayMode === DisplayMode.Edit;
  const isHtmlInlineEmpty = !htmlInput    || htmlInput.trim()    === '';
  const isHtmlFileEmpty   = !htmlFileName || htmlFileName.trim() === '';
  const nothingProvided   = isHtmlInlineEmpty && isHtmlFileEmpty;

  const htmlContainerRef  = React.useRef<HTMLDivElement>(null);
  const [htmlReady, setHtmlReady] = React.useState<boolean>(false);
  const pendingScriptsRef = React.useRef<string[]>([]);

  // ── Execute JS in global scope — alert works ───────────────────────────
  const executeScript = React.useCallback((code: string, source: string): void => {
    if (!code || code.trim() === '') return;
    try {
      // eslint-disable-next-line no-eval
      (0, eval)(code.trim());
      console.info(`[ScriptInjector:${instanceId}] JS executed from ${source}`);
    } catch (err) {
      console.error(`[ScriptInjector:${instanceId}] JS error from ${source}:`, err);
    }
  }, [instanceId]);

  // ── Strip inline handlers using DOMParser ─────────────────────────────
  const stripInlineHandlers = React.useCallback((html: string): string => {
    const eventAttrs: string[] = [
      'onclick', 'onchange', 'onkeyup', 'onkeydown', 'onkeypress',
      'onmouseover', 'onmouseout', 'onmouseenter', 'onmouseleave',
      'onfocus', 'onblur', 'oninput', 'onsubmit', 'ondblclick',
      'oncontextmenu', 'onscroll', 'onmousedown', 'onmouseup',
      'onload', 'onerror', 'ondragstart', 'ondragend', 'ondrop',
      'ondragover', 'ondragleave'
    ];
    const parser: DOMParser    = new DOMParser();
    const doc: Document        = parser.parseFromString('<div id="__root__">' + html + '</div>', 'text/html');
    const root: Element | null = doc.getElementById('__root__');
    if (!root) return html;
    const allEls: HTMLElement[] = [].slice.call(root.querySelectorAll('*'));
    for (let i: number = 0; i < allEls.length; i++) {
      const el: HTMLElement = allEls[i];
      for (let j: number = 0; j < eventAttrs.length; j++) {
        const attr: string       = eventAttrs[j];
        const val: string | null = el.getAttribute(attr);
        if (val !== null) {
          el.setAttribute('data-handler-' + attr, val);
          el.removeAttribute(attr);
        }
      }
    }
    return root.innerHTML;
  }, []);

  // ── Re-attach events from data-handler-* attributes ───────────────────
  const reattachEvents = React.useCallback((container: HTMLElement): void => {
    const eventAttrs: string[] = [
      'onclick', 'onchange', 'onkeyup', 'onkeydown', 'onkeypress',
      'onmouseover', 'onmouseout', 'onmouseenter', 'onmouseleave',
      'onfocus', 'onblur', 'oninput', 'onsubmit', 'ondblclick',
      'oncontextmenu', 'onscroll', 'onmousedown', 'onmouseup',
      'onload', 'onerror', 'ondragstart', 'ondragend', 'ondrop',
      'ondragover', 'ondragleave'
    ];
    const eventNameMap: { [key: string]: string } = {
      'onclick': 'click', 'onchange': 'change', 'onkeyup': 'keyup',
      'onkeydown': 'keydown', 'onkeypress': 'keypress',
      'onmouseover': 'mouseover', 'onmouseout': 'mouseout',
      'onmouseenter': 'mouseenter', 'onmouseleave': 'mouseleave',
      'onfocus': 'focus', 'onblur': 'blur', 'oninput': 'input',
      'onsubmit': 'submit', 'ondblclick': 'dblclick',
      'oncontextmenu': 'contextmenu', 'onscroll': 'scroll',
      'onmousedown': 'mousedown', 'onmouseup': 'mouseup',
      'onload': 'load', 'onerror': 'error',
      'ondragstart': 'dragstart', 'ondragend': 'dragend',
      'ondrop': 'drop', 'ondragover': 'dragover', 'ondragleave': 'dragleave'
    };
    const allElements: HTMLElement[] = [].slice.call(container.querySelectorAll('*'));
    for (let i: number = 0; i < allElements.length; i++) {
      const el: HTMLElement = allElements[i];
      for (let j: number = 0; j < eventAttrs.length; j++) {
        const attr: string      = eventAttrs[j];
        const dataAttr: string  = 'data-handler-' + attr;
        const handlerCode: string | null = el.getAttribute(dataAttr);
        if (handlerCode && handlerCode.trim() !== '') {
          el.removeAttribute(dataAttr);
          (function(element: HTMLElement, code: string, eventName: string): void {
            element.addEventListener(eventName, function(event: Event): void {
              try {
                // eslint-disable-next-line no-new-func
                const fn: Function = new Function('event', 'with(window){ ' + code + ' }');
                fn.call(element, event);
              } catch (err) {
                console.error(`[ScriptInjector:${instanceId}] Event error [${eventName}]:`, err);
              }
            });
          })(el, handlerCode, eventNameMap[attr]);
        }
      }
      if (el.tagName === 'A') {
        const href: string | null = el.getAttribute('href');
        if (href && href.trim().toLowerCase().indexOf('javascript:') === 0) {
          const jsCode: string = href.trim().substring(11);
          el.removeAttribute('href');
          el.setAttribute('style', (el.getAttribute('style') || '') + ';cursor:pointer;');
          (function(element: HTMLElement, code: string): void {
            element.addEventListener('click', function(event: Event): void {
              event.preventDefault();
              try {
                // eslint-disable-next-line no-eval
                (0, eval)(code);
              } catch (err) {
                console.error(`[ScriptInjector:${instanceId}] href JS error:`, err);
              }
            });
          })(el, jsCode);
        }
      }
    }
    console.info(`[ScriptInjector:${instanceId}] Events re-attached`);
  }, [instanceId]);

  // ── MutationObserver — intercepts dynamically added inline handlers ────
  const startMutationObserver = React.useCallback((container: HTMLElement): void => {
    const eventAttrs: string[] = [
      'onclick', 'onchange', 'onkeyup', 'onkeydown', 'onkeypress',
      'onmouseover', 'onmouseout', 'onmouseenter', 'onmouseleave',
      'onfocus', 'onblur', 'oninput', 'onsubmit', 'ondblclick',
      'oncontextmenu', 'onscroll', 'onmousedown', 'onmouseup',
      'onload', 'onerror', 'ondragstart', 'ondragend', 'ondrop',
      'ondragover', 'ondragleave'
    ];
    const eventNameMap: { [key: string]: string } = {
      'onclick': 'click', 'onchange': 'change', 'onkeyup': 'keyup',
      'onkeydown': 'keydown', 'onkeypress': 'keypress',
      'onmouseover': 'mouseover', 'onmouseout': 'mouseout',
      'onmouseenter': 'mouseenter', 'onmouseleave': 'mouseleave',
      'onfocus': 'focus', 'onblur': 'blur', 'oninput': 'input',
      'onsubmit': 'submit', 'ondblclick': 'dblclick',
      'oncontextmenu': 'contextmenu', 'onscroll': 'scroll',
      'onmousedown': 'mousedown', 'onmouseup': 'mouseup',
      'onload': 'load', 'onerror': 'error',
      'ondragstart': 'dragstart', 'ondragend': 'dragend',
      'ondrop': 'drop', 'ondragover': 'dragover', 'ondragleave': 'dragleave'
    };
    const processElement = (el: HTMLElement): void => {
      for (let j: number = 0; j < eventAttrs.length; j++) {
        const attr: string = eventAttrs[j];
        const handlerCode: string | null = el.getAttribute(attr);
        if (handlerCode && handlerCode.trim() !== '') {
          el.removeAttribute(attr);
          (function(element: HTMLElement, code: string, eventName: string): void {
            element.addEventListener(eventName, function(event: Event): void {
              try {
                // eslint-disable-next-line no-new-func
                const fn: Function = new Function('event', 'with(window){ ' + code + ' }');
                fn.call(element, event);
              } catch (err) {
                console.error(`[ScriptInjector:${instanceId}] Dynamic event error [${eventName}]:`, err);
              }
            });
          })(el, handlerCode, eventNameMap[attr]);
        }
      }
      if (el.tagName === 'A') {
        const href: string | null = el.getAttribute('href');
        if (href && href.trim().toLowerCase().indexOf('javascript:') === 0) {
          const jsCode: string = href.trim().substring(11);
          el.removeAttribute('href');
          el.setAttribute('style', (el.getAttribute('style') || '') + ';cursor:pointer;');
          (function(element: HTMLElement, code: string): void {
            element.addEventListener('click', function(event: Event): void {
              event.preventDefault();
              try {
                // eslint-disable-next-line no-eval
                (0, eval)(code);
              } catch (err) {
                console.error(`[ScriptInjector:${instanceId}] href JS error:`, err);
              }
            });
          })(el, jsCode);
        }
      }
    };
    const observer: MutationObserver = new MutationObserver(
      function(mutations: MutationRecord[]): void {
        for (let m: number = 0; m < mutations.length; m++) {
          const mutation: MutationRecord = mutations[m];
          const added: NodeList = mutation.addedNodes;
          for (let n: number = 0; n < added.length; n++) {
            const node: Node = added[n];
            if (node.nodeType === 1) {
              const el: HTMLElement = node as HTMLElement;
              processElement(el);
              const children: HTMLElement[] = [].slice.call(el.querySelectorAll('*'));
              for (let c: number = 0; c < children.length; c++) {
                processElement(children[c]);
              }
            }
          }
          if (mutation.type === 'attributes') {
            const target: HTMLElement = mutation.target as HTMLElement;
            const attrName: string    = mutation.attributeName || '';
            if (eventAttrs.indexOf(attrName) !== -1) {
              processElement(target);
            }
          }
        }
      }
    );
    observer.observe(container, {
      childList: true, subtree: true,
      attributes: true, attributeFilter: eventAttrs
    });
    observer.observe(document.body, {
      childList: true, subtree: true,
      attributes: true, attributeFilter: eventAttrs
    });
    console.info(`[ScriptInjector:${instanceId}] MutationObserver started`);
  }, [instanceId]);

  // ── Load external CSS — fetch content if link tag blocked by CSP ───────
  const loadExternalCss = React.useCallback((href: string): void => {
    if (!href || href.trim() === '') return;
    const existing: Element | null = document.querySelector(`link[data-injector-css="${href}"]`);
    if (existing) return;
    const link: HTMLLinkElement = document.createElement('link');
    link.rel  = 'stylesheet';
    link.type = 'text/css';
    link.setAttribute('data-injector-css', href);
    link.href = href;
    link.onerror = (): void => {
      console.warn(`[ScriptInjector:${instanceId}] CSS link blocked: ${href} — fetching directly`);
      fetch(href, { method: 'GET' })
        .then(function(r: Response): Promise<string> {
          if (!r.ok) throw new Error('fetch failed');
          return r.text();
        })
        .then(function(css: string): void {
          const baseUrl: string = href.substring(0, href.lastIndexOf('/') + 1);
          const fixedCss: string = css.replace(
            /url\(\s*['"]?(?!https?:\/\/)(?!data:)([^'")]+)['"]?\s*\)/gi,
            function(_match: string, path: string): string {
              return 'url("' + baseUrl + path + '")';
            }
          );
          const alreadyFetched: Element | null = document.querySelector(`style[data-injector-fetched-css="${href}"]`);
          if (alreadyFetched) return;
          const styleTag: HTMLStyleElement = document.createElement('style');
          styleTag.setAttribute('data-injector-fetched-css', href);
          styleTag.textContent = fixedCss;
          document.head.appendChild(styleTag);
          console.info(`[ScriptInjector:${instanceId}] CSS fetched and injected: ${href}`);
        })
        .catch(function(): void {
          console.warn(`[ScriptInjector:${instanceId}] CSS fetch also failed: ${href}`);
        });
    };
    document.head.appendChild(link);
  }, [instanceId]);

  // ── Inject HTML safely ─────────────────────────────────────────────────
  const injectHtmlSafely = React.useCallback((html: string, scripts: string[]): void => {
    if (!htmlContainerRef.current) return;
    const safeHtml: string = stripInlineHandlers(html);
    htmlContainerRef.current.innerHTML = safeHtml;
    reattachEvents(htmlContainerRef.current);
    startMutationObserver(htmlContainerRef.current);
    setTimeout((): void => {
      for (let i: number = 0; i < scripts.length; i++) {
        if (scripts[i].trim() !== '') {
          executeScript(scripts[i], 'html-file-script');
        }
      }
    }, 300);
  }, [stripInlineHandlers, reattachEvents, startMutationObserver, executeScript]);

  // ── Load external script ───────────────────────────────────────────────
  const loadExternalScript = React.useCallback((src: string): Promise<void> => {
    return new Promise<void>((resolve: () => void) => {
      const existing: Element | null = document.querySelector(`script[data-injector-ext="${src}"]`);
      if (existing) { resolve(); return; }
      const script: HTMLScriptElement = document.createElement('script');
      script.src = src;
      script.setAttribute('data-injector-ext', src);
      script.onload  = (): void => { resolve(); };
      script.onerror = (): void => { resolve(); };
      document.head.appendChild(script);
    });
  }, []);



  // ── Process full HTML file ─────────────────────────────────────────────
  const processHtmlFile = React.useCallback(async (html: string, source: string): Promise<void> => {
    const parser: DOMParser = new DOMParser();
    const doc: Document     = parser.parseFromString(html, 'text/html');

    // Step 1 — load all external CSS link tags
    const linkTags: HTMLLinkElement[] = [].slice.call(doc.querySelectorAll('link[rel="stylesheet"]'));
    for (let i: number = 0; i < linkTags.length; i++) {
      loadExternalCss(linkTags[i].href);
    }

    // Step 2 — load external scripts
    const externalScripts: HTMLScriptElement[] = [].slice.call(doc.querySelectorAll('script[src]'));
    for (let i: number = 0; i < externalScripts.length; i++) {
      await loadExternalScript(externalScripts[i].src);
    }

    // Step 3 — apply inline style tags
    const styleTags: HTMLStyleElement[] = [].slice.call(doc.querySelectorAll('style'));
    for (let i: number = 0; i < styleTags.length; i++) {
      const existing: Element | null = document.querySelector(`style[data-injector-html-css="${instanceId}-${i}"]`);
      if (existing) existing.remove();
      const styleTag: HTMLStyleElement = document.createElement('style');
      styleTag.setAttribute('data-injector-html-css', `${instanceId}-${i}`);
      styleTag.textContent = styleTags[i].textContent || '';
      document.head.appendChild(styleTag);
    }

    // Step 4 — collect inline scripts
    const inlineScripts: HTMLScriptElement[] = [].slice.call(doc.querySelectorAll('script:not([src])'));
    const scriptCodes: string[] = [];
    for (let i: number = 0; i < inlineScripts.length; i++) {
      const code: string = inlineScripts[i].textContent || '';
      if (code.trim() !== '') scriptCodes.push(code);
    }

    pendingScriptsRef.current = scriptCodes;
    setHtmlReady(true);
    console.info(`[ScriptInjector:${instanceId}] HTML processed from ${source}`);
  }, [instanceId, loadExternalScript, loadExternalCss]);

  // ── When htmlReady fires inject HTML safely ────────────────────────────
  React.useEffect((): void => {
    if (htmlReady && htmlContainerRef.current && !isHtmlFileEmpty) {
      fetch(htmlFileName.trim(), {
        method: 'GET',
        credentials: 'same-origin',
        headers: { 'Accept': 'text/html, */*' },
      })
      .then(function(r: Response) { return r.text(); })
      .then(function(html: string): void {
        const parser: DOMParser   = new DOMParser();
        const doc: Document       = parser.parseFromString(html, 'text/html');
        const bodyContent: string = doc.body ? doc.body.innerHTML : '';
        injectHtmlSafely(bodyContent, pendingScriptsRef.current);
      })
      .catch(function(err: unknown): void {
        console.error(`[ScriptInjector:${instanceId}] Re-fetch error:`, err);
      });
    }
  }, [htmlReady]);

  // ── HTML file effect ───────────────────────────────────────────────────
  React.useEffect((): void => {
    if (!isHtmlFileEmpty) {
      setHtmlReady(false);
      fetch(htmlFileName.trim(), {
        method: 'GET',
        credentials: 'same-origin',
        headers: { 'Accept': 'text/html, */*' },
      })
      .then(function(r: Response) { return r.text(); })
      .then(function(html: string): void { processHtmlFile(html, htmlFileName).catch(console.error); })
      .catch(function(err: unknown): void { console.error(`[ScriptInjector:${instanceId}] HTML fetch error:`, err); });
    }
  }, [htmlFileName]);

  // ── Inline HTML effect ─────────────────────────────────────────────────
  React.useEffect((): void => {
    if (!isHtmlInlineEmpty && htmlContainerRef.current) {
      const raw: string = htmlInput.trim();
      const hasAnyHtmlTag: boolean = /<[a-zA-Z][\s\S]*?>/i.test(raw);

      if (!hasAnyHtmlTag) {
        const looksLikeCss: boolean = /[a-zA-Z0-9#.\[\]*:@,\s>~+]+\s*\{[\s\S]*?\}/.test(raw);
        if (looksLikeCss) {
          const existing: Element | null = document.querySelector(`style[data-injector-inline-css="${instanceId}"]`);
          if (existing) existing.remove();
          const styleTag: HTMLStyleElement = document.createElement('style');
          styleTag.setAttribute('data-injector-inline-css', instanceId);
          styleTag.textContent = raw;
          document.head.appendChild(styleTag);
        } else {
          executeScript(raw, 'inline-raw-js');
        }
        return;
      }

      // DOMParser moves <script> tags to <head>, so query the whole document
      const parser: DOMParser = new DOMParser();
      const doc: Document     = parser.parseFromString(raw, 'text/html');
      const allScriptEls: HTMLScriptElement[] = [].slice.call(doc.querySelectorAll('script:not([src])'));
      const scriptCodes: string[] = [];
      for (let i: number = 0; i < allScriptEls.length; i++) {
        const code: string = allScriptEls[i].textContent || '';
        if (code.trim() !== '') scriptCodes.push(code);
      }

      const safeHtml: string = stripInlineHandlers(raw);
      htmlContainerRef.current.innerHTML = safeHtml;
      reattachEvents(htmlContainerRef.current);
      startMutationObserver(htmlContainerRef.current);

      setTimeout((): void => {
        for (let i: number = 0; i < scriptCodes.length; i++) {
          if (scriptCodes[i].trim() !== '') {
            executeScript(scriptCodes[i], 'inline-html-script');
          }
        }
      }, 300);
    }
  }, [htmlInput]);

  // ── Edit mode — blue card ──────────────────────────────────────────────
  if (isEditMode) {
    return (
      <section className={`${styles.scriptinjector} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.editMode}>
          <div style={{ fontWeight: 600, marginBottom: '4px' }}>
            ⚙️ ScriptInjector Web Part
          </div>
          {nothingProvided
            ? <span>⚠️ No HTML provided. Open property pane to configure.</span>
            : <>
                {!isHtmlInlineEmpty && <div>✅ Inline HTML — active</div>}
                {!isHtmlFileEmpty   && <div>🌐 HTML URL: <strong>{htmlFileName.trim()}</strong></div>}
              </>
          }
        </div>
      </section>
    );
  }

  // ── Published mode ─────────────────────────────────────────────────────
  return (
    <div ref={htmlContainerRef} />
  );
};

export default Scriptinjector;