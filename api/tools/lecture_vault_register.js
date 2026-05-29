// モジュールの読み込み
const path = require('path');
const fs = require('fs');
const puppeteer = require('puppeteer');

const {
  GoogleSpreadsheet
} = require('google-spreadsheet');

console.log = (...args) => console.error(...args);

async function readStdinJson() {
  const chunks = [];
  for await (const chunk of process.stdin) chunks.push(chunk);
  const raw = Buffer.concat(chunks).toString('utf8').trim();
  if (!raw) throw new Error('Vault登録入力JSONが空です。');
  return JSON.parse(raw);
}

function loadGoogleCredentials() {
  if (process.env.GSA_JSON) return JSON.parse(process.env.GSA_JSON);

  const candidates = [
    process.env.GOOGLE_SA_PATH,
    path.join(__dirname, '..', '_master', 'client_secret.json'),
    path.join(__dirname, 'client_secret.json'),
  ].filter(Boolean);

  const credentialPath = candidates.find(candidate => fs.existsSync(candidate));
  if (!credentialPath) {
    throw new Error('GSA_JSON / GOOGLE_SA_PATH / client_secret.json が見つかりません。');
  }
  return JSON.parse(fs.readFileSync(credentialPath, 'utf8'));
}

async function withTimeout(promise, timeoutMs, label) {
  let timer;
  try {
    return await Promise.race([
      promise,
      new Promise((_, reject) => {
        timer = setTimeout(() => reject(new Error(`${label} timed out after ${timeoutMs}ms`)), timeoutMs);
      }),
    ]);
  } finally {
    clearTimeout(timer);
  }
}



(async function () {

  var red = '\u001b[31m';
  var reset = '\u001b[0m';
  const slowMo = Number(process.env.PUPPETEER_SLOW_MO || 0) || 0;

  const input = await readStdinJson();
  const rawPackages = Array.isArray(input.packages) ? input.packages : [input];
  const LOGIN_USER = String(input.vaultAccount || rawPackages[0]?.vaultAccount || '').trim();
  const packages = rawPackages.map((item, index) => {
    const packageInput = {
      presentationName: String(item.presentationName || '').trim(),
      presentationId: String(item.presentationId || '').trim(),
      product: String(item.product || '').trim(),
      absoluteZipPath: String(item.absoluteZipPath || '').trim(),
      mediaFileName: String(item.mediaFileName || item.presentationName || `package_${index + 1}`).trim(),
    };
    if (!packageInput.presentationName || !packageInput.presentationId || !packageInput.product || !packageInput.absoluteZipPath || !LOGIN_USER) {
      throw new Error(`Vault登録入力が不足しています: ${packageInput.mediaFileName}`);
    }
    if (!fs.existsSync(packageInput.absoluteZipPath)) {
      throw new Error(`ZIPファイルが見つかりません: ${packageInput.absoluteZipPath}`);
    }
    return packageInput;
  });

  console.log(`Vault登録ランナーを開始します。対象 ${packages.length} 件 / アカウント: ${LOGIN_USER}`);

  const creds2 = loadGoogleCredentials();
  console.log("Google認証情報を読み込みました。");

  const doc2 = new GoogleSpreadsheet('1zM2cBmSXBc_kpShPZtf9N9FfJq2Ohc8TLUnwCOWDZSc');



  doc2.useServiceAccountAuth(creds2);







  // アカウント管理表読み込み
  console.log("Vaultアカウント管理表を読み込んでいます。");
  await doc2.loadInfo();


  let logPass

  if (LOGIN_USER !== "Hayato.Seto@vv-agency.com") {

    const sheet2 = doc2.sheetsByTitle["MSD"];
    const rows2 = await sheet2.getRows();

    let shiftRow = rows2.find(row =>
      row.アカウント名 === LOGIN_USER && row.環境 === '本番/UAT' && row.サービス === 'Vault(iDetail)');

    logPass = shiftRow.パスワード;

  } else {

    const sheet2 = doc2.sheetsByTitle["嵐丸"];
    const rows2 = await sheet2.getRows();

    let shiftRow = rows2.find(row =>
      row.アカウント名 === LOGIN_USER && row.メーカー === 'Vault');

    logPass = shiftRow.パスワード;

  }

  const LOGIN_PASS = logPass;

  if (!LOGIN_PASS) {
    throw new Error(`Vaultアカウントのパスワードが見つかりません: ${LOGIN_USER}`);
  }
  console.log("Vaultパスワードを取得しました。");






  const options = [
    '--disable-gpu',
    '--disable-dev-shm-usage',
    '--disable-setuid-sandbox',
    '--no-first-run',
    '--no-sandbox',
    '--no-zygote',
  ];

  console.log("Chromeを起動しています。");
  const headlessValue = String(process.env.LECTURE_TOOL_HEADLESS || process.env.PUPPETEER_HEADLESS || '').trim().toLowerCase();
  const headless = ['1', 'true', 'yes', 'new'].includes(headlessValue) ? true : false;
  const executablePath = String(process.env.LECTURE_TOOL_CHROME_EXECUTABLE_PATH || process.env.PUPPETEER_EXECUTABLE_PATH || '').trim();
  const launchOptions = {
    headless,
    ignoreDefaultArgs: ['--disable-extensions'],
    args: options,
    slowMo,
  };
  if (executablePath) {
    launchOptions.executablePath = executablePath;
    console.log(`Chromium実行パスを使用します: ${executablePath}`);
  } else {
    launchOptions.channel = process.env.PUPPETEER_CHANNEL || 'chrome';
  }
  const browser = await puppeteer.launch(launchOptions)

  const results = [];
  try {
    for (let i = 0; i < packages.length; i += 1) {
      const packageInput = packages[i];
      console.log(`Vault登録 ${i + 1}/${packages.length} を開始します: ${packageInput.mediaFileName}`);
      const common = {
        mediaFileName: packageInput.mediaFileName,
        presentationName: packageInput.presentationName,
        presentationId: packageInput.presentationId,
        vaultAccount: LOGIN_USER,
      };
      let vaultResult;
      try {
        vaultResult = await createVAULT(
          packageInput.presentationName,
          packageInput.presentationId,
          packageInput.product,
          packageInput.absoluteZipPath
        )
      } catch (err) {
        vaultResult = ['作成失敗', '', '', err?.stack || err?.toString?.() || String(err)];
      }
      if (vaultResult?.[3]) {
        const resultItem = {
          ...common,
          status: vaultResult?.[0] || '作成失敗',
          url: vaultResult?.[1] || '',
          slideUrl: vaultResult?.[2] || '',
          error: vaultResult?.[3] || '',
        };
        results.push(resultItem);
        process.stderr.write(`__VAULT_RESULT__${JSON.stringify(resultItem)}\n`);
        console.log(`Vault登録をスキップして次へ進みます: ${packageInput.mediaFileName} / ${vaultResult?.[3]}`);
      } else {
        const resultItem = {
          ...common,
          status: vaultResult?.[0] || '作成完了',
          url: vaultResult?.[1] || '',
          slideUrl: vaultResult?.[2] || '',
          error: '',
        };
        results.push(resultItem);
        process.stderr.write(`__VAULT_RESULT__${JSON.stringify(resultItem)}\n`);
        console.log(`Vault登録 ${i + 1}/${packages.length} が完了しました: ${packageInput.mediaFileName}`);
      }
    }
  } finally {
    console.log("Chromeを終了しています。");
    try {
      await withTimeout(browser.close(), 15000, "Chrome終了");
      console.log("Chrome終了が完了しました。");
    } catch (e) {
      console.log(`Chrome終了でタイムアウトしました。処理結果を返します: ${e.message}`);
    }
  }

  await new Promise(resolve => process.stdout.write(JSON.stringify({
    results,
    vaultAccount: LOGIN_USER,
  }, null, 2), resolve));
  process.exit(0);
  return;


  async function createVAULT(presentationName, presentationId, product, zIPfolder) {

    let status = "作成失敗";
    let slideURL = "";

    const LOGIN_USER_SELECTOR = '#j_username';
    const LOGIN_PASS_SELECTOR = 'input[type=password]';
    const LOGIN_CONTINUE_SELECTOR = 'button[name=continue]';
    const LOGIN_SUBMIT_SELECTOR = 'button[name=login]';

    const CREATE_SUBMIT_SELECTOR = 'button[tooltip=Create]';
    const PLACEHOLDER_SUBMIT_SELECTOR = 'li[data-value=Placeholder]';
    const BINOCULARS_SUBMIT_SELECTOR = '.binoculars';
    const TPYE_SUBMIT_SELECTOR = '#uploadTypeSelect';

    const OK_SUBMIT_SELECTOR = '.ok';
    const NEXT_SUBMIT_SELECTOR = '.nextStep';

    const COUNTRY_SELECTOR = 'div[name=country_b] > .vv_pill_container > input';
    const COUNTRY = 'Japan';

    const LANGUAGE_SELECTOR = 'div[name=language_b] > .vv_pill_container > input';
    const LANGUAGE = 'Japanese';


    const NAME_SELECTOR = 'textarea[name=name]';
    const NAME = presentationName;

    const PRODUCT_SELECTOR = 'div[name=crmProduct_b] > .vv_pill_container > input';
    const PRODUCT = product;

    const PRODUCT_SELECTOR_TEST = 'div[name=product] > .vv_pill_container > input';
    const PRODUCT_TEST = "Cholecap";

    const PRODUCTTEXT_SELECTOR = 'textarea[name=productText]';

    const presentationId_SELECTOR = 'textarea[name=crmPresentationId_b]';


    const ACCESSCONTROL_SELECTOR = '#delegateAccessControl input';
    const ACCESSCONTROL = 'arashimaru@msd.com';


    const DETAILGROUP_SELECTOR = 'div[name=crmDetailGroup_b] > .vv_pill_container > input';
    const DETAILGROUP = 'JP Detail Group';


    const DESCRIPTION_SELECTOR = 'textarea[name=title]';
    const DESCRIPTION = "";

    const EVENTID_SELECTOR = 'textarea[name=eventID]';
    const EVENTID = "";

    const FRAGMENTCATEGORY_SELECTOR = 'textarea[name=fragmentCategory]';
    const FRAGMENTCATEGORY = 'VMMTG';

    // const DATE_SELECTOR = 'input[datatype=Date]';
    // date = new Date(date);
    // date = date.toFormat('M/D/YYYY');
    // const DATE = date;

    const SAVE_SUBMIT_SELECTOR = '.save';

    const UPLOAD_SUBMIT_SELECTOR = '.addContentButton';
    const UPLOAD_INPUT_SELECTOR = '.file_upload_control';

    const ADDASSET_SUBMIT_SELECTOR = '.add-assets';
    const ADDASSET_INPUT_SELECTOR = '.fileUploadControl';

    const ADDASSET_UPLOAD_SELECTOR = '.upload';
    const ASSET_FILE_SIZE = '.vv_mime_zip';

    const EDIT_SUBMIT_SELECTOR = '.docInfoEditButton';

    const STATUS_SUBMIT_SELECTOR = '.vv-action-bar-dropdown-menu > button';
    const STAGED_SUBMIT_SELECTOR = 'li[data-value=dynamicAction\\:stage]';


    const MENU_CONTAINER_SELECTOR = 'ul.ui-autocomplete.ui-menu, ul.vv_vof_lookup_panel.ui-menu, ul.ui-menu.ui-autocomplete';
    const MENU_ITEM_SELECTOR = 'li.ui-menu-item';
    const page = await browser.newPage();
    const DCL = { waitUntil: 'domcontentloaded' };
    const SHORT_WAIT = 3000;
    const NORMAL_WAIT = 30000;
    const LONG_WAIT = 90000;

    async function waitForOptionalSelector(selector, timeout = SHORT_WAIT, options = {}) {
      try {
        return await page.waitForSelector(selector, { timeout, ...options });
      } catch (err) {
        return null;
      }
    }

    async function waitForAnySelector(selectors, timeout = NORMAL_WAIT, options = {}) {
      const result = await Promise.race(
        selectors.map(selector =>
          page.waitForSelector(selector, { timeout, ...options }).then(handle => ({ selector, handle }))
        )
      );
      return result;
    }

    async function waitForElementCount(selector, minCount, timeout = NORMAL_WAIT) {
      await page.waitForFunction(
        (targetSelector, targetCount) => document.querySelectorAll(targetSelector).length >= targetCount,
        { timeout },
        selector,
        minCount
      );
    }

    async function waitForDomMutation(selector, timeout = NORMAL_WAIT) {
      await page.evaluate((targetSelector, maxWait) => new Promise(resolve => {
        const target = document.querySelector(targetSelector) || document.body;
        let done = false;
        const observer = new MutationObserver(() => finish());
        const timer = setTimeout(() => finish(), maxWait);

        function finish() {
          if (done) return;
          done = true;
          clearTimeout(timer);
          observer.disconnect();
          resolve();
        }

        observer.observe(target, { childList: true, subtree: true, characterData: true });
      }), selector, timeout);
    }

    async function waitForClickable(selector, timeout = NORMAL_WAIT) {
      await page.waitForFunction(
        targetSelector => {
          return Array.from(document.querySelectorAll(targetSelector)).some(element => {
            const style = window.getComputedStyle(element);
            const rect = element.getBoundingClientRect();
            return rect.width > 0 &&
              rect.height > 0 &&
              !element.disabled &&
              element.getAttribute('aria-disabled') !== 'true' &&
              !element.classList.contains('disabled') &&
              style.visibility !== 'hidden' &&
              style.display !== 'none' &&
              style.pointerEvents !== 'none';
          });
        },
        { timeout },
        selector
      );
    }

    async function clickWhenReady(selector, timeout = NORMAL_WAIT) {
      await waitForClickable(selector, timeout);
      const targetIndex = await page.evaluate(targetSelector => {
        const elements = Array.from(document.querySelectorAll(targetSelector));
        return elements.findIndex(candidate => {
          const style = window.getComputedStyle(candidate);
          const rect = candidate.getBoundingClientRect();
          return rect.width > 0 &&
            rect.height > 0 &&
            !candidate.disabled &&
            candidate.getAttribute('aria-disabled') !== 'true' &&
            !candidate.classList.contains('disabled') &&
            style.visibility !== 'hidden' &&
            style.display !== 'none' &&
            style.pointerEvents !== 'none';
        });
      }, selector);
      if (targetIndex < 0) {
        throw new Error(`${selector} をクリックできませんでした。`);
      }
      const handles = await page.$$(selector);
      await handles[targetIndex].click({ delay: 50 });
    }

    async function openBinderAddFilesMenu() {
      const addButtonSelector = "button.addToBinder, .addToBinder";
      const addFilesSelector = ".addFiles a, li.addFiles a, a[val='upload'], .vv_menu_item a[val='upload']";

      for (let attempt = 1; attempt <= 3; attempt += 1) {
        console.log(`BinderのAddメニューを開いています。(${attempt}/3)`);
        await waitForOptionalSelector(addButtonSelector, LONG_WAIT, { visible: true });
        const clicked = await page.evaluate(selector => {
          const candidates = Array.from(document.querySelectorAll(selector));
          const element = candidates.find(candidate => {
            const style = window.getComputedStyle(candidate);
            const rect = candidate.getBoundingClientRect();
            return rect.width > 0 &&
              rect.height > 0 &&
              !candidate.disabled &&
              candidate.getAttribute('aria-disabled') !== 'true' &&
              !candidate.classList.contains('disabled') &&
              style.visibility !== 'hidden' &&
              style.display !== 'none' &&
              style.pointerEvents !== 'none';
          });
          if (!element) return false;
          element.scrollIntoView({ block: 'center', inline: 'center' });
          ['mouseover', 'mouseenter', 'mousemove', 'mousedown', 'mouseup', 'click'].forEach(type => {
            element.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, view: window }));
          });
          return true;
        }, addButtonSelector);
        if (!clicked) {
          await clickWhenReady(addButtonSelector, NORMAL_WAIT);
        }

        const clickedUpload = await page.waitForFunction(selector => {
          const visible = element => {
            const style = window.getComputedStyle(element);
            const rect = element.getBoundingClientRect();
            return rect.width > 0 &&
              rect.height > 0 &&
              style.visibility !== 'hidden' &&
              style.display !== 'none' &&
              style.pointerEvents !== 'none';
          };
          return Array.from(document.querySelectorAll(selector)).some(element => {
            const text = (element.textContent || '').replace(/\s+/g, ' ').trim();
            return visible(element) && (
              element.getAttribute('val') === 'upload' ||
              /Upload File/i.test(text)
            );
          });
        }, { timeout: 10000 }, addFilesSelector).then(async () => {
          return page.evaluate(selector => {
            const visible = element => {
              const style = window.getComputedStyle(element);
              const rect = element.getBoundingClientRect();
              return rect.width > 0 &&
                rect.height > 0 &&
                style.visibility !== 'hidden' &&
                style.display !== 'none' &&
                style.pointerEvents !== 'none';
            };
            const element = Array.from(document.querySelectorAll(selector)).find(candidate => {
              const text = (candidate.textContent || '').replace(/\s+/g, ' ').trim();
              return visible(candidate) && (
                candidate.getAttribute('val') === 'upload' ||
                /Upload File/i.test(text)
              );
            });
            if (!element) return false;
            element.scrollIntoView({ block: 'center', inline: 'center' });
            ['mouseover', 'mouseenter', 'mousemove', 'mousedown', 'mouseup', 'click'].forEach(type => {
              element.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, view: window }));
            });
            return true;
          }, addFilesSelector);
        }).catch(() => false);

        if (clickedUpload) {
          return;
        }

        console.log("Add Filesメニューが表示されなかったため再試行します。");
        await waitForDomMutation("body", 1500);
      }

      throw new Error("BinderのAddメニューからAdd Filesを開けませんでした。");
    }

    async function clickActiveSaveButton(selector = SAVE_SUBMIT_SELECTOR, label = 'Save', timeout = LONG_WAIT, options = {}) {
      const waitForGlobalBusy = !!options.waitForGlobalBusy;
      let targetIndex = -1;
      let targetInfo = null;
      try {
        const targetResult = await page.waitForFunction(
          (targetSelector, shouldWaitForGlobalBusy) => {
            const visible = element => {
              const style = window.getComputedStyle(element);
              const rect = element.getBoundingClientRect();
              return rect.width > 0 &&
                rect.height > 0 &&
                style.visibility !== 'hidden' &&
                style.display !== 'none' &&
                style.pointerEvents !== 'none';
            };
            const disabled = element => {
              const classText = String(element.className || '');
              const parentDisabled = element.closest('.disabled, .vv_disabled, .vv-button-disabled, .vv_button_disabled, .ui-state-disabled, [aria-disabled="true"]');
              return !!element.disabled ||
                element.hasAttribute('disabled') ||
                element.getAttribute('aria-disabled') === 'true' ||
                /(^|\s)(disabled|vv_disabled|vv-button-disabled|vv_button_disabled|ui-state-disabled|inactive|loading|processing)(\s|$)/i.test(classText) ||
                !!parentDisabled;
            };
            if (shouldWaitForGlobalBusy) {
              const busySelectors = [
                '[role="progressbar"]',
                '.file_upload_progress',
                '.fileUploadProgress',
                '.uploadProgress',
                '.upload_progress',
                '.fileUploadSpinner',
                '.library_loading',
                '.loadingIndicator',
                '.vv_spinner',
                '.spinner',
                '[class*="uploading"]',
                '[class*="Uploading"]',
                '[class*="uploadStatus"]',
                '[class*="UploadStatus"]',
                '[class*="progressBar"]',
                '[class*="ProgressBar"]',
              ];
              const uploadBusy = busySelectors.some(selector => {
                return Array.from(document.querySelectorAll(selector)).some(element => {
                  if (!visible(element)) return false;
                  const text = [
                    element.textContent || '',
                    element.getAttribute('title') || '',
                    element.getAttribute('aria-label') || '',
                    element.className || '',
                  ].join(' ');
                  return /upload|uploading|progress|loading|spinner|processing|generating|rendition|アップロード|処理|読み込み/i.test(text);
                });
              });
              if (uploadBusy) return false;
            }

            const elements = Array.from(document.querySelectorAll(targetSelector));
            const index = elements.findIndex(element => visible(element) && !disabled(element));
            if (index < 0) return false;
            const element = elements[index];
            const rect = element.getBoundingClientRect();
            return {
              index,
              text: (element.textContent || '').trim(),
              title: element.getAttribute('title') || '',
              ariaLabel: element.getAttribute('aria-label') || '',
              className: String(element.className || ''),
              rect: {
                x: Math.round(rect.left),
                y: Math.round(rect.top),
                width: Math.round(rect.width),
                height: Math.round(rect.height),
              },
            };
          },
          { timeout },
          selector,
          waitForGlobalBusy
        ).then(handle => handle.jsonValue());
        targetIndex = targetResult.index;
        targetInfo = targetResult;
      } catch (e) {
        const state = await page.evaluate(targetSelector => {
          const visible = element => {
            const style = window.getComputedStyle(element);
            const rect = element.getBoundingClientRect();
            return rect.width > 0 &&
              rect.height > 0 &&
              style.visibility !== 'hidden' &&
              style.display !== 'none';
          };
          const elements = Array.from(document.querySelectorAll(targetSelector));
          return {
            url: location.href,
            selector: targetSelector,
            candidates: elements.map((element, index) => ({
              index,
              text: (element.textContent || '').trim(),
              className: String(element.className || ''),
              disabled: !!element.disabled || element.hasAttribute('disabled'),
              ariaDisabled: element.getAttribute('aria-disabled') || '',
              visible: visible(element),
            })).slice(0, 10),
            activeTag: document.activeElement?.tagName || '',
            activeText: (document.activeElement?.textContent || '').trim().slice(0, 80),
          };
        }, selector).catch(err => ({ error: String(err) }));
        throw new Error(`${label}ボタンがアクティブになりませんでした。現在の状態: ${JSON.stringify(state)}`);
      }

      if (targetIndex < 0) {
        throw new Error(`${label}ボタンがアクティブになりませんでした。`);
      }

      const handles = await page.$$(selector);
      const handle = handles[targetIndex];
      if (!handle) {
        throw new Error(`${label}ボタンを取得できませんでした。index=${targetIndex}`);
      }
      console.log(`${label}ボタンがアクティブになりました。クリックします。対象: ${JSON.stringify(targetInfo)}`);
      await handle.evaluate(element => {
        element.scrollIntoView({ block: 'center', inline: 'center' });
      });
      await handle.click({ delay: 50 });

      if (options.confirmAfterClick) {
        const beforeUrl = options.beforeUrl || "";
        try {
          await page.waitForFunction(
            previousUrl => {
              const visible = element => {
                if (!element) return false;
                const style = window.getComputedStyle(element);
                const rect = element.getBoundingClientRect();
                return rect.width > 0 &&
                  rect.height > 0 &&
                  style.visibility !== 'hidden' &&
                  style.display !== 'none';
              };
              const urlChanged = previousUrl && location.href !== previousUrl;
              const docInfoVisible = !!document.querySelector('.vv_docstatus_wrapper, li[data-target-key=doc_info_relationships__sys]');
              const renditionVisible = Array.from(document.querySelectorAll('.generatingRenditionSpinner, .generatingRenditionLabel'))
                .some(visible);
              const saveDialogGone = !Array.from(document.querySelectorAll('.save')).some(element => visible(element));
              return urlChanged || docInfoVisible || renditionVisible || saveDialogGone;
            },
            { timeout: 45000 },
            beforeUrl
          );
          console.log(`${label}クリック後の保存開始を確認しました。`);
        } catch (e) {
          const state = await page.evaluate(() => ({
            url: location.href,
            saveButtons: Array.from(document.querySelectorAll('.save')).map((element, index) => {
              const style = window.getComputedStyle(element);
              const rect = element.getBoundingClientRect();
              return {
                index,
                text: (element.textContent || '').trim(),
                title: element.getAttribute('title') || '',
                className: String(element.className || ''),
                disabled: !!element.disabled || element.hasAttribute('disabled'),
                ariaDisabled: element.getAttribute('aria-disabled') || '',
                visible: rect.width > 0 && rect.height > 0 && style.visibility !== 'hidden' && style.display !== 'none',
              };
            }).slice(0, 10),
            bodyText: (document.body?.textContent || '').replace(/\s+/g, ' ').trim().slice(0, 600),
          })).catch(err => ({ error: String(err) }));
          throw new Error(`${label}クリック後に保存開始を確認できませんでした。現在の状態: ${JSON.stringify(state)}`);
        }
      }
    }

    async function clearAndType(selector, value, options = {}) {
      const textValue = String(value ?? '');
      const { sectionTitle = '', ...typeOptions } = options;
      const targetIndex = await getVisibleElementIndex(selector, NORMAL_WAIT, { sectionTitle });
      const handles = await page.$$(selector);
      const handle = handles[targetIndex];
      if (!handle) {
        throw new Error(`${selector} の入力対象を取得できませんでした。`);
      }
      let typedByKeyboard = false;

      try {
        await handle.evaluate(el => {
          el.scrollIntoView({ block: 'center', inline: 'center' });
          el.focus();
        });
        await handle.click({ clickCount: 3, delay: 30 });
        const modifier = process.platform === 'darwin' ? 'Meta' : 'Control';
        await page.keyboard.down(modifier);
        await page.keyboard.press('A');
        await page.keyboard.up(modifier);
        await page.keyboard.press('Backspace');
        if (textValue) {
          await page.keyboard.type(textValue, { delay: 30, ...typeOptions });
        }
        typedByKeyboard = true;
      } catch (e) {
        console.log(`${selector} は直接クリックできないため、値を直接設定します。`);
      }

      let actualValue = await handle.evaluate(el => el.value);
      if (typedByKeyboard && actualValue === textValue) return;

      await handle.evaluate((el, expected) => {
        el.scrollIntoView({ block: 'center', inline: 'center' });
        el.focus();
        const prototype = el.tagName === 'TEXTAREA'
          ? HTMLTextAreaElement.prototype
          : HTMLInputElement.prototype;
        const valueSetter = Object.getOwnPropertyDescriptor(prototype, 'value')?.set;
        if (valueSetter) {
          valueSetter.call(el, expected);
        } else {
          el.value = expected;
        }
        el.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText', data: expected }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
      }, textValue);

      await page.waitForFunction(
        (targetSelector, index, expected) => Array.from(document.querySelectorAll(targetSelector))[index]?.value === expected,
        { timeout: SHORT_WAIT },
        selector,
        targetIndex,
        textValue
      );
    }

    async function clearAndTypeRequired(selector, value, label = selector, options = {}) {
      const textValue = String(value ?? '');
      try {
        await clearAndType(selector, textValue, options);
        return;
      } catch (e) {
        const setState = await page.evaluate((targetSelector, expected) => {
          const elements = Array.from(document.querySelectorAll(targetSelector));
          const target = elements[0];
          if (!target) {
            return {
              ok: false,
              count: 0,
              reason: 'element not found',
            };
          }
          target.scrollIntoView?.({ block: 'center', inline: 'center' });
          target.focus?.();
          const prototype = target.tagName === 'TEXTAREA'
            ? HTMLTextAreaElement.prototype
            : HTMLInputElement.prototype;
          const valueSetter = Object.getOwnPropertyDescriptor(prototype, 'value')?.set;
          if (valueSetter) {
            valueSetter.call(target, expected);
          } else {
            target.value = expected;
          }
          target.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText', data: expected }));
          target.dispatchEvent(new Event('change', { bubbles: true }));
          const style = window.getComputedStyle(target);
          const rect = target.getBoundingClientRect();
          return {
            ok: target.value === expected,
            count: elements.length,
            value: target.value || '',
            visible: rect.width > 0 && rect.height > 0 && style.visibility !== 'hidden' && style.display !== 'none',
            className: String(target.className || ''),
          };
        }, selector, textValue).catch(err => ({ ok: false, error: String(err) }));
        console.log(`${label} 表示フィールド入力に失敗したため直接設定しました: ${e.message} / ${JSON.stringify(setState)}`);
        if (!setState.ok) {
          throw new Error(`${label} の入力に失敗しました: ${e.message} / ${JSON.stringify(setState)}`);
        }
        await page.waitForFunction(
          (targetSelector, expected) => Array.from(document.querySelectorAll(targetSelector)).some(element => element.value === expected),
          { timeout: SHORT_WAIT },
          selector,
          textValue
        );
      }
    }

    async function selectRadioById(radioId) {
      const selector = `#${radioId}`;
      await page.waitForSelector(selector, { timeout: NORMAL_WAIT });
      await page.$eval(selector, el => {
        if (!el.checked) el.click();
        el.checked = true;
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
      });
      await page.waitForFunction(
        targetSelector => document.querySelector(targetSelector)?.checked === true,
        { timeout: SHORT_WAIT },
        selector
      );
    }

    async function getVisibleElementIndex(selector, timeout = NORMAL_WAIT, options = {}) {
      await waitForClickable(selector, timeout);
      const index = await page.evaluate((targetSelector, preferredSectionTitle) => {
        const isVisible = element => {
          if (!element) return false;
          const style = window.getComputedStyle(element);
          const rect = element.getBoundingClientRect();
          return rect.width > 0 &&
            rect.height > 0 &&
            !element.disabled &&
            element.getAttribute('aria-disabled') !== 'true' &&
            !element.classList.contains('disabled') &&
            style.visibility !== 'hidden' &&
            style.display !== 'none' &&
            style.pointerEvents !== 'none';
        };
        const centerY = window.innerHeight / 2;
        const centerX = window.innerWidth / 2;
        const normalize = text => String(text || '').trim().replace(/\s+/g, ' ').toUpperCase();
        const expectedSection = normalize(preferredSectionTitle);
        const section = expectedSection
          ? Array.from(document.querySelectorAll('h1, h2, h3, h4, [title]'))
            .filter(isVisible)
            .find(element => normalize(element.getAttribute('title') || element.textContent) === expectedSection)
          : null;
        const sectionRect = section ? section.getBoundingClientRect() : null;
        const candidates = Array.from(document.querySelectorAll(targetSelector));
        const visible = candidates
          .map((element, index) => {
            const rect = element.getBoundingClientRect();
            const ok = isVisible(element);
            const centerDistance = Math.abs((rect.top + rect.height / 2) - centerY) +
              Math.abs((rect.left + rect.width / 2) - centerX);
            const sectionDistance = sectionRect
              ? (rect.top >= sectionRect.top - 80 && rect.top <= sectionRect.bottom + 900
                ? Math.abs(rect.top - sectionRect.bottom) + Math.abs(rect.left - sectionRect.left)
                : 100000 + Math.abs(rect.top - sectionRect.bottom) + Math.abs(rect.left - sectionRect.left))
              : centerDistance;
            const distance = sectionRect ? sectionDistance : centerDistance;
            return { index, ok, distance };
          })
          .filter(item => item.ok)
          .sort((a, b) => a.distance - b.distance);
        return visible[0]?.index ?? -1;
      }, selector, options.sectionTitle || '');
      if (index < 0) {
        throw new Error(`${selector} の表示中フィールドが見つかりませんでした。`);
      }
      return index;
    }

    async function getVisibleMenuItems(inputSelector = null) {
      return page.evaluate((menuContainerSelector, menuItemSelector, fieldSelector) => {
        const visibleCenterInput = selector => {
          const centerY = window.innerHeight / 2;
          const centerX = window.innerWidth / 2;
          return Array.from(document.querySelectorAll(selector))
            .map(element => {
              const style = window.getComputedStyle(element);
              const rect = element.getBoundingClientRect();
              const visible = rect.width > 0 &&
                rect.height > 0 &&
                !element.disabled &&
                element.getAttribute('aria-disabled') !== 'true' &&
                style.visibility !== 'hidden' &&
                style.display !== 'none';
              const distance = Math.abs((rect.top + rect.height / 2) - centerY) +
                Math.abs((rect.left + rect.width / 2) - centerX);
              return { element, visible, distance };
            })
            .filter(item => item.visible)
            .sort((a, b) => a.distance - b.distance)[0]?.element || null;
        };
        const marker = document.querySelector('[data-lecture-active-lookup="true"]');
        const active = document.activeElement;
        const input = fieldSelector
          ? (marker?.matches?.(fieldSelector) ? marker : active?.matches?.(fieldSelector) ? active : visibleCenterInput(fieldSelector))
          : null;
        const inputRect = input ? input.getBoundingClientRect() : null;
        const visibleMenus = Array.from(document.querySelectorAll(menuContainerSelector))
          .map((menu, menuIndex) => {
            const style = window.getComputedStyle(menu);
            const rect = menu.getBoundingClientRect();
            const visible = rect.width > 0 &&
              rect.height > 0 &&
              style.visibility !== 'hidden' &&
              style.display !== 'none';
            const nearInput = !inputRect || (
              rect.top >= inputRect.top - 40 &&
              rect.top <= inputRect.bottom + 500 &&
              rect.right >= inputRect.left - 80 &&
              rect.left <= inputRect.right + 500
            );
            const distance = inputRect
              ? Math.abs(rect.top - inputRect.bottom) + Math.abs(rect.left - inputRect.left)
              : 0;
            return {
              menu,
              menuIndex,
              visible,
              nearInput,
              distance,
              zIndex: Number.parseInt(style.zIndex, 10) || 0,
            };
          })
          .filter(menu => menu.visible);
        return visibleMenus
          .sort((a, b) => Number(b.nearInput) - Number(a.nearInput) ||
            (b.zIndex - a.zIndex) ||
            (a.distance - b.distance))
          .flatMap(menuInfo => Array.from(menuInfo.menu.querySelectorAll(menuItemSelector))
            .map((element, itemIndex) => {
              const style = window.getComputedStyle(element);
              const rect = element.getBoundingClientRect();
              const visible = rect.width > 0 &&
                rect.height > 0 &&
                style.visibility !== 'hidden' &&
                style.display !== 'none';
              return {
                menuIndex: menuInfo.menuIndex,
                itemIndex,
                text: (element.textContent || '').trim(),
                visible,
                nearInput: menuInfo.nearInput,
                distance: menuInfo.distance,
                zIndex: menuInfo.zIndex,
              };
            })
            .filter(item => item.visible && item.text));
      }, MENU_CONTAINER_SELECTOR, MENU_ITEM_SELECTOR, inputSelector);
    }

    async function closeVisibleMenus() {
      await page.keyboard.press('Escape').catch(() => null);
      await page.evaluate(() => document.activeElement?.blur()).catch(() => null);
      await waitForOptionalSelector('.ui-menu-item', SHORT_WAIT, { hidden: true });
    }

    function normalizeMenuText(text) {
      return String(text || '')
        .trim()
        .replace(/\s+/g, ' ')
        .toUpperCase();
    }

    function menuTextMatchesTarget(text, expectedText) {
      const normalized = normalizeMenuText(text);
      return normalized === expectedText;
    }

    function sleep(ms) {
      return new Promise(resolve => setTimeout(resolve, ms));
    }

    async function waitForLookupCandidateReady(inputSelector, targetText, label, timeout = NORMAL_WAIT) {
      const expected = normalizeMenuText(targetText);
      const startedAt = Date.now();
      let lastState = null;

      while (Date.now() - startedAt < timeout) {
        lastState = await page.evaluate((menuContainerSelector, menuItemSelector, fieldSelector, expectedText) => {
          const normalize = text => String(text || '')
            .trim()
            .replace(/\s+/g, ' ')
            .toUpperCase();
          const visibleCenterInput = selector => {
            const centerY = window.innerHeight / 2;
            const centerX = window.innerWidth / 2;
            return Array.from(document.querySelectorAll(selector))
              .map(element => {
                const style = window.getComputedStyle(element);
                const rect = element.getBoundingClientRect();
                const visible = rect.width > 0 &&
                  rect.height > 0 &&
                  !element.disabled &&
                  element.getAttribute('aria-disabled') !== 'true' &&
                  style.visibility !== 'hidden' &&
                  style.display !== 'none';
                const distance = Math.abs((rect.top + rect.height / 2) - centerY) +
                  Math.abs((rect.left + rect.width / 2) - centerX);
                return { element, visible, distance };
              })
              .filter(item => item.visible)
              .sort((a, b) => a.distance - b.distance)[0]?.element || null;
          };
          const marker = document.querySelector('[data-lecture-active-lookup="true"]');
          const active = document.activeElement;
          const input = fieldSelector
            ? (marker?.matches?.(fieldSelector) ? marker : active?.matches?.(fieldSelector) ? active : visibleCenterInput(fieldSelector))
            : null;
          const inputRect = input ? input.getBoundingClientRect() : null;
          const visibleMenus = Array.from(document.querySelectorAll(menuContainerSelector))
            .map((menu, menuIndex) => {
              const style = window.getComputedStyle(menu);
              const rect = menu.getBoundingClientRect();
              const visible = rect.width > 0 &&
                rect.height > 0 &&
                style.visibility !== 'hidden' &&
                style.display !== 'none';
              const nearInput = !inputRect || (
                rect.top >= inputRect.top - 40 &&
                rect.top <= inputRect.bottom + 500 &&
                rect.right >= inputRect.left - 80 &&
                rect.left <= inputRect.right + 500
              );
              const distance = inputRect
                ? Math.abs(rect.top - inputRect.bottom) + Math.abs(rect.left - inputRect.left)
                : 0;
              return { menu, menuIndex, visible, nearInput, distance, zIndex: Number.parseInt(style.zIndex, 10) || 0 };
            })
            .filter(menu => menu.visible)
            .sort((a, b) => Number(b.nearInput) - Number(a.nearInput) ||
              (a.distance - b.distance) ||
              (b.zIndex - a.zIndex) ||
              (a.menuIndex - b.menuIndex));
          const candidates = visibleMenus
            .flatMap(menuInfo => Array.from(menuInfo.menu.querySelectorAll(menuItemSelector)).map(item => {
              const style = window.getComputedStyle(item);
              const rect = item.getBoundingClientRect();
              const visible = rect.width > 0 &&
                rect.height > 0 &&
                style.visibility !== 'hidden' &&
                style.display !== 'none';
              const text = (item.textContent || '').trim();
              const $ = window.jQuery || window.$;
              const data = $ ? ($(item).data('ui-autocomplete-item') ||
                $(item).data('item.autocomplete') ||
                $(item.querySelector('.ui-menu-item-wrapper')).data('ui-autocomplete-item') ||
                null) : null;
              const dataText = data && typeof data === 'object'
                ? [data.value, data.label, data.name, data.text].filter(Boolean).join(' ')
                : String(data || '');
              return {
                text,
                visible,
                nearInput: menuInfo.nearInput,
                dataText,
                textMatches: normalize(text) === expectedText,
                dataMatches: normalize(dataText) === expectedText || normalize(dataText).split(' ').includes(expectedText),
                hasAutocompleteData: !!dataText,
              };
            }))
            .filter(item => item.visible && item.text);
          const exact = candidates.filter(item => item.textMatches && (!inputRect || item.nearInput));
          const hasAutocomplete = !!input && !!((window.jQuery || window.$)?.(input).data('ui-autocomplete') ||
            (window.jQuery || window.$)?.(input).data('autocomplete'));
          const exactReady = exact.some(item => !hasAutocomplete || item.dataMatches);
          return {
            ready: exact.length > 0 && exactReady,
            exactCount: exact.length,
            hasAutocomplete,
            candidateCount: candidates.length,
            firstExact: !!candidates[0]?.textMatches,
            firstCandidates: candidates.slice(0, 8).map(item => item.text),
            candidates: candidates.slice(0, 30),
          };
        }, MENU_CONTAINER_SELECTOR, MENU_ITEM_SELECTOR, inputSelector, expected);

        if (lastState?.ready) {
          await sleep(350);
          return lastState;
        }
        await sleep(150);
      }

      throw new Error(`${label}「${targetText}」の候補データ準備が完了しませんでした。候補状態: ${JSON.stringify(lastState)}`);
    }

    async function waitAndSelectMenuItem(target, options = {}) {
      const targetText = String(target);
      const expected = normalizeMenuText(targetText);
      const allowedExpectedTexts = [expected];
      const label = options.label || '候補';
      const inputSelector = options.inputSelector || null;
      const timeout = options.timeout || NORMAL_WAIT;
      const selectionMode = options.selectionMode || 'auto';

      try {
        await page.waitForFunction(
          (menuContainerSelector, menuItemSelector, fieldSelector, expectedTexts) => {
            const normalize = text => String(text || '')
              .trim()
              .replace(/\s+/g, ' ')
              .toUpperCase();
            const matchesMenuText = text => expectedTexts.some(expectedText => normalize(text) === expectedText);
            const visibleCenterInput = selector => {
              const centerY = window.innerHeight / 2;
              const centerX = window.innerWidth / 2;
              return Array.from(document.querySelectorAll(selector))
                .map(element => {
                  const style = window.getComputedStyle(element);
                  const rect = element.getBoundingClientRect();
                  const visible = rect.width > 0 &&
                    rect.height > 0 &&
                    !element.disabled &&
                    element.getAttribute('aria-disabled') !== 'true' &&
                    style.visibility !== 'hidden' &&
                    style.display !== 'none';
                  const distance = Math.abs((rect.top + rect.height / 2) - centerY) +
                    Math.abs((rect.left + rect.width / 2) - centerX);
                  return { element, visible, distance };
                })
                .filter(item => item.visible)
                .sort((a, b) => a.distance - b.distance)[0]?.element || null;
            };
            const marker = document.querySelector('[data-lecture-active-lookup="true"]');
            const active = document.activeElement;
            const input = fieldSelector
              ? (marker?.matches?.(fieldSelector) ? marker : active?.matches?.(fieldSelector) ? active : visibleCenterInput(fieldSelector))
              : null;
            const inputRect = input ? input.getBoundingClientRect() : null;
            const visibleMenus = Array.from(document.querySelectorAll(menuContainerSelector))
              .map(menu => {
                const style = window.getComputedStyle(menu);
                const rect = menu.getBoundingClientRect();
                const visible = rect.width > 0 &&
                  rect.height > 0 &&
                  style.visibility !== 'hidden' &&
                  style.display !== 'none';
                const nearInput = !inputRect || (
                  rect.top >= inputRect.top - 40 &&
                  rect.top <= inputRect.bottom + 500 &&
                  rect.right >= inputRect.left - 80 &&
                  rect.left <= inputRect.right + 500
                );
                return { menu, visible, nearInput };
              })
              .filter(menu => menu.visible);
            return visibleMenus.some(menuInfo => (!inputRect || menuInfo.nearInput) &&
              Array.from(menuInfo.menu.querySelectorAll(menuItemSelector)).some(element => {
                const style = window.getComputedStyle(element);
                const rect = element.getBoundingClientRect();
                return rect.width > 0 &&
                  rect.height > 0 &&
                  style.visibility !== 'hidden' &&
                  style.display !== 'none' &&
                  matchesMenuText(element.textContent || '');
              }));
          },
          { timeout },
          MENU_CONTAINER_SELECTOR,
          MENU_ITEM_SELECTOR,
          inputSelector,
          allowedExpectedTexts
        );
      } catch (e) {
        const visibleItems = await getVisibleMenuItems(inputSelector);
        throw new Error(`${label}「${targetText}」の一致候補が表示されませんでした。候補一覧: ${visibleItems.map(item => item.text).join(' / ') || 'なし'}`);
      }

      const visibleItems = await getVisibleMenuItems(inputSelector);
      console.log(`${label} 候補一覧: ${visibleItems.map(item => item.text).join(' / ')}`);
      const matches = visibleItems
        .map(item => {
          const matchedIndex = allowedExpectedTexts.findIndex(expectedText => menuTextMatchesTarget(item.text, expectedText));
          return { ...item, matchedIndex };
        })
        .filter(item => item.matchedIndex >= 0)
        .sort((a, b) => (a.matchedIndex - b.matchedIndex) || (a.distance - b.distance));

      if (matches.length === 0) {
        throw new Error(`${label}「${targetText}」の一致候補が0件です。候補一覧: ${visibleItems.map(item => item.text).join(' / ') || 'なし'}`);
      }
      if (matches.length > 1) {
        console.log(`${label}「${targetText}」の一致候補が${matches.length}件あります。入力欄に近い表示中の候補を選択します。`);
      }
      let readyState = null;
      if (options.waitForReady !== false) {
        readyState = await waitForLookupCandidateReady(inputSelector, targetText, label, options.readyTimeout || NORMAL_WAIT);
      } else {
        await sleep(options.stabilizeMs || 500);
      }
      if (readyState &&
        readyState.candidateCount > 20 &&
        !readyState.firstExact &&
        options.warnUnfiltered !== false) {
        console.log(`${label} 候補が絞り込まれていない可能性があります: 件数${readyState.candidateCount} / 先頭候補 ${readyState.firstCandidates.join(' / ')}`);
      }

      const selection = await page.evaluate((menuContainerSelector, menuItemSelector, fieldSelector, expectedTexts, mode) => {
        const normalize = text => String(text || '')
          .trim()
          .replace(/\s+/g, ' ')
          .toUpperCase();
        const matchedIndexFor = text => expectedTexts.findIndex(expectedText => normalize(text) === expectedText);
        const visibleCenterInput = selector => {
          const centerY = window.innerHeight / 2;
          const centerX = window.innerWidth / 2;
          return Array.from(document.querySelectorAll(selector))
            .map(element => {
              const style = window.getComputedStyle(element);
              const rect = element.getBoundingClientRect();
              const visible = rect.width > 0 &&
                rect.height > 0 &&
                !element.disabled &&
                element.getAttribute('aria-disabled') !== 'true' &&
                style.visibility !== 'hidden' &&
                style.display !== 'none';
              const distance = Math.abs((rect.top + rect.height / 2) - centerY) +
                Math.abs((rect.left + rect.width / 2) - centerX);
              return { element, visible, distance };
            })
            .filter(item => item.visible)
            .sort((a, b) => a.distance - b.distance)[0]?.element || null;
        };
        const marker = document.querySelector('[data-lecture-active-lookup="true"]');
        const active = document.activeElement;
        const input = fieldSelector
          ? (marker?.matches?.(fieldSelector) ? marker : active?.matches?.(fieldSelector) ? active : visibleCenterInput(fieldSelector))
          : null;
        const inputRect = input ? input.getBoundingClientRect() : null;
        const visibleMenus = Array.from(document.querySelectorAll(menuContainerSelector))
          .map((menu, menuIndex) => {
            const style = window.getComputedStyle(menu);
            const rect = menu.getBoundingClientRect();
            const visible = rect.width > 0 &&
              rect.height > 0 &&
              style.visibility !== 'hidden' &&
              style.display !== 'none';
            const nearInput = !inputRect || (
              rect.top >= inputRect.top - 40 &&
              rect.top <= inputRect.bottom + 500 &&
              rect.right >= inputRect.left - 80 &&
              rect.left <= inputRect.right + 500
            );
            const distance = inputRect
              ? Math.abs(rect.top - inputRect.bottom) + Math.abs(rect.left - inputRect.left)
              : 0;
            return {
              menu,
              menuIndex,
              visible,
              nearInput,
              distance,
              zIndex: Number.parseInt(style.zIndex, 10) || 0,
            };
          })
          .filter(menu => menu.visible);
        const sortedMenus = visibleMenus
          .sort((a, b) => Number(b.nearInput) - Number(a.nearInput) ||
            (a.distance - b.distance) ||
            (b.zIndex - a.zIndex) ||
            (a.menuIndex - b.menuIndex));
        const candidates = sortedMenus
          .flatMap((menuInfo, sortedMenuIndex) => Array.from(menuInfo.menu.querySelectorAll(menuItemSelector))
            .map((item, itemIndex) => {
              const style = window.getComputedStyle(item);
              const rect = item.getBoundingClientRect();
              const matchedIndex = matchedIndexFor(item.textContent || '');
              const active = item.classList.contains('ui-state-active') ||
                !!item.querySelector('.ui-state-active, .ui-menu-item-wrapper.ui-state-active');
              const $ = window.jQuery || window.$;
              const clickTarget = item.querySelector('.ui-menu-item-wrapper') || item;
              const itemData = $ ? ($(item).data('ui-autocomplete-item') ||
                $(item).data('item.autocomplete') ||
                $(clickTarget).data('ui-autocomplete-item') ||
                null) : null;
              const itemDataText = itemData && typeof itemData === 'object'
                ? [itemData.value, itemData.label, itemData.name, itemData.text].filter(Boolean).join(' ')
                : String(itemData || '');
              const itemDataMatches = normalize(itemDataText) === expectedTexts[0] ||
                normalize(itemDataText).split(' ').includes(expectedTexts[0]);
              const matches = (!inputRect || menuInfo.nearInput) &&
                rect.width > 0 &&
                rect.height > 0 &&
                style.visibility !== 'hidden' &&
                style.display !== 'none' &&
                matchedIndex >= 0;
              return {
                item,
                itemIndex,
                menuIndex: menuInfo.menuIndex,
                sortedMenuIndex,
                matchedIndex,
                itemDataMatches,
                matches,
                active,
                nearInput: menuInfo.nearInput,
                distance: menuInfo.distance,
                zIndex: menuInfo.zIndex,
              };
            }))
          .filter(candidate => candidate.matches)
          .sort((a, b) => (a.matchedIndex - b.matchedIndex) ||
            (Number(b.itemDataMatches) - Number(a.itemDataMatches)) ||
            (Number(b.nearInput) - Number(a.nearInput)) ||
            (a.distance - b.distance) ||
            (b.zIndex - a.zIndex) ||
            (a.menuIndex - b.menuIndex) ||
            (a.itemIndex - b.itemIndex));
        const selected = candidates[0];
        if (!selected) return null;
        const targetMenu = sortedMenus[selected.sortedMenuIndex]?.menu;
        const items = targetMenu ? Array.from(targetMenu.querySelectorAll(menuItemSelector)).filter(item => {
          const style = window.getComputedStyle(item);
          const rect = item.getBoundingClientRect();
          return rect.width > 0 &&
            rect.height > 0 &&
            style.visibility !== 'hidden' &&
            style.display !== 'none';
        }) : [];
        const activeIndex = items.findIndex(item => item.classList.contains('ui-state-active') ||
          !!item.querySelector('.ui-state-active, .ui-menu-item-wrapper.ui-state-active'));
        const targetIndex = items.findIndex(item => item === selected.item);
        if (targetIndex < 0) return null;
        input?.focus();
        const clickTarget = selected.item.querySelector('.ui-menu-item-wrapper') || selected.item;
        if (mode === 'mouse') {
          clickTarget.scrollIntoView({ block: 'nearest', inline: 'nearest' });
          const rect = clickTarget.getBoundingClientRect();
          const eventOptions = {
            bubbles: true,
            cancelable: true,
            view: window,
            clientX: rect.left + rect.width / 2,
            clientY: rect.top + rect.height / 2,
            button: 0,
          };
          ['mouseover', 'mouseenter', 'mousemove'].forEach(type => {
            clickTarget.dispatchEvent(new MouseEvent(type, eventOptions));
            if (clickTarget !== selected.item) selected.item.dispatchEvent(new MouseEvent(type, eventOptions));
          });
          document.querySelectorAll('[data-lecture-menu-click-target="true"]').forEach(element => {
            element.removeAttribute('data-lecture-menu-click-target');
          });
          clickTarget.setAttribute('data-lecture-menu-click-target', 'true');
          return {
            method: 'mouse',
            text: (selected.item.textContent || '').trim(),
            x: rect.left + rect.width / 2,
            y: rect.top + rect.height / 2,
          };
        }
        const $ = window.jQuery || window.$;
        if ($ && input) {
          const $input = $(input);
          const autocomplete = $input.data('ui-autocomplete') || $input.data('autocomplete');
          const $item = $(selected.item);
          const itemData = $item.data('ui-autocomplete-item') ||
            $item.data('item.autocomplete') ||
            $(clickTarget).data('ui-autocomplete-item') ||
            null;
          const itemDataText = itemData && typeof itemData === 'object'
            ? [itemData.value, itemData.label, itemData.name, itemData.text].filter(Boolean).join(' ')
            : String(itemData || '');
          const itemDataMatches = normalize(itemDataText) === expectedTexts[0] ||
            normalize(itemDataText).split(' ').includes(expectedTexts[0]);
          if (autocomplete?.menu?.focus && autocomplete?.menu?.select) {
            try {
              if (!itemDataMatches) throw new Error('autocomplete item data is not ready');
              clickTarget.scrollIntoView({ block: 'nearest', inline: 'nearest' });
              const focusEvent = $.Event('mouseover');
              focusEvent.target = clickTarget;
              autocomplete.menu.focus(focusEvent, $item);
              const selectEvent = $.Event('click');
              selectEvent.target = clickTarget;
              autocomplete.menu.select(selectEvent);
              return {
                method: 'jquery-ui',
                text: (selected.item.textContent || '').trim(),
              };
            } catch (e) {
              // Fall through to keyboard selection.
            }
          }
        }
        return {
          method: 'keyboard',
          activeIndex,
          targetIndex,
          text: (selected.item.textContent || '').trim(),
        };
      }, MENU_CONTAINER_SELECTOR, MENU_ITEM_SELECTOR, inputSelector, allowedExpectedTexts, selectionMode);
      if (!selection) {
        throw new Error(`${label}「${targetText}」の候補選択に失敗しました。`);
      }
      if (selection.method === 'jquery-ui') {
        console.log(`${label} 候補「${selection.text}」をautocompleteイベントで選択します。`);
      } else if (selection.method === 'mouse') {
        console.log(`${label} 候補「${selection.text}」をマウスイベントで選択します。`);
        const focusState = await page.evaluate(selector => {
            const marker = document.querySelector('[data-lecture-active-lookup="true"]');
            const active = document.activeElement;
            const input = marker?.matches?.(selector) ? marker : active?.matches?.(selector) ? active : null;
            if (!input) {
              return {
                ok: false,
                activeTag: document.activeElement?.tagName || '',
                activeText: (document.activeElement?.textContent || '').trim().slice(0, 80),
              };
            }
            input.focus();
            const root = input.closest('div[name]') || input.closest('.vv_pill_container') || input.parentElement;
            return {
              ok: document.activeElement === input,
              fieldName: root?.getAttribute('name') || '',
              rootTitle: root?.getAttribute('title') || '',
              activeTag: document.activeElement?.tagName || '',
              inputValue: input.value || '',
            };
        }, inputSelector);
        console.log(`${label} 候補クリック前フォーカス: ${JSON.stringify(focusState)}`);

        const clickPoint = await page.evaluate((menuContainerSelector, menuItemSelector, fieldSelector, expectedText) => {
          const normalize = text => String(text || '')
            .trim()
            .replace(/\s+/g, ' ')
            .toUpperCase();
          const visible = element => {
            if (!element) return false;
            const style = window.getComputedStyle(element);
            const rect = element.getBoundingClientRect();
            return rect.width > 0 &&
              rect.height > 0 &&
              style.visibility !== 'hidden' &&
              style.display !== 'none';
          };
          const expected = normalize(expectedText);
          const marker = document.querySelector('[data-lecture-active-lookup="true"]');
          const active = document.activeElement;
          const input = marker?.matches?.(fieldSelector) ? marker : active?.matches?.(fieldSelector) ? active : null;
          const inputRect = input ? input.getBoundingClientRect() : null;
          const visibleMenus = Array.from(document.querySelectorAll(menuContainerSelector))
            .map((menu, menuIndex) => {
              const rect = menu.getBoundingClientRect();
              const style = window.getComputedStyle(menu);
              const nearInput = !inputRect || (
                rect.top >= inputRect.top - 40 &&
                rect.top <= inputRect.bottom + 500 &&
                rect.right >= inputRect.left - 80 &&
                rect.left <= inputRect.right + 500
              );
              const distance = inputRect
                ? Math.abs(rect.top - inputRect.bottom) + Math.abs(rect.left - inputRect.left)
                : 0;
              return {
                menu,
                menuIndex,
                visible: visible(menu),
                nearInput,
                distance,
                zIndex: Number.parseInt(style.zIndex, 10) || 0,
              };
            })
            .filter(menu => menu.visible);
          const candidates = visibleMenus
            .flatMap(menuInfo => Array.from(menuInfo.menu.querySelectorAll(menuItemSelector)).map((item, itemIndex) => {
              const target = item.querySelector('.ui-menu-item-wrapper') || item;
              const text = (item.textContent || '').trim();
              const rect = target.getBoundingClientRect();
              return {
                item,
                target,
                itemIndex,
                menuIndex: menuInfo.menuIndex,
                text,
                matches: visible(target) && normalize(text) === expected && (!inputRect || menuInfo.nearInput),
                nearInput: menuInfo.nearInput,
                distance: menuInfo.distance,
                zIndex: menuInfo.zIndex,
                y: rect.top,
              };
            }))
            .filter(candidate => candidate.matches)
            .sort((a, b) => Number(b.nearInput) - Number(a.nearInput) ||
              (a.distance - b.distance) ||
              (b.zIndex - a.zIndex) ||
              (a.y - b.y) ||
              (a.menuIndex - b.menuIndex) ||
              (a.itemIndex - b.itemIndex));
          const selected = candidates[0];
          if (!selected) {
            return {
              ok: false,
              reason: 'candidate not found',
              activeTag: document.activeElement?.tagName || '',
              visibleItems: visibleMenus
                .flatMap(menuInfo => Array.from(menuInfo.menu.querySelectorAll(menuItemSelector)).map(item => (item.textContent || '').trim()))
                .filter(Boolean)
                .slice(0, 30),
            };
          }
          input?.focus();
          selected.target.scrollIntoView({ block: 'nearest', inline: 'nearest' });
          const rect = selected.target.getBoundingClientRect();
          const x = rect.left + rect.width / 2;
          const y = rect.top + rect.height / 2;
          const eventOptions = { bubbles: true, cancelable: true, view: window, clientX: x, clientY: y, button: 0 };
          ['mouseover', 'mouseenter', 'mousemove'].forEach(type => {
            selected.target.dispatchEvent(new MouseEvent(type, eventOptions));
            if (selected.target !== selected.item) selected.item.dispatchEvent(new MouseEvent(type, eventOptions));
          });
          const elementAtPoint = document.elementFromPoint(x, y);
          return {
            ok: true,
            text: selected.text,
            x,
            y,
            elementAtPointText: (elementAtPoint?.textContent || '').trim().slice(0, 120),
            elementAtPointTag: elementAtPoint?.tagName || '',
            elementAtPointClass: String(elementAtPoint?.className || '').slice(0, 120),
          };
        }, MENU_CONTAINER_SELECTOR, MENU_ITEM_SELECTOR, inputSelector, targetText);

        console.log(`${label} 候補クリック座標: ${JSON.stringify(clickPoint)}`);
        if (!clickPoint.ok || !Number.isFinite(clickPoint.x) || !Number.isFinite(clickPoint.y)) {
          throw new Error(`${label}「${targetText}」のクリック座標を取得できませんでした: ${JSON.stringify(clickPoint)}`);
        }
        await page.mouse.move(clickPoint.x, clickPoint.y);
        await page.mouse.down();
        await sleep(80);
        await page.mouse.up();
        await page.evaluate(() => {
          document.querySelectorAll('[data-lecture-menu-click-target="true"]').forEach(element => {
            element.removeAttribute('data-lecture-menu-click-target');
          });
        }).catch(() => null);
        await sleep(500);
        const afterState = await page.evaluate(selector => {
          const marker = document.querySelector('[data-lecture-active-lookup="true"]');
          const input = marker?.matches?.(selector) ? marker : document.querySelector(selector);
          const root = input?.closest('div[name]') || input?.closest('.vv_pill_container') || input?.parentElement;
          const rootClone = root?.cloneNode(true);
          rootClone?.querySelectorAll('input, tester, button, svg, .multiItemSelectButtonsContainer, .data-config').forEach(element => element.remove());
          return {
            activeTag: document.activeElement?.tagName || '',
            activeName: document.activeElement?.getAttribute('name') || '',
            fieldName: root?.getAttribute('name') || '',
            selectedText: (rootClone?.textContent || '').trim(),
            hiddenValues: Array.from(root?.querySelectorAll('input[type="hidden"], input.singleDataInputName') || [])
              .map(element => element.value)
              .filter(Boolean),
          };
        }, inputSelector).catch(err => ({ error: String(err) }));
        console.log(`${label} 候補クリック後状態: ${JSON.stringify(afterState)}`);
      } else {
        const steps = selection.activeIndex >= 0
          ? selection.targetIndex - selection.activeIndex
          : selection.targetIndex + 1;
        const key = steps < 0 ? 'ArrowUp' : 'ArrowDown';
        console.log(`${label} 候補「${selection.text}」をキーボード選択します。`);
        for (let i = 0; i < Math.abs(steps); i += 1) {
          await page.keyboard.press(key);
        }
        await page.keyboard.press('Enter');
      }
    }

    async function focusClearAndTypeLookup(inputSelector, targetText, delay = 120, options = {}) {
      const targetIndex = await getVisibleElementIndex(inputSelector, NORMAL_WAIT, {
        sectionTitle: options.sectionTitle || '',
      });
      const handles = await page.$$(inputSelector);
      let handle = handles[targetIndex];
      if (!handle) {
        throw new Error(`${inputSelector} の入力対象を取得できませんでした。`);
      }
      let targetInfo = await page.evaluate((selector, index, clearSelectedItems) => {
        Array.from(document.querySelectorAll('[data-lecture-active-lookup="true"]')).forEach(element => {
          element.removeAttribute('data-lecture-active-lookup');
        });
        const allTargets = Array.from(document.querySelectorAll(selector));
        const target = allTargets[index];
        const visibleCount = allTargets.filter(element => {
          const style = window.getComputedStyle(element);
          const rect = element.getBoundingClientRect();
          return rect.width > 0 &&
            rect.height > 0 &&
            !element.disabled &&
            element.getAttribute('aria-disabled') !== 'true' &&
            style.visibility !== 'hidden' &&
            style.display !== 'none';
        }).length;
        if (!target) return;
        target.setAttribute('data-lecture-active-lookup', 'true');
        target.scrollIntoView({ block: 'center', inline: 'center' });
        target.focus();
        const root = target.closest('div[name]') || target.closest('.vv_pill_container') || target.parentElement;
        const rootClone = root?.cloneNode(true);
        rootClone?.querySelectorAll('input, tester, button, svg, .multiItemSelectButtonsContainer, .data-config').forEach(element => element.remove());
        const hiddenValues = Array.from(root?.querySelectorAll('input[type="hidden"], input.singleDataInputName') || [])
          .map(element => element.value)
          .filter(Boolean);
        let removedCount = 0;
        if (clearSelectedItems && root) {
          const removeTargets = Array.from(root.querySelectorAll('.multiItemSelectAutoComplete.removeItem, .removeItem'))
            .filter(element => {
              const style = window.getComputedStyle(element);
              const rect = element.getBoundingClientRect();
              return rect.width > 0 &&
                rect.height > 0 &&
                style.visibility !== 'hidden' &&
                style.display !== 'none';
            });
          removeTargets.forEach(element => {
            const rect = element.getBoundingClientRect();
            const eventOptions = {
              bubbles: true,
              cancelable: true,
              view: window,
              clientX: rect.left + rect.width / 2,
              clientY: rect.top + rect.height / 2,
              button: 0,
            };
            ['mousedown', 'mouseup', 'click'].forEach(type => {
              element.dispatchEvent(new MouseEvent(type, eventOptions));
            });
            removedCount += 1;
          });
          target.focus();
        }
        return {
          index: index + 1,
          total: allTargets.length,
          visibleCount,
          fieldName: root?.getAttribute('name') || '',
          rootTitle: root?.getAttribute('title') || '',
          rootClassName: String(root?.className || '').slice(0, 120),
          currentValue: target.value || '',
          selectedText: (rootClone?.textContent || '').trim(),
          hiddenValues,
          removedCount,
        };
      }, inputSelector, targetIndex, options.clearSelectedItems === true);
      if (options.label && targetInfo) {
        console.log(`${options.label} 入力対象: ${targetInfo.fieldName || inputSelector} / ${targetInfo.rootTitle || '(titleなし)'} (${targetInfo.index}/${targetInfo.total}, 表示${targetInfo.visibleCount})`);
      }
      if (options.clearSelectedItems && targetInfo?.removedCount) {
        console.log(`${options.label || targetText} 既存選択を${targetInfo.removedCount}件クリアしました。`);
        await sleep(500);
        const refreshedHandles = await page.$$(inputSelector);
        handle = refreshedHandles[targetIndex] || handle;
        targetInfo = await page.evaluate(selector => {
          const marker = document.querySelector('[data-lecture-active-lookup="true"]');
          const target = marker?.matches?.(selector) ? marker : document.querySelector(selector);
          const root = target?.closest('div[name]') || target?.closest('.vv_pill_container') || target?.parentElement;
          const rootClone = root?.cloneNode(true);
          rootClone?.querySelectorAll('input, tester, button, svg, .multiItemSelectButtonsContainer, .data-config').forEach(element => element.remove());
          const hiddenValues = Array.from(root?.querySelectorAll('input[type="hidden"], input.singleDataInputName') || [])
            .map(element => element.value)
            .filter(Boolean);
          return {
            fieldName: root?.getAttribute('name') || '',
            rootTitle: root?.getAttribute('title') || '',
            rootClassName: String(root?.className || '').slice(0, 120),
            currentValue: target?.value || '',
            selectedText: (rootClone?.textContent || '').trim(),
            hiddenValues,
          };
        }, inputSelector);
      }
      await handle.click({ clickCount: 3 });
      await handle.evaluate(el => {
        el.focus();
        const prototype = el.tagName === 'TEXTAREA'
          ? HTMLTextAreaElement.prototype
          : HTMLInputElement.prototype;
        const valueSetter = Object.getOwnPropertyDescriptor(prototype, 'value')?.set;
        if (valueSetter) {
          valueSetter.call(el, '');
        } else {
          el.value = '';
        }
        el.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'deleteContentBackward' }));
        el.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true, key: 'Backspace' }));
      });
      const modifier = process.platform === 'darwin' ? 'Meta' : 'Control';
      await page.keyboard.down(modifier);
      await page.keyboard.press('A');
      await page.keyboard.up(modifier);
      await page.keyboard.press('Backspace');
      await handle.type(targetText, { delay });
      try {
        await page.waitForFunction(
          expected => {
            const input = document.querySelector('[data-lecture-active-lookup="true"]');
            return document.activeElement === input && input?.value === expected;
          },
          { timeout: SHORT_WAIT },
          targetText
        );
      } catch (e) {
        const forcedState = await page.evaluate((selector, expected) => {
          const marker = document.querySelector('[data-lecture-active-lookup="true"]');
          const active = document.activeElement;
          const input = marker?.matches?.(selector) ? marker : active?.matches?.(selector) ? active : null;
          if (!input) return { ok: false, reason: 'active lookup input not found' };
          input.focus();
          const valueSetter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set;
          if (valueSetter) {
            valueSetter.call(input, '');
            input.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'deleteContentBackward' }));
            valueSetter.call(input, expected);
          } else {
            input.value = expected;
          }
          input.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText', data: expected }));
          input.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true, key: expected.slice(-1) || ' ' }));
          const $ = window.jQuery || window.$;
          if ($) {
            const $input = $(input);
            $input.trigger('input').trigger('keyup');
            try {
              if (typeof $input.autocomplete === 'function') {
                $input.autocomplete('search', expected);
              }
            } catch (err) {
              // Ignore and fall back to native input events.
            }
            const autocomplete = $input.data('ui-autocomplete') || $input.data('autocomplete');
            try {
              autocomplete?.search?.(expected);
            } catch (err) {
              // Ignore and fall back to native input events.
            }
          }
          return {
            ok: true,
            value: input.value,
            active: document.activeElement === input,
          };
        }, inputSelector, targetText);
        console.log(`${options.label || targetText} キーボード入力を確認できなかったためlookup値を再設定しました: ${JSON.stringify(forcedState)}`);
        await page.waitForFunction(
          expected => {
            const input = document.querySelector('[data-lecture-active-lookup="true"]');
            return document.activeElement === input && input?.value === expected;
          },
          { timeout: NORMAL_WAIT },
          targetText
        );
      }
      const searchState = await page.evaluate((selector, expected) => {
        const marker = document.querySelector('[data-lecture-active-lookup="true"]');
        const active = document.activeElement;
        const input = marker?.matches?.(selector) ? marker : active?.matches?.(selector) ? active : document.querySelector(selector);
        if (!input) {
          return {
            ok: false,
            reason: 'active lookup input not found',
            activeTag: document.activeElement?.tagName || '',
          };
        }
        input.focus();
        const valueSetter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set;
        if (input.value !== expected) {
          if (valueSetter) {
            valueSetter.call(input, expected);
          } else {
            input.value = expected;
          }
        }
        const eventOptions = { bubbles: true, cancelable: true };
        if (valueSetter) {
          valueSetter.call(input, expected);
        } else {
          input.value = expected;
        }
        input.dispatchEvent(new KeyboardEvent('keydown', { ...eventOptions, key: expected.slice(-1) || ' ' }));
        input.dispatchEvent(new InputEvent('input', { ...eventOptions, inputType: 'insertText', data: expected }));
        input.dispatchEvent(new KeyboardEvent('keyup', { ...eventOptions, key: expected.slice(-1) || ' ' }));
        const $ = window.jQuery || window.$;
        let autocompleteTriggered = false;
        let jqueryTriggered = false;
        if ($) {
          const $input = $(input);
          try {
            $input.val(expected);
            $input.trigger('input').trigger('keydown').trigger('keyup');
            jqueryTriggered = true;
          } catch (err) {
            // Native events above are enough for non-jQuery fields.
          }
          try {
            if (typeof $input.autocomplete === 'function') {
              $input.autocomplete('search', expected);
              autocompleteTriggered = true;
            }
          } catch (err) {
            // Some Veeva lookup widgets expose autocomplete only through instance data.
          }
          const autocomplete = $input.data('ui-autocomplete') || $input.data('autocomplete');
          try {
            if (autocomplete) {
              autocomplete.term = null;
              autocomplete.search?.(expected);
              autocompleteTriggered = true;
            }
          } catch (err) {
            // Ignore and keep the typed value.
          }
        }
        if (valueSetter) {
          valueSetter.call(input, expected);
        } else {
          input.value = expected;
        }
        return {
          ok: true,
          value: input.value,
          active: document.activeElement === input,
          jqueryTriggered,
          autocompleteTriggered,
        };
      }, inputSelector, targetText);
      console.log(`${options.label || targetText} lookup検索を発火しました: ${JSON.stringify(searchState)}`);
      await sleep(options.searchStabilizeMs || 700);
      return targetInfo;
    }

    async function typeAndSelectMenuItem(inputSelector, target, options = {}) {
      const targetText = String(target ?? '');
      const label = options.label || targetText;
      await closeVisibleMenus();
      let lookupBaseline = await focusClearAndTypeLookup(inputSelector, targetText, 120, {
        label,
        sectionTitle: options.sectionTitle || '',
        clearSelectedItems: options.clearSelectedItems === true,
      });
      let selectionMode = options.selectionMode ||
        (/multiItemSelectContainer|multiItemSelect/i.test(lookupBaseline?.rootClassName || '') ? 'mouse' : 'auto');
      try {
        await waitAndSelectMenuItem(targetText, {
          inputSelector,
          label,
          timeout: options.timeout || NORMAL_WAIT,
          selectionMode,
        });
      } catch (firstError) {
        console.log(`${label}「${targetText}」の候補取得を再試行します: ${firstError.message}`);
        await closeVisibleMenus();
        lookupBaseline = await focusClearAndTypeLookup(inputSelector, targetText, 160, {
          label,
          sectionTitle: options.sectionTitle || '',
          clearSelectedItems: options.clearSelectedItems === true,
        });
        selectionMode = options.selectionMode ||
          (/multiItemSelectContainer|multiItemSelect/i.test(lookupBaseline?.rootClassName || '') ? 'mouse' : 'auto');
        await page.keyboard.press('ArrowDown').catch(() => null);
        await waitAndSelectMenuItem(targetText, {
          inputSelector,
          label,
          timeout: options.timeout || NORMAL_WAIT,
          selectionMode,
        });
      }
      try {
        await page.waitForFunction(
          (selector, expected, baselineHiddenValues, baselineSelectedText, requireSelectedText) => {
            const normalize = text => String(text || '')
              .trim()
              .replace(/\s+/g, ' ')
              .toUpperCase();
            const matchesExpected = text => {
              const normalized = normalize(text);
              if (normalized === expected) return true;
              const escaped = expected.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
              return new RegExp(`(^|[^A-Z0-9])${escaped}($|[^A-Z0-9])`).test(normalized);
            };
            const marker = document.querySelector('[data-lecture-active-lookup="true"]');
            const active = document.activeElement;
            const input = marker?.matches?.(selector) ? marker : active?.matches?.(selector) ? active : null;
            const root = input?.closest('div[name]') || input?.closest('.vv_pill_container') || input?.parentElement;
            const rootClone = root?.cloneNode(true);
            rootClone?.querySelectorAll('input, tester, button, svg, .multiItemSelectButtonsContainer, .data-config').forEach(element => element.remove());
            const selectedText = rootClone?.textContent || '';
            const hiddenValues = Array.from(root?.querySelectorAll('input[type="hidden"], input.singleDataInputName') || [])
              .map(element => element.value)
              .filter(Boolean);
            const hiddenChanged = hiddenValues.join('|') !== (baselineHiddenValues || []).join('|');
            const visibleMenus = Array.from(document.querySelectorAll('.ui-menu-item')).filter(element => {
              const style = window.getComputedStyle(element);
              const rect = element.getBoundingClientRect();
              return rect.width > 0 &&
                rect.height > 0 &&
                style.visibility !== 'hidden' &&
                style.display !== 'none';
            });
            const values = [
              input?.getAttribute('title'),
              input?.getAttribute('aria-label'),
              selectedText,
              root?.getAttribute('title'),
              root?.getAttribute('aria-label'),
            ];
            root?.querySelectorAll?.('input, [title], [aria-label], [data-label], [data-value], [data-name], [data-text]').forEach(element => {
              if (element === input) return;
              values.push(
                element.value,
                element.textContent,
                element.getAttribute('title'),
                element.getAttribute('aria-label'),
                element.getAttribute('data-label'),
                element.getAttribute('data-value'),
                element.getAttribute('data-name'),
                element.getAttribute('data-text')
              );
            });
            const hasSelectedText = values.some(value => matchesExpected(value));
            const inputHasUncommittedText = normalize(input?.value || '') === expected;
            const committedBySelectedText = hasSelectedText && hiddenValues.length > 0;
            const committedByHiddenChange = !requireSelectedText &&
              hiddenValues.length > 0 &&
              hiddenChanged &&
              !inputHasUncommittedText;
            return committedBySelectedText || (visibleMenus.length === 0 && committedByHiddenChange);
          },
          { timeout: NORMAL_WAIT },
          inputSelector,
          normalizeMenuText(targetText),
          lookupBaseline?.hiddenValues || [],
          lookupBaseline?.selectedText || '',
          options.requireSelectedText === true
        );
        console.log(`${options.label || targetText}「${targetText}」を選択しました。`);
      } catch (e) {
        const fieldState = await page.$eval(inputSelector, (el, selector) => {
          const marker = document.querySelector('[data-lecture-active-lookup="true"]');
          const target = marker?.matches?.(selector) ? marker : el;
          const root = target.closest('div[name]') || target.closest('.vv_pill_container') || target.parentElement;
          const rootClone = root?.cloneNode(true);
          rootClone?.querySelectorAll('input, tester, button, svg, .multiItemSelectButtonsContainer, .data-config').forEach(element => element.remove());
          const hiddenValues = Array.from(root?.querySelectorAll('input[type="hidden"], input.singleDataInputName') || [])
            .map(element => element.value)
            .filter(Boolean);
          const textSources = [];
          root?.querySelectorAll?.('input, [title], [aria-label], [data-label], [data-value], [data-name], [data-text]').forEach(element => {
            textSources.push({
              tag: element.tagName,
              className: String(element.className || ''),
              value: element.value || '',
              text: (element.textContent || '').trim(),
              title: element.getAttribute('title') || '',
              ariaLabel: element.getAttribute('aria-label') || '',
              dataLabel: element.getAttribute('data-label') || '',
              dataValue: element.getAttribute('data-value') || '',
              dataName: element.getAttribute('data-name') || '',
              dataText: element.getAttribute('data-text') || '',
            });
          });
          return {
            inputValue: target.value,
            rootText: (root?.textContent || '').trim(),
            rootTitle: root?.getAttribute('title') || '',
            rootClassName: String(root?.className || ''),
            fieldName: root?.getAttribute('name') || '',
            selectedText: (rootClone?.textContent || '').trim(),
            hiddenValues,
            rootHtml: (root?.outerHTML || '').slice(0, 1200),
            textSources: textSources.slice(0, 12),
            activeTag: document.activeElement?.tagName,
            activeName: document.activeElement?.getAttribute('name'),
          };
        }, inputSelector);
        const normalizedExpected = normalizeMenuText(targetText);
        const selectedAlreadyCommitted = fieldState.hiddenValues?.length > 0 &&
          [fieldState.selectedText, fieldState.rootText].some(value => menuTextMatchesTarget(value, normalizedExpected));
        if (selectedAlreadyCommitted) {
          console.log(`${options.label || targetText}「${targetText}」は選択済みとして続行します。状態: ${JSON.stringify({
            selectedText: fieldState.selectedText,
            hiddenValues: fieldState.hiddenValues,
          })}`);
        } else {
          throw new Error(`候補「${targetText}」を選択できませんでした。現在の入力欄: ${JSON.stringify(fieldState)}`);
        }
      }
      await page.evaluate(() => {
        const marker = document.querySelector('[data-lecture-active-lookup="true"]');
        if (marker) {
          marker.blur();
          return;
        }
        document.activeElement?.blur();
      }).catch(() => null);
      try {
        await page.waitForFunction(() => {
          return !Array.from(document.querySelectorAll('.ui-menu-item')).some(element => {
            const style = window.getComputedStyle(element);
            const rect = element.getBoundingClientRect();
            return rect.width > 0 &&
              rect.height > 0 &&
              style.visibility !== 'hidden' &&
              style.display !== 'none';
          });
        }, { timeout: SHORT_WAIT });
      } catch (e) {
        await page.keyboard.press('Escape');
        await page.evaluate(() => document.activeElement?.blur());
      }
    }

    async function waitForDelegateUserSwitch(expectedUser, timeout = LONG_WAIT) {
      const expected = normalizeMenuText(expectedUser);
      try {
        await page.waitForFunction(
          expectedText => {
            const normalize = text => String(text || '')
              .trim()
              .replace(/\s+/g, ' ')
              .toUpperCase();
            const username = normalize(document.querySelector(".vv_username")?.textContent || '');
            return username === expectedText || username.includes(expectedText);
          },
          { timeout },
          expected
        );
      } catch (e) {
        const state = await page.evaluate(() => ({
          url: location.href,
          username: (document.querySelector(".vv_username")?.textContent || '').trim(),
          searchBoxVisible: !!document.querySelector("#search_main_box"),
          delegateInputValue: document.querySelector("#delegateAccessControl input")?.value || '',
        })).catch(err => ({ error: String(err) }));
        throw new Error(`代理アクセスユーザーへの切り替えを確認できませんでした。現在の状態: ${JSON.stringify(state)}`);
      }
    }

    async function waitForRenditionToFinish(timeout = LONG_WAIT) {
      await Promise.race([
        waitForOptionalSelector(".generatingRenditionSpinner", SHORT_WAIT, { visible: true }),
        waitForOptionalSelector(".generatingRenditionLabel", SHORT_WAIT, { visible: true }),
      ]);
      await page.waitForSelector(".generatingRenditionSpinner", { hidden: true, timeout });
    }

    async function getSearchResultSummary(baseUrl) {
      return page.evaluate(vaultUrl => {
        const paginatorText = document.querySelector(".vv-expanded-search-paginator")?.textContent?.trim() || '';
        const gridText = document.querySelector(".vv-document-search-vcl-data-grid")?.textContent?.trim() || '';
        const normalizedGridText = gridText.toLowerCase();
        const noItems = /no items found|no results found|no documents found/.test(normalizedGridText);
        const linkIds = Array.from(document.querySelectorAll(".vv-document-search-vcl-data-grid a[data-linkid]"))
          .map(link => link.getAttribute("data-linkid"))
          .filter(Boolean);
        const uniqueLinkIds = Array.from(new Set(linkIds));
        const countMatch = paginatorText.match(/of\s+(?:about\s+)?([\d,]+)/i);
        const count = countMatch
          ? Number(countMatch[1].replace(/,/g, ''))
          : noItems
            ? 0
            : uniqueLinkIds.length > 0
              ? uniqueLinkIds.length
              : null;

        return {
          count,
          countSource: countMatch ? 'paginator' : noItems ? 'no-items' : uniqueLinkIds.length > 0 ? 'grid-links' : 'unknown',
          paginatorText,
          gridText,
          resultRows: Array.from(document.querySelectorAll(".vv-document-search-vcl-data-grid tr, .vv-document-search-vcl-data-grid .vv_grid_row, .vv-document-search-vcl-data-grid [role=row]"))
            .map(row => (row.textContent || '').replace(/\s+/g, ' ').trim())
            .filter(Boolean)
            .slice(0, 10),
          urls: uniqueLinkIds.map(id => `${vaultUrl}/ui/#doc_info/${id}`),
        };
      }, baseUrl);
    }



    async function searchExistingDocument(searchName, baseUrl) {
      const expectedSearchName = String(searchName || '').trim();
      if (!expectedSearchName) {
        return {
          count: null,
          countSource: 'invalid-search-name',
          paginatorText: '',
          gridText: '',
          resultRows: [],
          urls: [],
          error: '検索語が空です。',
        };
      }
      let lastSummary = null;
      let lastInputValue = '';
      let lastError = '';
      for (let attempt = 1; attempt <= 3; attempt++) {
        try {
          await clearAndType("#search_main_box", expectedSearchName, { delay: attempt === 1 ? 120 : 200 });
        } catch (e) {
          lastError = e.message;
          console.log(`検索ボックス入力に失敗しました。再入力します。(${attempt}/3): ${e.message}`);
          continue;
        }

        const inputValue = await page.$eval("#search_main_box", el => el.value.trim());
        lastInputValue = inputValue;
        if (inputValue !== expectedSearchName) {
          lastError = `検索ボックス入力値が期待値と一致しません。期待値: "${expectedSearchName}", 入力値: "${inputValue}"`;
          console.log(`${lastError} 再入力します。(${attempt}/3)`);
          continue;
        }

        const beforeSummary = await getSearchResultSummary(baseUrl).catch(() => null);
        console.log(`${expectedSearchName}で既存登録を検索します。(${attempt}/3)`);
        await page.click("#search_main_button");

        try {
          await page.waitForFunction((expected, previousGridText, previousPaginatorText) => {
            const inputValue = document.querySelector("#search_main_box")?.value?.trim() || '';
            if (inputValue !== expected) return false;
            const paginatorText = document.querySelector(".vv-expanded-search-paginator")?.textContent || '';
            const gridText = document.querySelector(".vv-document-search-vcl-data-grid")?.textContent || '';
            const resultReady = /of\s+[\d,]+/i.test(paginatorText) ||
              /no items found|no results found|no documents found/i.test(gridText) ||
              !!document.querySelector(".vv-document-search-vcl-data-grid a[data-linkid]");
            if (!resultReady) return false;
            return gridText !== previousGridText || paginatorText !== previousPaginatorText || /no items found|no results found|no documents found/i.test(gridText);
          }, { timeout: 45000 }, expectedSearchName, beforeSummary?.gridText || '', beforeSummary?.paginatorText || '');
        } catch (e) {
          // ヘッダーだけ表示される中途半端な状態なら下のsummary判定でリトライする
        }

        const postInputValue = await page.$eval("#search_main_box", el => el.value.trim()).catch(() => '');
        lastInputValue = postInputValue;
        if (postInputValue !== expectedSearchName) {
          lastError = `検索実行後の入力値が期待値と一致しません。期待値: "${expectedSearchName}", 入力値: "${postInputValue}"`;
          console.log(`${lastError} 検索結果を採用せずリトライします。(${attempt}/3)`);
          continue;
        }

        lastSummary = await getSearchResultSummary(baseUrl);
        if (lastSummary.count !== null) {
          return {
            ...lastSummary,
            searchName: expectedSearchName,
            verified: true,
            inputValue: postInputValue,
          };
        }

        console.log(`検索結果の件数をまだ判定できません。検索をリトライします。(${attempt}/3)`);
        console.log("paginator:", lastSummary.paginatorText || "(なし)");
        console.log("grid:", lastSummary.gridText ? lastSummary.gridText.slice(0, 200) : "(なし)");
      }

      return {
        ...(lastSummary || {
          count: null,
          countSource: 'not-verified',
          paginatorText: '',
          gridText: '',
          resultRows: [],
          urls: [],
        }),
        count: null,
        searchName: expectedSearchName,
        verified: false,
        inputValue: lastInputValue,
        error: lastError || '検索結果を厳格に確認できませんでした。',
      };
    }

    async function findVisibleSelector(selectors, timeout = NORMAL_WAIT) {
      await page.waitForFunction(selectorList => {
        return selectorList.some(selector => {
          return Array.from(document.querySelectorAll(selector)).some(element => {
            const style = window.getComputedStyle(element);
            const rect = element.getBoundingClientRect();
            return rect.width > 0 &&
              rect.height > 0 &&
              !element.disabled &&
              element.getAttribute('aria-disabled') !== 'true' &&
              style.visibility !== 'hidden' &&
              style.display !== 'none';
          });
        });
      }, { timeout }, selectors);

      const selector = await page.evaluate(selectorList => {
        return selectorList.find(candidate => {
          return Array.from(document.querySelectorAll(candidate)).some(element => {
            const style = window.getComputedStyle(element);
            const rect = element.getBoundingClientRect();
            return rect.width > 0 &&
              rect.height > 0 &&
              !element.disabled &&
              element.getAttribute('aria-disabled') !== 'true' &&
              style.visibility !== 'hidden' &&
              style.display !== 'none';
          });
        });
      }, selectors);

      if (!selector) {
        throw new Error(`表示中の要素が見つかりませんでした: ${selectors.join(', ')}`);
      }
      return selector;
    }

    async function getVisibleElementHandle(selector, timeout = NORMAL_WAIT) {
      await waitForClickable(selector, timeout);
      const handles = await page.$$(selector);
      for (const handle of handles) {
        const visible = await handle.evaluate(element => {
          const style = window.getComputedStyle(element);
          const rect = element.getBoundingClientRect();
          return rect.width > 0 &&
            rect.height > 0 &&
            !element.disabled &&
            element.getAttribute('aria-disabled') !== 'true' &&
            style.visibility !== 'hidden' &&
            style.display !== 'none';
        });
        if (visible) return handle;
      }
      throw new Error(`${selector} の表示中要素を取得できませんでした。`);
    }

    async function clearAndTypeHandle(handle, value, options = {}, label = 'input') {
      const textValue = String(value ?? '');

      let typedByKeyboard = false;
      try {
        await handle.evaluate(element => {
          element.scrollIntoView({ block: 'center', inline: 'center' });
          element.focus();
        });
        await handle.click({ clickCount: 3, delay: 30 });
        const modifier = process.platform === 'darwin' ? 'Meta' : 'Control';
        await page.keyboard.down(modifier);
        await page.keyboard.press('A');
        await page.keyboard.up(modifier);
        await page.keyboard.press('Backspace');
        if (textValue) {
          await page.keyboard.type(textValue, { delay: 120, ...options });
        }
        typedByKeyboard = true;
      } catch (e) {
        console.log(`${label} は直接クリックできないため、値を直接設定します。`);
      }

      const actualValue = await handle.evaluate(element => element.value);
      if (typedByKeyboard && actualValue === textValue) {
        await page.waitForFunction(
          (element, expected) => element.isConnected && element.value === expected,
          { timeout: SHORT_WAIT },
          handle,
          textValue
        );
        return handle;
      }

      await handle.evaluate((element, expected) => {
        element.scrollIntoView({ block: 'center', inline: 'center' });
        element.focus();
        const prototype = element.tagName === 'TEXTAREA'
          ? HTMLTextAreaElement.prototype
          : HTMLInputElement.prototype;
        const valueSetter = Object.getOwnPropertyDescriptor(prototype, 'value')?.set;
        if (valueSetter) {
          valueSetter.call(element, expected);
        } else {
          element.value = expected;
        }
        element.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText', data: expected }));
        element.dispatchEvent(new Event('change', { bubbles: true }));
      }, textValue);

      await page.waitForFunction(
        (element, expected) => element.isConnected && element.value === expected,
        { timeout: SHORT_WAIT },
        handle,
        textValue
      );
      return handle;
    }

    async function clearAndTypeVisible(selector, value, options = {}) {
      const handle = await getVisibleElementHandle(selector, NORMAL_WAIT);
      return clearAndTypeHandle(handle, value, options, selector);
    }

    async function getSharedResourceSearchInputHandle() {
      const inputSelector = [
        '.vv_search_box input',
        '.vv-search-box input',
        '.vv-search input',
        '.NonResponsiveDialog input[type=text]',
        '[role="dialog"] input[type=text]',
        '.ui-dialog input[type=text]',
        'input[type=search]',
        'input[class*="search"]',
      ].join(', ');

      await page.waitForFunction(selector => {
        const isVisible = element => {
          const style = window.getComputedStyle(element);
          const rect = element.getBoundingClientRect();
          return rect.width > 0 &&
            rect.height > 0 &&
            !element.disabled &&
            element.getAttribute('aria-disabled') !== 'true' &&
            style.visibility !== 'hidden' &&
            style.display !== 'none';
        };
        const isGlobalSearch = input => {
          const label = [
            input.id || '',
            input.getAttribute('placeholder') || '',
            input.getAttribute('aria-label') || '',
            input.className || '',
          ].join(' ');
          return input.id === 'search_main_box' ||
            /マイドキュメント|my\s*documents/i.test(label);
        };
        return Array.from(document.querySelectorAll(selector)).some(input => isVisible(input) && !isGlobalSearch(input));
      }, { timeout: LONG_WAIT }, inputSelector);

      const targetIndex = await page.evaluate(selector => {
        const isVisible = element => {
          const style = window.getComputedStyle(element);
          const rect = element.getBoundingClientRect();
          return rect.width > 0 &&
            rect.height > 0 &&
            !element.disabled &&
            element.getAttribute('aria-disabled') !== 'true' &&
            style.visibility !== 'hidden' &&
            style.display !== 'none';
        };
        const isGlobalSearch = input => {
          const label = [
            input.id || '',
            input.getAttribute('placeholder') || '',
            input.getAttribute('aria-label') || '',
            input.className || '',
          ].join(' ');
          return input.id === 'search_main_box' ||
            /マイドキュメント|my\s*documents/i.test(label);
        };
        const scoreInput = input => {
          const dialog = input.closest('.NonResponsiveDialog, [role="dialog"], .ui-dialog, [class*="Dialog"]');
          const localSearchBox = input.closest('.vv_search_box, .vv-search-box, .vv-search');
          const placeholder = input.getAttribute('placeholder') || '';
          const rect = input.getBoundingClientRect();
          let score = 0;
          if (dialog) score += 100;
          if (localSearchBox) score += 50;
          if (/search|検索/i.test(placeholder)) score += 10;
          score -= Math.max(0, rect.top / 1000);
          return score;
        };

        const inputs = Array.from(document.querySelectorAll(selector));
        const candidates = inputs
          .map((input, index) => ({ input, index, score: scoreInput(input) }))
          .filter(item => isVisible(item.input) && !isGlobalSearch(item.input))
          .sort((a, b) => b.score - a.score);
        return candidates[0]?.index ?? -1;
      }, inputSelector);

      if (targetIndex < 0) {
        throw new Error("Shared Resource検索欄が見つかりませんでした。背面のグローバル検索欄は除外しています。");
      }

      const handles = await page.$$(inputSelector);
      return handles[targetIndex];
    }

    async function triggerSearchFromInput(inputHandle) {
      await inputHandle.focus();
      await page.keyboard.press('Enter').catch(() => null);

      return inputHandle.evaluate(input => {
        const isVisible = element => {
          if (!element) return false;
          const style = window.getComputedStyle(element);
          const rect = element.getBoundingClientRect();
          return rect.width > 0 &&
            rect.height > 0 &&
            !element.disabled &&
            element.getAttribute('aria-disabled') !== 'true' &&
            style.visibility !== 'hidden' &&
            style.display !== 'none' &&
            style.pointerEvents !== 'none';
        };

        const inputRect = input.getBoundingClientRect();
        const roots = [
          input.closest('.vv_search_box'),
          input.closest('.vv-search-box'),
          input.closest('.vv-search'),
          input.parentElement,
          input.closest('.NonResponsiveDialog'),
          input.closest('[role="dialog"]'),
        ].filter(Boolean);

        for (const root of roots) {
          const candidates = Array.from(root.querySelectorAll('button, a, [role="button"]'))
            .filter(isVisible)
            .map(element => {
              const rect = element.getBoundingClientRect();
              const text = element.textContent || '';
              const title = element.getAttribute('title') || '';
              const ariaLabel = element.getAttribute('aria-label') || '';
              const classText = element.getAttribute('class') || '';
              const label = [
                text,
                title,
                ariaLabel,
                classText,
              ].join(' ');
              const sameRow = rect.bottom >= inputRect.top - 20 &&
                rect.top <= inputRect.bottom + 40 &&
                rect.left >= inputRect.left - 20 &&
                rect.left <= inputRect.right + 240;
              const isAdvancedSearch = /advancedSearchLink|vv_search_lookup|詳細検索|advanced\s*search|lookup/i.test(label);
              const exactSearchLabel = [text, title, ariaLabel].some(value => /^(search|検索)$/i.test(String(value || '').trim()));
              const looksSearch = exactSearchLabel ||
                /(^|[\s_-])(search|searchButton|search-button|search_icon|fa-search|magnif)([\s_-]|$)/i.test(label);
              const distance = Math.abs(rect.left - inputRect.right) + Math.abs(rect.top - inputRect.top);
              return { element, sameRow, looksSearch, isAdvancedSearch, distance, label: label.trim() };
            })
            .filter(item => !item.isAdvancedSearch && item.looksSearch && (item.sameRow || item.distance < 300))
            .sort((a, b) => {
              if (a.looksSearch !== b.looksSearch) return a.looksSearch ? -1 : 1;
              return a.distance - b.distance;
            });

          const target = candidates[0];
          if (target) {
            target.element.scrollIntoView({ block: 'center', inline: 'center' });
            target.element.click();
            return { clicked: true, label: target.label };
          }
        }

        input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', bubbles: true, cancelable: true }));
        input.dispatchEvent(new KeyboardEvent('keypress', { key: 'Enter', code: 'Enter', bubbles: true, cancelable: true }));
        input.dispatchEvent(new KeyboardEvent('keyup', { key: 'Enter', code: 'Enter', bubbles: true, cancelable: true }));
        return { clicked: false, label: 'Enter only (通常検索ボタンなし)' };
      });
    }

    async function getSharedSearchSummary(sharedName) {
      return page.evaluate((paginationSelector, resultListSelector, expectedName) => {
        const normalizeName = value => String(value || '')
          .replace(/\s*\(v[\d.]+\)\s*$/i, '')
          .replace(/\s+/g, ' ')
          .trim();
        const inferName = item => {
          const explicitName = item.querySelector('.vv-search-result-name, .docName, .docNameLink, .vv_doc_title_name')?.textContent?.trim() || '';
          if (explicitName) return explicitName;
          const text = item.textContent?.trim() || '';
          const versionMatch = text.match(/^(.+?\s*\(v[\d.]+\))/i);
          if (versionMatch) return versionMatch[1].trim();
          const docNumberMatch = text.match(/^(.*?)(VV-\d+|MCS-\d+|ドラフト|Draft|Approved|承認済)/i);
          return (docNumberMatch?.[1] || text).trim();
        };
        const getResultRows = resultList => {
          if (!resultList) return [];
          const directRows = Array.from(resultList.children).filter(child =>
            child.matches('li, .binderDocRow, .vv-doc-compact-item, .vv_veeva_document')
          );
          if (directRows.length > 0) return directRows;
          const nestedRows = Array.from(resultList.querySelectorAll('li, .binderDocRow, .vv-doc-compact-item, .vv_veeva_document'));
          return nestedRows.length > 0 ? nestedRows : [resultList];
        };
        const paginationText = document.querySelector(paginationSelector)?.textContent?.trim() || '';
        const normalizedPaginationText = paginationText.replace(/[〜～]/g, '~');
        const resultListText = document.querySelector(resultListSelector)?.textContent?.trim() || '';
        const noItemsText = document.querySelector('.vv-compact-result-no-items')?.textContent?.trim() || '';
        const countMatch = paginationText.match(/of\s+([\d,]+)/i) ||
          normalizedPaginationText.match(/\/\s*([\d,]+)\s*$/);
        const count = countMatch
          ? Number(countMatch[1].replace(/,/g, ''))
          : /no items found|no results found|no documents found/i.test(noItemsText)
            ? 0
            : null;
        const resultList = document.querySelector(resultListSelector);
        const resultItems = getResultRows(resultList).map((item, index) => {
          const name = inferName(item);
          const text = item.textContent?.trim() || '';
          const normalizedName = normalizeName(name);
          return {
            index,
            name,
            normalizedName,
            exactNameMatched: normalizedName === expectedName,
            text,
          };
        });
        const matchedItems = resultItems.filter(item => item.exactNameMatched);

        return {
          count,
          paginationText,
          resultListText,
          noItemsText,
          resultItems,
          matchedCount: matchedItems.length,
        };
      }, '.vv-search-results-platform-pagination span', '.vv-compact-result-list', sharedName);
    }

    async function selectSharedSearchResult(sharedName) {
      const result = await page.evaluate((resultListSelector, expectedName) => {
        const normalizeName = value => String(value || '')
          .replace(/\s*\(v[\d.]+\)\s*$/i, '')
          .replace(/\s+/g, ' ')
          .trim();
        const inferName = item => {
          const explicitName = item.querySelector('.vv-search-result-name, .docName, .docNameLink, .vv_doc_title_name')?.textContent?.trim() || '';
          if (explicitName) return explicitName;
          const text = item.textContent?.trim() || '';
          const versionMatch = text.match(/^(.+?\s*\(v[\d.]+\))/i);
          if (versionMatch) return versionMatch[1].trim();
          const docNumberMatch = text.match(/^(.*?)(VV-\d+|MCS-\d+|ドラフト|Draft|Approved|承認済)/i);
          return (docNumberMatch?.[1] || text).trim();
        };
        const resultList = document.querySelector(resultListSelector);
        const directRows = Array.from(resultList?.children || []).filter(child =>
          child.matches('li, .binderDocRow, .vv-doc-compact-item, .vv_veeva_document')
        );
        const items = directRows.length > 0
          ? directRows
          : Array.from(resultList?.querySelectorAll('li, .binderDocRow, .vv-doc-compact-item, .vv_veeva_document') || []);
        const rows = items.length > 0 ? items : resultList ? [resultList] : [];
        const targetItem = rows.find(item => {
          const name = inferName(item);
          return normalizeName(name) === expectedName;
        });
        if (!targetItem) {
          return { selected: false, reason: 'target-not-found' };
        }

        const selectionArea = targetItem.querySelector('.vv-doc-compact-item-selection');
        const input = targetItem.querySelector('input[type=radio]');
        const clickable = input?.closest('[data-corgix-internal="RADIO"]') ||
          input?.nextElementSibling ||
          selectionArea ||
          targetItem;

        clickable.scrollIntoView({ block: 'center', inline: 'center' });
        clickable.click();
        if (input) {
          input.click();
          if (!input.checked) {
            input.checked = true;
            input.dispatchEvent(new Event('input', { bubbles: true }));
            input.dispatchEvent(new Event('change', { bubbles: true }));
          }
        }

        return {
          selected: true,
          checked: input ? input.checked : null,
          text: targetItem.textContent?.trim() || '',
        };
      }, '.vv-compact-result-list', sharedName);

      if (!result.selected) {
        throw new Error(`Shared Resource「${sharedName}」の検索結果行を選択できませんでした。`);
      }

      try {
        await page.waitForFunction((resultListSelector, expectedName) => {
          const normalizeName = value => String(value || '')
            .replace(/\s*\(v[\d.]+\)\s*$/i, '')
            .replace(/\s+/g, ' ')
            .trim();
          const inferName = item => {
            const explicitName = item.querySelector('.vv-search-result-name, .docName, .docNameLink, .vv_doc_title_name')?.textContent?.trim() || '';
            if (explicitName) return explicitName;
            const text = item.textContent?.trim() || '';
            const versionMatch = text.match(/^(.+?\s*\(v[\d.]+\))/i);
            if (versionMatch) return versionMatch[1].trim();
            const docNumberMatch = text.match(/^(.*?)(VV-\d+|MCS-\d+|ドラフト|Draft|Approved|承認済)/i);
            return (docNumberMatch?.[1] || text).trim();
          };
          const resultList = document.querySelector(resultListSelector);
          const directRows = Array.from(resultList?.children || []).filter(child =>
            child.matches('li, .binderDocRow, .vv-doc-compact-item, .vv_veeva_document')
          );
          const items = directRows.length > 0
            ? directRows
            : Array.from(resultList?.querySelectorAll('li, .binderDocRow, .vv-doc-compact-item, .vv_veeva_document') || []);
          const rows = items.length > 0 ? items : resultList ? [resultList] : [];
          const targetItem = rows.find(item => {
            const name = inferName(item);
            return normalizeName(name) === expectedName;
          });
          if (!targetItem) return false;
          const input = targetItem.querySelector('input[type=radio]');
          const checkedByInput = input?.checked === true;
          const checkedByAria = targetItem.querySelector('[aria-checked="true"], [data-checked="true"]');
          return checkedByInput || !!checkedByAria;
        }, { timeout: SHORT_WAIT }, '.vv-compact-result-list', sharedName);
      } catch (e) {
        console.log(`Shared Resource「${sharedName}」のラジオ選択状態を画面上で確認できませんでしたが、選択クリックは実行済みです。`);
      }
    }

    async function addSharedResource(sharedName) {
      let lastSummary = null;

      for (let attempt = 1; attempt <= 3; attempt++) {
        const searchInputHandle = await getSharedResourceSearchInputHandle();
        await clearAndTypeHandle(searchInputHandle, sharedName, { delay: attempt === 1 ? 120 : 200 }, 'Shared Resource検索欄');

        const inputValue = await searchInputHandle.evaluate(input => input.value?.trim() || '');

        if (inputValue !== sharedName) {
          console.log(`Shared検索ボックスへの入力が正しくありません。再入力します。(${attempt}/3) 入力値: "${inputValue}"`);
          if (attempt === 3) {
            throw new Error(`Shared検索ボックスの入力値が完全一致しません。期待値: "${sharedName}", 入力値: "${inputValue}"`);
          }
          continue;
        }

        const previousPaginationText = await page.$eval('.vv-search-results-platform-pagination span', element => element.textContent?.trim() || '').catch(() => '');
        console.log(`${sharedName}でShared Resourceを検索します。(${attempt}/3)`);

        const searchTrigger = await triggerSearchFromInput(searchInputHandle);
        console.log(`Shared Resource検索を実行しました: ${searchTrigger.clicked ? 'button' : 'enter'} ${searchTrigger.label || ''}`);

        try {
          await page.waitForFunction((paginationSelector, resultListSelector, expectedName, beforeText) => {
            const normalizeName = value => String(value || '')
              .replace(/\s*\(v[\d.]+\)\s*$/i, '')
              .replace(/\s+/g, ' ')
              .trim();
            const inferName = item => {
              const explicitName = item.querySelector('.vv-search-result-name, .docName, .docNameLink, .vv_doc_title_name')?.textContent?.trim() || '';
              if (explicitName) return explicitName;
              const text = item.textContent?.trim() || '';
              const versionMatch = text.match(/^(.+?\s*\(v[\d.]+\))/i);
              if (versionMatch) return versionMatch[1].trim();
              const docNumberMatch = text.match(/^(.*?)(VV-\d+|MCS-\d+|ドラフト|Draft|Approved|承認済)/i);
              return (docNumberMatch?.[1] || text).trim();
            };
            const getRows = resultList => {
              if (!resultList) return [];
              const directRows = Array.from(resultList.children).filter(child =>
                child.matches('li, .binderDocRow, .vv-doc-compact-item, .vv_veeva_document')
              );
              if (directRows.length > 0) return directRows;
              const nestedRows = Array.from(resultList.querySelectorAll('li, .binderDocRow, .vv-doc-compact-item, .vv_veeva_document'));
              return nestedRows.length > 0 ? nestedRows : [resultList];
            };
            const paginationText = document.querySelector(paginationSelector)?.textContent?.trim() || '';
            const resultList = document.querySelector(resultListSelector);
            const noItemsText = document.querySelector('.vv-compact-result-no-items')?.textContent?.trim() || '';
            const exactResultExists = getRows(resultList).some(item => {
              const name = inferName(item);
              return normalizeName(name) === expectedName;
            });
            const normalizedPaginationText = paginationText.replace(/[〜～]/g, '~');
            const hasCount = /of\s+[\d,]+/i.test(paginationText) ||
              /\/\s*[\d,]+\s*$/.test(normalizedPaginationText);
            const hasNoItems = /no items found|no results found|no documents found/i.test(noItemsText);
            return (hasCount && (exactResultExists || paginationText !== beforeText)) || hasNoItems;
          }, { timeout: NORMAL_WAIT }, '.vv-search-results-platform-pagination span', '.vv-compact-result-list', sharedName, previousPaginationText);
        } catch (e) {
          console.log(`Shared Resource検索結果の読み込みを待ちきれませんでした。状態を確認してリトライします。(${attempt}/3)`);
        }

        lastSummary = await getSharedSearchSummary(sharedName);
        if (lastSummary.count === 1 && lastSummary.matchedCount === 1) {
          const matchedName = lastSummary.resultItems.find(item => item.exactNameMatched)?.name || sharedName;
          console.log(`${sharedName}のShared Resource検索結果が1件で、名前も完全一致しました: ${matchedName}`);
          console.log(`${sharedName}のShared Resourceのラジオボタンを選択します。`);
          await selectSharedSearchResult(sharedName);
          const saveButtonSelector = 'button[title="Save"][data-corgix-internal="BUTTON"], button[title="Save"], button[title="保存"]';
          await clickWhenReady(saveButtonSelector, LONG_WAIT);
          await Promise.race([
            page.waitForFunction(element => {
              if (!element.isConnected) return true;
              const style = window.getComputedStyle(element);
              const rect = element.getBoundingClientRect();
              return rect.width === 0 ||
                rect.height === 0 ||
                style.visibility === 'hidden' ||
                style.display === 'none';
            }, { timeout: 15000 }, searchInputHandle),
            waitForOptionalSelector('.vv-compact-result-list', 15000, { hidden: true }),
          ]);
          console.log(`${sharedName}のShared Resourceを保存しました。`);
          return lastSummary;
        }

        if (lastSummary.count !== null) {
          const names = lastSummary.resultItems.map(item => item.name || item.text).join(' / ') || '(なし)';
          console.log(`${sharedName}のShared Resource検索結果が期待と違います。検索をリトライします。(${attempt}/3) 件数: ${lastSummary.count}, 一致件数: ${lastSummary.matchedCount}`);
          console.log("候補:", names);
          if (attempt === 3) {
            throw new Error(`${sharedName}のShared Resource検索結果が1件ではありません。件数: ${lastSummary.count}, 一致件数: ${lastSummary.matchedCount}, 候補: ${names}`);
          }
          continue;
        }

        console.log(`Shared Resource検索結果の件数をまだ判定できません。検索をリトライします。(${attempt}/3)`);
        console.log("pagination:", lastSummary.paginationText || "(なし)");
        console.log("results:", lastSummary.resultListText ? lastSummary.resultListText.slice(0, 200) : "(なし)");
      }

      throw new Error(`${sharedName}のShared Resource検索結果の件数を判定できませんでした。pagination: ${lastSummary?.paginationText || '(なし)'}`);
    }

    async function changeDocumentStatusToStaged() {
      const statusButtonSelector = '.vv_docstatus_wrapper .documentLifecycleStateBadgeContainer button[data-corgix-internal="BUTTON"], .vv_docstatus_wrapper .vv-picker-badge button';
      const stagedMenuItemSelector = '.vv-picker-badge-menu li[data-value="dynamicAction:LifecycleUserAction3"], .vv-picker-badge-menu [role="option"], [data-corgix-internal="MENU-ITEM"]';
      const dialogSelector = '.ui-dialog, [role="dialog"], [data-corgix-internal="DIALOG"], .vv-dialog, .vv_modal, .modal';

      const visibleTextState = async () => page.evaluate((buttonSelector, menuSelector, modalSelector) => ({
        statusButtons: Array.from(document.querySelectorAll(buttonSelector)).map(button => ({
          text: (button.textContent || '').trim(),
          title: button.getAttribute('title') || '',
          expanded: button.getAttribute('aria-expanded') || '',
        })),
        menuItems: Array.from(document.querySelectorAll(menuSelector)).map(item => ({
          text: (item.textContent || '').trim(),
          value: item.getAttribute('data-value') || '',
        })),
        dialogs: Array.from(document.querySelectorAll(modalSelector)).map(dialog => ({
          text: (dialog.textContent || '').replace(/\s+/g, ' ').trim().slice(0, 500),
        })),
      }), statusButtonSelector, stagedMenuItemSelector, dialogSelector).catch(err => ({ error: String(err) }));

      const statusIsStaged = async () => page.evaluate(selector => {
        return Array.from(document.querySelectorAll(selector)).some(button => {
          const label = [
            button.getAttribute('title') || '',
            button.textContent || '',
            button.getAttribute('aria-label') || '',
          ].join(' ');
          return /staged/i.test(label);
        });
      }, statusButtonSelector).catch(() => false);

      const clickStatusButton = async () => {
        await waitForClickable(statusButtonSelector, LONG_WAIT);
        const point = await page.evaluate(selector => {
          const visible = element => {
            if (!element) return false;
            const style = window.getComputedStyle(element);
            const rect = element.getBoundingClientRect();
            return rect.width > 0 &&
              rect.height > 0 &&
              style.visibility !== 'hidden' &&
              style.display !== 'none';
          };
          const buttons = Array.from(document.querySelectorAll(selector));
          const target = buttons.find(button => {
            const label = [
              button.getAttribute('title') || '',
              button.textContent || '',
            ].join(' ');
            return visible(button) && /draft/i.test(label);
          }) || buttons.find(visible);
          if (!target) return { ok: false, reason: 'status button not found' };
          target.scrollIntoView({ block: 'center', inline: 'center' });
          const rect = target.getBoundingClientRect();
          return {
            ok: true,
            x: rect.left + rect.width / 2,
            y: rect.top + rect.height / 2,
            label: [target.getAttribute('title') || '', target.textContent || ''].join(' ').trim(),
          };
        }, statusButtonSelector);
        if (!point.ok) throw new Error(`ドキュメントステータスボタンが見つかりません: ${JSON.stringify(point)}`);
        console.log(`ドキュメントステータスのDraftボタンを押します: ${JSON.stringify(point)}`);
        await page.mouse.move(point.x, point.y);
        await page.mouse.down();
        await sleep(80);
        await page.mouse.up();
      };

      const clickStagedMenuItem = async () => {
        await page.waitForFunction(selector => {
          return Array.from(document.querySelectorAll(selector)).some(element => {
            const style = window.getComputedStyle(element);
            const rect = element.getBoundingClientRect();
            return rect.width > 0 &&
              rect.height > 0 &&
              style.visibility !== 'hidden' &&
              style.display !== 'none' &&
              (element.getAttribute('data-value') === 'dynamicAction:LifecycleUserAction3' ||
                (element.textContent || '').trim().toUpperCase() === 'STAGED');
          });
        }, { timeout: NORMAL_WAIT }, stagedMenuItemSelector);

        const point = await page.evaluate(selector => {
          const visible = element => {
            if (!element) return false;
            const style = window.getComputedStyle(element);
            const rect = element.getBoundingClientRect();
            return rect.width > 0 &&
              rect.height > 0 &&
              style.visibility !== 'hidden' &&
              style.display !== 'none';
          };
          const items = Array.from(document.querySelectorAll(selector));
          const target = items.find(element => visible(element) &&
            (element.getAttribute('data-value') === 'dynamicAction:LifecycleUserAction3' ||
              (element.textContent || '').trim().toUpperCase() === 'STAGED'));
          if (!target) return { ok: false, reason: 'staged menu item not found' };
          target.scrollIntoView({ block: 'center', inline: 'center' });
          const rect = target.getBoundingClientRect();
          const x = rect.left + rect.width / 2;
          const y = rect.top + rect.height / 2;
          const eventOptions = { bubbles: true, cancelable: true, view: window, clientX: x, clientY: y, button: 0 };
          ['mouseover', 'mouseenter', 'mousemove'].forEach(type => target.dispatchEvent(new MouseEvent(type, eventOptions)));
          return {
            ok: true,
            x,
            y,
            text: (target.textContent || '').trim(),
            tag: target.tagName,
            className: String(target.className || ''),
            elementAtPointText: (document.elementFromPoint(x, y)?.textContent || '').trim(),
          };
        }, stagedMenuItemSelector);
        if (!point.ok) throw new Error(`Stagedメニュー項目が見つかりません: ${JSON.stringify(point)}`);
        console.log(`ステータスメニューのStagedを押します: ${JSON.stringify(point)}`);
        await page.mouse.move(point.x, point.y);
        await page.mouse.down();
        await sleep(80);
        await page.mouse.up();
        await sleep(800);
      };

      const hasStatusDialog = async () => page.evaluate(selector => {
        const visible = element => {
          if (!element) return false;
          const style = window.getComputedStyle(element);
          const rect = element.getBoundingClientRect();
          return rect.width > 0 &&
            rect.height > 0 &&
            style.visibility !== 'hidden' &&
            style.display !== 'none' &&
            style.visibility !== 'hidden' &&
            style.display !== 'none';
        };
        return Array.from(document.querySelectorAll(selector)).some(dialog => {
          if (!visible(dialog)) return false;
          const text = [
            dialog.querySelector('.ui-dialog-title')?.textContent || '',
            dialog.textContent || '',
          ].join(' ');
          const hasYes = Array.from(dialog.querySelectorAll('button, a, .save, [role="button"]')).some(button => {
            const label = button.textContent || button.getAttribute('title') || button.getAttribute('aria-label') || '';
            return visible(button) && /yes|はい|ok/i.test(label);
          });
          return /Change Document Status|Change State to Staged|Staged/i.test(text) && hasYes;
        });
      }, dialogSelector).catch(() => false);

      for (let attempt = 1; attempt <= 3; attempt += 1) {
        if (await statusIsStaged()) {
          console.log("ドキュメントステータスは既にStagedです。");
          return;
        }
        await clickStatusButton();
        await clickStagedMenuItem();
        if (await hasStatusDialog() || await statusIsStaged()) break;
        console.log(`Stagedクリック後に確認ダイアログが出ませんでした。再試行します。(${attempt}/3)`);
        await page.keyboard.press('Escape').catch(() => null);
        await sleep(1000);
        if (attempt === 3) {
          throw new Error(`Stagedクリック後に確認ダイアログを表示できませんでした。状態: ${JSON.stringify(await visibleTextState())}`);
        }
      }

      if (await statusIsStaged()) {
        console.log("ドキュメントステータスをStagedへ変更しました。");
        return;
      }

      console.log("Change Document StatusダイアログのYesを押します。");
      const yesClickPoint = await page.evaluate(selector => {
        const visible = element => {
          if (!element) return false;
          const style = window.getComputedStyle(element);
          const rect = element.getBoundingClientRect();
          return rect.width > 0 &&
            rect.height > 0 &&
            style.visibility !== 'hidden' &&
            style.display !== 'none';
        };
        const dialogs = Array.from(document.querySelectorAll(selector));
        const dialog = dialogs.find(candidate => {
          if (!visible(candidate)) return false;
          const text = [
            candidate.querySelector('.ui-dialog-title')?.textContent || '',
            candidate.textContent || '',
          ].join(' ');
          return /Change Document Status|Change State to Staged|Staged/i.test(text);
        });
        const yesButton = Array.from(dialog?.querySelectorAll('button, a, .save, [role="button"]') || []).find(button => {
          const label = button.textContent || button.getAttribute('title') || button.getAttribute('aria-label') || '';
          return visible(button) && /yes|はい|ok/i.test(label);
        });
        if (!yesButton) return { ok: false, reason: 'yes button not found', dialogText: (dialog?.textContent || '').trim().slice(0, 500) };
        yesButton.scrollIntoView({ block: 'center', inline: 'center' });
        const rect = yesButton.getBoundingClientRect();
        return {
          ok: true,
          x: rect.left + rect.width / 2,
          y: rect.top + rect.height / 2,
          text: (yesButton.textContent || '').trim(),
        };
      }, dialogSelector);
      if (!yesClickPoint.ok) {
        throw new Error(`Change Document StatusダイアログのYesボタンが見つかりません: ${JSON.stringify(yesClickPoint)}`);
      }
      console.log(`Change Document StatusダイアログのYesを押します: ${JSON.stringify(yesClickPoint)}`);
      const statusNavigation = page.waitForNavigation({ waitUntil: ['load', 'networkidle2'], timeout: 120000 }).catch(() => null);
      await page.mouse.move(yesClickPoint.x, yesClickPoint.y);
      await page.mouse.down();
      await sleep(80);
      await page.mouse.up();

      await Promise.race([
        page.waitForFunction(selector => {
          return !Array.from(document.querySelectorAll(selector)).some(dialog => {
            const style = window.getComputedStyle(dialog);
            const rect = dialog.getBoundingClientRect();
            const text = dialog.textContent || '';
            return rect.width > 0 &&
              rect.height > 0 &&
              style.visibility !== 'hidden' &&
              style.display !== 'none' &&
              /Change Document Status|Change State to Staged|Staged/i.test(text);
          });
        }, { timeout: NORMAL_WAIT }, dialogSelector),
        page.waitForFunction(selector => {
          return Array.from(document.querySelectorAll(selector)).some(button => {
            const label = [
              button.getAttribute('title') || '',
              button.textContent || '',
            ].join(' ');
            return /staged/i.test(label);
          });
        }, { timeout: NORMAL_WAIT }, statusButtonSelector),
      ]).catch(() => null);

      await page.waitForFunction(selector => {
        return Array.from(document.querySelectorAll(selector)).some(button => {
          const label = [
            button.getAttribute('title') || '',
            button.textContent || '',
            button.getAttribute('aria-label') || '',
          ].join(' ');
          return /staged/i.test(label);
        });
      }, { timeout: LONG_WAIT }, statusButtonSelector).catch(() => null);

      if (!(await statusIsStaged())) {
        const beforeReloadState = await visibleTextState();
        console.log(`Staged反映を画面上で確認できないためページを再読み込みして確認します。状態: ${JSON.stringify(beforeReloadState)}`);
        await statusNavigation;
        const currentUrl = page.url();
        await page.goto(currentUrl, DCL).catch(() => null);
        await waitForOptionalSelector(statusButtonSelector, LONG_WAIT, { visible: true });
        await page.waitForFunction(selector => {
          return Array.from(document.querySelectorAll(selector)).some(button => {
            const label = [
              button.getAttribute('title') || '',
              button.textContent || '',
              button.getAttribute('aria-label') || '',
            ].join(' ');
            return /staged/i.test(label);
          });
        }, { timeout: NORMAL_WAIT }, statusButtonSelector).catch(() => null);
      }
      if (!(await statusIsStaged())) {
        throw new Error(`ドキュメントステータスをStagedへ変更できませんでした。状態: ${JSON.stringify(await visibleTextState())}`);
      }
      console.log("ドキュメントステータスをStagedへ変更しました。");
    }

    async function openSharedResourceAddDialog() {
      const sectionHeaderSelector = '.section-related_shared_resource__pm';
      const addLinkSelector = '.add-related_shared_resource__pm';
      const addButtonSelector = '.add-related_shared_resource__pm button, .add-related_shared_resource__pm';

      await page.waitForSelector(sectionHeaderSelector, { timeout: LONG_WAIT });
      await page.$eval(sectionHeaderSelector, element => {
        element.scrollIntoView({ block: 'center', inline: 'center' });
      });

      for (let attempt = 1; attempt <= 3; attempt++) {
        console.log(`Related Shared ResourceのAddボタンを表示します。(${attempt}/3)`);
        const header = await page.$(sectionHeaderSelector);
        if (!header) break;

        await header.hover();
        const box = await header.boundingBox();
        if (box) {
          await page.mouse.move(box.x + box.width - 24, box.y + box.height / 2);
        }

        let addVisible = await waitForOptionalSelector(addButtonSelector, SHORT_WAIT, { visible: true });
        if (!addVisible) {
          await page.$eval(sectionHeaderSelector, element => {
            ['mouseover', 'mouseenter', 'mousemove'].forEach(type => {
              element.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, view: window }));
            });
            const addLink = element.querySelector('.add-related_shared_resource__pm') ||
              document.querySelector('.add-related_shared_resource__pm');
            if (addLink) {
              addLink.style.display = '';
            }
          });
          addVisible = await waitForOptionalSelector(addButtonSelector, SHORT_WAIT, { visible: true });
        }

        if (addVisible) {
          await clickWhenReady(addButtonSelector, NORMAL_WAIT);
        } else {
          const clicked = await page.evaluate(selector => {
            const element = document.querySelector(selector);
            if (!element) return false;
            element.click();
            return true;
          }, addLinkSelector);
          if (!clicked) {
            continue;
          }
        }

        try {
          await getSharedResourceSearchInputHandle();
          return;
        } catch (e) {
          console.log(`Shared Resource追加ダイアログがまだ開いていません。再試行します。(${attempt}/3)`);
        }
      }

      throw new Error("Related Shared ResourceのAddダイアログを開けませんでした。");
    }

    async function openBinderCreateWizard(baseUrl) {
      for (let attempt = 1; attempt <= 3; attempt++) {
        console.log(`Binder作成画面へ移動します。(${attempt}/3)`);
        await page.goto(`${baseUrl}/ui/#inbox/binder`, DCL);

        const createButton = await waitForOptionalSelector(BINOCULARS_SUBMIT_SELECTOR, LONG_WAIT, { visible: true });
        if (!createButton) {
          console.log(red + `Binder作成ボタンが表示されませんでした。再読み込みします。(${attempt}/3)` + reset);
          continue;
        }

        console.log("Binder作成ボタンを押します。");
        await clickWhenReady(BINOCULARS_SUBMIT_SELECTOR, LONG_WAIT);

        const typeSelect = await waitForOptionalSelector(TPYE_SUBMIT_SELECTOR, 30000, { visible: true });
        if (typeSelect) {
          console.log("ドキュメントタイプ選択欄が表示されました。");
          return;
        }

        console.log(red + `ドキュメントタイプ選択欄が表示されませんでした。再試行します。(${attempt}/3)` + reset);
      }

      throw new Error("Binder作成ウィザードを開けませんでした。");
    }


    async function ensureSectionFieldVisible(sectionTitle, fieldSelector) {
      if (!sectionTitle) {
        const targetIndex = await getVisibleElementIndex(fieldSelector, LONG_WAIT);
        await page.evaluate((selector, index) => {
          const target = Array.from(document.querySelectorAll(selector))[index];
          target?.scrollIntoView({ block: 'center', inline: 'center' });
        }, fieldSelector, targetIndex);
        return;
      }

      const sectionSelector = `h3[title="${sectionTitle}"]`;
      await page.waitForSelector(sectionSelector, { timeout: NORMAL_WAIT });
      await page.$eval(sectionSelector, el => el.scrollIntoView({ block: 'center', inline: 'center' }));
      const isVisibleNearSection = async () => page.evaluate((headingSelector, targetSelector) => {
        const isVisible = element => {
          if (!element) return false;
          const style = window.getComputedStyle(element);
          const rect = element.getBoundingClientRect();
          return rect.width > 0 &&
            rect.height > 0 &&
            !element.disabled &&
            element.getAttribute('aria-disabled') !== 'true' &&
            style.visibility !== 'hidden' &&
            style.display !== 'none';
        };
        const heading = Array.from(document.querySelectorAll(headingSelector)).find(isVisible) ||
          document.querySelector(headingSelector);
        if (!heading) return false;
        const headingRect = heading.getBoundingClientRect();
        const target = Array.from(document.querySelectorAll(targetSelector))
          .map(element => {
            const rect = element.getBoundingClientRect();
            const near = rect.top >= headingRect.top - 80 && rect.top <= headingRect.bottom + 900;
            const distance = Math.abs(rect.top - headingRect.bottom) + Math.abs(rect.left - headingRect.left);
            return { element, visible: isVisible(element), near, distance };
          })
          .filter(item => item.visible && item.near)
          .sort((a, b) => a.distance - b.distance)[0]?.element;
        if (!target) return false;
        target.scrollIntoView({ block: 'center', inline: 'center' });
        return true;
      }, sectionSelector, fieldSelector);

      if (await isVisibleNearSection()) return;

      await clickWhenReady(sectionSelector, LONG_WAIT);
      await page.waitForFunction(
        (headingSelector, targetSelector) => {
          const isVisible = element => {
            if (!element) return false;
            const style = window.getComputedStyle(element);
            const rect = element.getBoundingClientRect();
            return rect.width > 0 &&
              rect.height > 0 &&
              !element.disabled &&
              element.getAttribute('aria-disabled') !== 'true' &&
              style.visibility !== 'hidden' &&
              style.display !== 'none';
          };
          const heading = Array.from(document.querySelectorAll(headingSelector)).find(isVisible) ||
            document.querySelector(headingSelector);
          if (!heading) return false;
          const headingRect = heading.getBoundingClientRect();
          return Array.from(document.querySelectorAll(targetSelector)).some(element => {
            const rect = element.getBoundingClientRect();
            return isVisible(element) &&
              rect.top >= headingRect.top - 80 &&
              rect.top <= headingRect.bottom + 900;
          });
        },
        { timeout: LONG_WAIT },
        sectionSelector,
        fieldSelector
      );
      await isVisibleNearSection();
    }

    async function ensurePresentationIdFieldVisible() {
      try {
        await ensureSectionFieldVisible('Multichannel Properties', presentationId_SELECTOR);
        return;
      } catch (e) {
        console.log(`Multichannel PropertiesセクションからPresentation ID欄を表示できませんでした。ラベルから再試行します: ${e.message}`);
      }

      const opened = await page.evaluate(selector => {
        const field = document.querySelector(selector);
        const label = document.querySelector('[attrkey="crmPresentationId_b"], .docInfoLabel-crmPresentationId_b');
        const anchor = field || label;
        const section = anchor?.closest?.('.fullDocSectionContainer');
        const header = section?.querySelector?.('h3.sectionAccordionHeader, h3[title], .sectionAccordionHeader');
        if (!header) {
          return {
            ok: false,
            reason: 'section header not found',
            hasField: !!field,
            hasLabel: !!label,
          };
        }
        header.scrollIntoView({ block: 'center', inline: 'center' });
        const body = section.querySelector('.doc_info_section_body, .ui-accordion-content');
        const style = body ? window.getComputedStyle(body) : null;
        const expanded = body && style.display !== 'none' && style.visibility !== 'hidden';
        if (!expanded) {
          header.click();
        }
        (field || label)?.scrollIntoView?.({ block: 'center', inline: 'center' });
        return {
          ok: true,
          title: header.getAttribute('title') || header.textContent?.trim() || '',
          hadExpandedBody: !!expanded,
        };
      }, presentationId_SELECTOR);
      console.log(`Presentation IDセクション表示状態: ${JSON.stringify(opened)}`);
      if (!opened.ok) {
        throw new Error(`Presentation ID欄のセクションを表示できませんでした: ${JSON.stringify(opened)}`);
      }
      await page.waitForFunction(selector => {
        const field = document.querySelector(selector);
        if (!field) return false;
        const style = window.getComputedStyle(field);
        const rect = field.getBoundingClientRect();
        return rect.width > 0 &&
          rect.height > 0 &&
          style.visibility !== 'hidden' &&
          style.display !== 'none';
      }, { timeout: LONG_WAIT }, presentationId_SELECTOR);
    }

    async function typeAndSelectMultipleMenuItems(inputSelector, targets, options = {}) {
      for (const target of targets) {
        await ensureSectionFieldVisible(options.sectionTitle || '', inputSelector);
        await typeAndSelectMenuItem(inputSelector, target, {
          label: options.label ? `${options.label}: ${target}` : target,
          timeout: options.timeout || NORMAL_WAIT,
          sectionTitle: options.sectionTitle || '',
          selectionMode: options.selectionMode || 'mouse',
        });
      }
    }

    function normalizeVeevaUrl(urlStr) {
      const url = new URL(urlStr);

      // #doc_info/7702/0/1
      const hash = url.hash.replace(/^#/, "");

      const parts = hash.split("/");

      // doc_info + docId だけ残す
      if (parts[0] === "doc_info" && parts.length >= 2) {
        url.hash = `doc_info/${parts[1]}`;
      }

      return url.toString();
    }

    async function gotoSlideFromBinderList(baseUrl, expectedName = '') {
      await page.waitForSelector('.vv_library_list .binderDocRow, .vv_library_list .vv_veeva_document', { timeout: LONG_WAIT });

      const collectBinderRows = async () => page.evaluate(targetName => {
        const normalizeName = value => String(value || '')
          .replace(/\s*\(v[\d.]+\)\s*$/i, '')
          .replace(/\s+/g, ' ')
          .trim();
        const normalizeToken = value => String(value || '').replace(/\s+/g, ' ').trim();
        const expected = normalizeName(targetName);

        return Array.from(document.querySelectorAll('.vv_library_list .binderDocRow, .vv_library_list .vv_veeva_document'))
          .map(row => {
            const docType = row.querySelector('.docType')?.textContent?.trim() || '';
            const name = row.querySelector('.docName')?.textContent?.trim() ||
              row.querySelector('.docNameLink')?.textContent?.trim() ||
              '';
            const link = row.querySelector('a.docNameLink[href*="#doc_info/"], a.docThumbnail[href*="#doc_info/"], a[href*="#doc_info/"]');
            const href = link?.getAttribute('href') || '';
            const hrefMatch = href.match(/#doc_info\/(\d+)/);
            const dockeyMatch = (row.getAttribute('dockey') || '').match(/^(\d+)-/);
            const docId = hrefMatch?.[1] || dockeyMatch?.[1] || '';
            const normalizedName = normalizeName(name);
            const rowText = normalizeToken(row.textContent || '');
            const typeText = normalizeToken(`${docType} ${row.getAttribute('class') || ''} ${row.getAttribute('data-doctype') || ''} ${rowText}`);
            const isSlide = /スライド|slide/i.test(typeText);

            return {
              docId,
              docType,
              href,
              name,
              normalizedName,
              rowText: rowText.slice(0, 500),
              isSlide,
              nameMatched: !!expected && normalizedName === expected,
            };
          })
          .filter(item => item.docId);
      }, expectedName);

      await page.waitForFunction(targetName => {
        const normalizeName = value => String(value || '')
          .replace(/\s*\(v[\d.]+\)\s*$/i, '')
          .replace(/\s+/g, ' ')
          .trim();
        const expected = normalizeName(targetName);
        return Array.from(document.querySelectorAll('.vv_library_list .binderDocRow, .vv_library_list .vv_veeva_document'))
          .some(row => {
            const docType = row.querySelector('.docType')?.textContent?.trim() || '';
            const name = row.querySelector('.docName')?.textContent?.trim() ||
              row.querySelector('.docNameLink')?.textContent?.trim() ||
              '';
            const href = row.querySelector('a.docNameLink[href*="#doc_info/"], a.docThumbnail[href*="#doc_info/"], a[href*="#doc_info/"]')?.getAttribute('href') || '';
            const hrefMatch = href.match(/#doc_info\/(\d+)/);
            const dockeyMatch = (row.getAttribute('dockey') || '').match(/^(\d+)-/);
            const docId = hrefMatch?.[1] || dockeyMatch?.[1] || '';
            const rowText = String(row.textContent || '').replace(/\s+/g, ' ').trim();
            const isSlide = /スライド|slide/i.test(`${docType} ${row.getAttribute('class') || ''} ${rowText}`);
            const nameMatched = !!expected && normalizeName(name) === expected;
            return !!docId && (isSlide || nameMatched);
          });
      }, { timeout: LONG_WAIT }, expectedName).catch(() => null);

      const allCandidates = await collectBinderRows();
      const slideCandidates = allCandidates.filter(item => item.isSlide || item.nameMatched);

      if (slideCandidates.length === 0) {
        const rowSummary = allCandidates.map(candidate => {
          return `${candidate.docId || '(idなし)'} / type=${candidate.docType || '(空)'} / name=${candidate.name || '(空)'} / text=${candidate.rowText || '(空)'}`;
        }).join(' || ');
        throw new Error(`Binder内にスライドのドキュメント行が見つかりませんでした。候補行: ${rowSummary || 'なし'}`);
      }

      const nameMatchedCandidates = slideCandidates.filter(candidate => candidate.nameMatched);
      let target = null;
      if (nameMatchedCandidates.length === 1) {
        target = nameMatchedCandidates[0];
      } else if (slideCandidates.length === 1) {
        target = slideCandidates[0];
      } else {
        const candidateText = slideCandidates.map(candidate => `${candidate.docId}: ${candidate.name}`).join(' / ');
        throw new Error(`Binder内のスライド候補が複数あり、対象を1件に絞れませんでした。候補: ${candidateText}`);
      }

      const slideUrl = `${baseUrl.replace(/\/$/, '')}/ui/#doc_info/${target.docId}`;
      console.log(`Binder内のスライドへ移動します: ${target.name} -> ${slideUrl}`);
      await page.goto(slideUrl, DCL);
      await waitForAnySelector([
        "li[data-target-key=doc_info_relationships__sys]",
        ".vv_docstatus_wrapper",
      ], LONG_WAIT, { visible: true });
      return slideUrl;
    }


    // 不要なリソースをブロックしてページ読み込みを高速化
    await page.setRequestInterception(true);
    page.on('request', (req) => {
      const type = req.resourceType();
      if (['image', 'font', 'media'].includes(type)) {
        req.abort();
      } else {
        req.continue();
      }
    });

    try {
      await page.setDefaultNavigationTimeout(120000);
      await page.setDefaultTimeout(90000); // 全体のデフォルトタイムアウトを90秒に

      let tUrl
      if (LOGIN_USER !== "Hayato.Seto@vv-agency.com") {
        tUrl = 'https://msd-promomats-ghh.veevavault.com';
      } else {
        tUrl = 'https://vvagency-arashimaru.veevavault.com';
      }
      // テスト用

      console.log(`Vaultへアクセスしています: ${tUrl}`);
      await page.goto(tUrl, DCL);
      await waitForAnySelector([LOGIN_USER_SELECTOR, "#search_main_box", ".vv_username"], LONG_WAIT);
      let u = await page.url();
      let loginForm = await page.$(LOGIN_USER_SELECTOR);

      if (u.match(/login/) || loginForm) {
        console.log(`Vaultへログインしています: ${LOGIN_USER}`);
        await page.type(LOGIN_USER_SELECTOR, LOGIN_USER);

        await Promise.all([

          page.waitForSelector(LOGIN_SUBMIT_SELECTOR),
          page.click(LOGIN_CONTINUE_SELECTOR),
        ]);

        await page.type(LOGIN_PASS_SELECTOR, LOGIN_PASS);

        await Promise.all([
          page.waitForNavigation({ waitUntil: ['load', 'networkidle2'], timeout: 120000 }).catch(() => null),
          page.click(LOGIN_SUBMIT_SELECTOR),
        ]);

        await waitForAnySelector(["#search_main_box", ".vv_username"], LONG_WAIT);
        console.log("Vaultログインが完了しました。");
      } else {
        console.log("Vaultログイン済みセッションを使用します。");
      }

      let lexNoThanks = await waitForOptionalSelector(".vv-callout-content-dismiss", 5000, { visible: true });

      if (lexNoThanks) {
        await lexNoThanks.click();
      }

      if (LOGIN_USER !== "Hayato.Seto@vv-agency.com") {

        await page.waitForSelector(".vv_username", { timeout: LONG_WAIT });
        let currentUser = await page.$(".vv_username");

        let currentUserValue = await (await currentUser.getProperty('textContent')).jsonValue();
        if (normalizeMenuText(currentUserValue) !== normalizeMenuText("Arashimaru Inc. Agency")) {
          console.log("代理アクセスユーザーへ切り替えています。");
          await clearAndType(ACCESSCONTROL_SELECTOR, ACCESSCONTROL);
          const delegateNavigation = page.waitForNavigation({ waitUntil: ['load', 'networkidle2'], timeout: 120000 }).catch(() => null);
          await waitAndSelectMenuItem('Arashimaru Inc. Agency arashimaru@msd.com', {
            label: '代理アクセスユーザー',
            inputSelector: ACCESSCONTROL_SELECTOR,
            selectionMode: 'mouse',
            waitForReady: false,
            stabilizeMs: 800,
          });
          await Promise.race([
            delegateNavigation,
            waitForDelegateUserSwitch("Arashimaru Inc. Agency", 15000).catch(() => null),
          ]);
          await waitForDelegateUserSwitch("Arashimaru Inc. Agency");
          await waitForAnySelector(["#search_main_box", ".vv_username"], LONG_WAIT);
          console.log("代理アクセスユーザーへの切り替えが完了しました。");
        }

        const duplicateSearchTargets = [
          { label: 'プレゼンテーションID', value: String(presentationId).trim() },
        ].filter((item, index, self) => item.value && self.findIndex(candidate => candidate.value === item.value) === index);

        console.log(`既存登録チェック対象: ${duplicateSearchTargets.map(item => `${item.label}="${item.value}"`).join(' / ')}`);
        for (const duplicateTarget of duplicateSearchTargets) {
          const searchName = duplicateTarget.value;
          console.log(`既存登録を厳格検索しています: ${duplicateTarget.label} "${searchName}"`);
          const searchSummary = await searchExistingDocument(searchName, tUrl);
          if (!searchSummary.verified || searchSummary.count === null) {
            const message = `${duplicateTarget.label} "${searchName}" の検索結果を厳格に判定できませんでした。既存登録の確認ができないため、処理を中断します。`;
            console.log(red + message + reset);
            if (searchSummary.error) console.log("error:", searchSummary.error);
            console.log("input:", searchSummary.inputValue || "(なし)");
            console.log("paginator:", searchSummary.paginatorText || "(なし)");
            console.log("grid:", searchSummary.gridText ? searchSummary.gridText.slice(0, 300) : "(なし)");
            if (searchSummary.resultRows?.length) {
              console.log("rows:", searchSummary.resultRows.join(" / ").slice(0, 500));
            }
            return ["作成失敗", page.url(), slideURL, message];
          }

          if (searchSummary.count === 0) {
            console.log(`${duplicateTarget.label} "${searchName}" で検索した結果、登録済みドキュメントは0件でした。`);
            continue;
          }

          const message = `${duplicateTarget.label} "${searchName}" で検索した結果、${searchSummary.count}件の登録済みドキュメントが見つかったため中断しました。`;
          console.log(red + message + reset);
          if (searchSummary.urls.length > 0) {
            searchSummary.urls.forEach(url => console.log(url));
            if (searchSummary.count > searchSummary.urls.length) {
              console.log("表示中のURLのみ出力しています。");
            }
          }
          if (searchSummary.resultRows?.length) {
            console.log("rows:", searchSummary.resultRows.join(" / ").slice(0, 500));
          }
          return ["作成失敗", page.url(), slideURL, message];
        }
        console.log("既存登録チェックが完了しました。登録処理を続行します。");

      }


      console.log("Engage PresentationのBinder作成画面を開いています。");
      await openBinderCreateWizard(tUrl);

      await page.select(TPYE_SUBMIT_SELECTOR, 'engagePresentation_b');
      await page.waitForFunction(
        selector => document.querySelector(selector)?.value === 'engagePresentation_b',
        { timeout: NORMAL_WAIT },
        TPYE_SUBMIT_SELECTOR
      );
      console.log("ドキュメントタイプ選択後のOKボタンを押します。");
      await clickWhenReady(OK_SUBMIT_SELECTOR, LONG_WAIT);
      try {
        await page.waitForSelector(OK_SUBMIT_SELECTOR, { hidden: true, timeout: 10000 });
      } catch (e) {
        console.log(red + "OK押下後もOKボタンが残っています。もう一度OKを押します。" + reset);
        await clickWhenReady(OK_SUBMIT_SELECTOR, LONG_WAIT);
        await page.waitForSelector(OK_SUBMIT_SELECTOR, { hidden: true, timeout: 30000 });
      }

      console.log("作成ウィザードのNextボタンを押します。");
      await clickWhenReady(NEXT_SUBMIT_SELECTOR, LONG_WAIT);
      const countryInput = await waitForOptionalSelector(COUNTRY_SELECTOR, 15000);
      if (!countryInput) {
        console.log(red + "Next押下後に国入力欄が表示されませんでした。もう一度Nextを押します。" + reset);
        await clickWhenReady(NEXT_SUBMIT_SELECTOR, LONG_WAIT);
        await page.waitForSelector(COUNTRY_SELECTOR, { timeout: LONG_WAIT });
      }

      console.log("Binder基本情報を入力しています。");
      await typeAndSelectMenuItem(COUNTRY_SELECTOR, COUNTRY);
      await clearAndType(NAME_SELECTOR, NAME);

      if (LOGIN_USER !== "Hayato.Seto@vv-agency.com") {
        console.log(`Productを選択しています: ${PRODUCT}`);
        await typeAndSelectMenuItem(PRODUCT_SELECTOR, PRODUCT, {
          label: 'Product',
          clearSelectedItems: true,
          requireSelectedText: true,
          selectionMode: 'mouse',
        });

        // CS環境
        console.log("Detail Groupと言語を設定しています。");
        await typeAndSelectMenuItem(DETAILGROUP_SELECTOR, DETAILGROUP);
        await typeAndSelectMenuItem(LANGUAGE_SELECTOR, LANGUAGE, {
          label: 'Binder Language',
        });
        await clearAndType(PRODUCTTEXT_SELECTOR, PRODUCT);
        await ensurePresentationIdFieldVisible();
        await clearAndTypeRequired(presentationId_SELECTOR, presentationId, 'Presentation ID');

      } else {

        console.log("嵐丸環境の製品情報を設定しています。");
        await typeAndSelectMultipleMenuItems(
          PRODUCT_SELECTOR_TEST,
          ["Cholecap"],
          { sectionTitle: "製品情報", label: "製品" }
        );
      }

      await selectRadioById('clmContent_bYES');
      await selectRadioById('crmHidden_bYES');



      console.log("Binderを保存しています。");
      await waitForClickable(SAVE_SUBMIT_SELECTOR);
      await page.click(SAVE_SUBMIT_SELECTOR);


      console.log("Binderのレンディション生成完了を待っています。");
      await waitForRenditionToFinish();

      // アップテスト
      const binderURL = normalizeVeevaUrl(page.url())
      console.log("Binderが作成されました:" + binderURL);


      console.log("Binder編集画面を開いてZIPを追加しています。");
      const createButton2 = await waitForOptionalSelector(".vv-edit-binder", LONG_WAIT, { visible: true });
      await clickWhenReady(".vv-edit-binder", LONG_WAIT);

      await openBinderAddFilesMenu();

      try {
        await page.waitForSelector('#inboxFileChooserHTML5', { timeout: 10000 });
      } catch (e) {
        console.log(red + "Upload File押下後もファイル選択欄が表示されません。もう一度Addメニューを開きます。" + reset);
        await openBinderAddFilesMenu();
        await page.waitForSelector('#inboxFileChooserHTML5', { timeout: 30000 });
      }



      await page.waitForSelector('#inboxFileChooserHTML5', { timeout: 60000 });
      const inputUploadHandle = await page.$('#inboxFileChooserHTML5');
      console.log(`ZIPをアップロードしています: ${zIPfolder}`);
      await inputUploadHandle.uploadFile(zIPfolder);

      console.log("スライド作成ウィザードを開いています。");
      console.log("ZIPアップロード完了とスライド作成ボタンの有効化を待っています。");
      await clickActiveSaveButton(BINOCULARS_SUBMIT_SELECTOR, "スライド作成", LONG_WAIT, {
        waitForGlobalBusy: true,
      });

      const typeSelect = await waitForOptionalSelector(TPYE_SUBMIT_SELECTOR, 30000, { visible: true });




      await page.select(TPYE_SUBMIT_SELECTOR, 'slide_b');
      await page.waitForFunction(
        selector => document.querySelector(selector)?.value === 'slide_b',
        { timeout: NORMAL_WAIT },
        TPYE_SUBMIT_SELECTOR
      );
      console.log("ドキュメントタイプ選択後のOKボタンを押します。");
      await clickWhenReady(OK_SUBMIT_SELECTOR, LONG_WAIT);
      try {
        await page.waitForSelector(OK_SUBMIT_SELECTOR, { hidden: true, timeout: 10000 });
      } catch (e) {
        console.log(red + "OK押下後もOKボタンが残っています。もう一度OKを押します。" + reset);
        await clickWhenReady(OK_SUBMIT_SELECTOR, LONG_WAIT);
        await page.waitForSelector(OK_SUBMIT_SELECTOR, { hidden: true, timeout: 30000 });
      }




      console.log("作成ウィザードのNextボタンを押します。");
      await clickWhenReady(NEXT_SUBMIT_SELECTOR, LONG_WAIT);
      const countryInput_s = await waitForOptionalSelector(COUNTRY_SELECTOR, 15000);
      if (!countryInput_s) {
        console.log(red + "Next押下後に国入力欄が表示されませんでした。もう一度Nextを押します。" + reset);
        await clickWhenReady(NEXT_SUBMIT_SELECTOR, LONG_WAIT);
        await page.waitForSelector(COUNTRY_SELECTOR, { timeout: LONG_WAIT });
      }



      console.log("スライド基本情報を入力しています。");
      await clearAndType(NAME_SELECTOR, NAME);

      if (LOGIN_USER !== "Hayato.Seto@vv-agency.com") {
        console.log("スライドの言語とDisable Actionsを設定しています。");
        await typeAndSelectMenuItem(LANGUAGE_SELECTOR, LANGUAGE, {
          label: 'スライド Language',
        });
        const DISABLE_ACTIONS_SELECTOR = "div[name=crmDisableActions_b] > .vv_pill_container > input";

        await typeAndSelectMultipleMenuItems(
          DISABLE_ACTIONS_SELECTOR,
          ["Pinch to Exit", "Rotation Lock"],
          { sectionTitle: "CLM Properties", label: "Disable Actions" }
        );
      } else {

        const DISABLE_ACTIONS_SELECTOR = "div[name=crmDisableActions_b] > .vv_pill_container > input";

        await typeAndSelectMultipleMenuItems(
          DISABLE_ACTIONS_SELECTOR,
          ["終了用ピンチアクション", "Rotation Lock"],
          { sectionTitle: "CLM のプロパティ", label: "アクションの無効化" }
        );
      }




      console.log("スライドを保存しています。");
      await clickActiveSaveButton(SAVE_SUBMIT_SELECTOR, "スライドのSave", LONG_WAIT, {
        beforeUrl: page.url(),
        confirmAfterClick: true,
      });

      console.log("スライドのレンディション生成完了を待っています。");
      await waitForRenditionToFinish();





      // 紐付けテスト

      await page.goto(binderURL, DCL);
      console.log("Binder内のスライドを探しています。");
      slideURL = await gotoSlideFromBinderList(tUrl, NAME);


      const createButton5 = await waitForOptionalSelector("li[data-target-key=doc_info_relationships__sys]", LONG_WAIT, { visible: true });
      await clickWhenReady("li[data-target-key=doc_info_relationships__sys]", LONG_WAIT);
      await openSharedResourceAddDialog();

      console.log("Shared Resourceを紐付けています: MSD_ONC_TOOL_SHARED");
      await addSharedResource("MSD_ONC_TOOL_SHARED");

      if (LOGIN_USER !== "Hayato.Seto@vv-agency.com") {
        console.log("スライドのステータスをStagedへ変更しています。");
        await changeDocumentStatusToStaged();
      }


      await page.goto(binderURL, DCL);
      if (LOGIN_USER !== "Hayato.Seto@vv-agency.com") {
        console.log("BinderのステータスをStagedへ変更しています。");
        await changeDocumentStatusToStaged();
      }


      console.log("登録が正常に完了しました。");
      console.log("BinderURL:", binderURL);
      console.log("SlideURL:", slideURL);

      status = "作成完了";
      return [status, binderURL, slideURL];

    } catch (err) {
      console.log(red + "エラーが発生した為、処理を終了します。" + reset);
      console.log(err);
      return [status, page.url(), slideURL, err.toString()];
    } finally {
      await withTimeout(page.close(), 10000, "ページ終了").catch(() => null);
    }




    // await page.waitForSelector(".pageimage");





    async function menuSelect(links, target) {
      const targetText = String(target).toUpperCase();
      for (var i = 0; i < links.length; i++) {
        let text = await (await links[i].getProperty('textContent')).jsonValue();
        let uptext = String(text).toUpperCase();
        if (uptext == targetText) {
          await links[i].click();
          // console.log(i);
          break;
        }
      }
    }



  }


  await browser.close();

}());
