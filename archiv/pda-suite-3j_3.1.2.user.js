// ==UserScript==
// @name         PDA Suite 3J (Jaro · HF Slovakia)
// @namespace    http://tampermonkey.net/pda-suite-3j
// @version      3.1.2
// @description  PDA Suite 3J - Jarova verzia (mimo gitu, C:\Claude Code\PDA verzia 3J). Vychadza z produkcneho buildu 2.3.0, vsetky funkcie zachovane, upravuje sa len dizajn.
// @author       Gabris, Tvarozek
// @updateURL    file:///C:/Claude%20Code/PDA%20verzia%203J/pda-suite-3j.user.js
// @downloadURL  file:///C:/Claude%20Code/PDA%20verzia%203J/pda-suite-3j.user.js
// @match        https://hf.simplifier.cloud/*
// @icon         https://www.google.com/s2/favicons?sz=64&domain=simplifier.cloud
// @run-at       document-start
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_registerMenuCommand
// @grant        GM_xmlhttpRequest
// @grant        unsafeWindow
// @connect      172.16.77.134
// @connect      mixinggroup.sharepoint.com
// @connect      sharepoint.com
// @connect      sharepointonline.com
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// ==/UserScript==

/*
 * ============================================================================
 *  PDA Suite - jeden skript namiesto osmich samostatnych
 * ============================================================================
 *
 *  Ako to funguje:
 *    - kazde vylepsenie je samostatny MODUL (dole v sekcii MODULY)
 *    - ktore moduly bezia sa nastavuje v paneli: ozubene koliesko vpravo dole
 *      na stranke PDA, alebo cez ikonu Tampermonkey -> "Nastavenia PDA Suite"
 *    - nastavenia sa ukladaju lokalne v Tampermonkey (GM storage), takze
 *      prezijeu aktualizaciu skriptu a NIE su sucastou repozitara
 *
 *  Citlive udaje (hesla kolegov, adresa sluzby vykresov) sa zadavaju
 *  v tom paneli - zamerne nie su v kode, lebo repozitar je verejny.
 *
 *  Pridanie noveho vylepsenia = pridat funkciu + jeden riadok do zoznamu
 *  MODULES. Nic ine netreba.
 * ============================================================================
 */

(function () {
    'use strict';

    const LOG = '[PDA 3J]';
    const W = unsafeWindow;

    /* ========================================================================
     *  1. ULOZISKO NASTAVENI
     * ====================================================================== */

    const KEY_MODULES = 'pda_modules_v1';
    const KEY_USERS = 'pda_users_v1';
    const KEY_PDM = 'pda_pdm_v1';
    const KEY_EXCEL = 'pda_excel_v1';
    const KEY_GROUPS = 'pda_groups_v1';
    const KEY_BUTTONS = 'pda_buttons_v1';
    const KEY_ADMIN = 'pda_admin_v1';

    function loadJson(key, fallback) {
        try {
            const raw = GM_getValue(key, null);
            if (raw === null || raw === undefined) return fallback;
            return typeof raw === 'string' ? JSON.parse(raw) : raw;
        } catch (e) {
            console.warn(LOG, 'nepodarilo sa nacitat nastavenie', key, e);
            return fallback;
        }
    }

    function saveJson(key, value) {
        try {
            GM_setValue(key, JSON.stringify(value));
        } catch (e) {
            console.warn(LOG, 'nepodarilo sa ulozit nastavenie', key, e);
        }
    }

    const settings = {
        modules: loadJson(KEY_MODULES, {}),
        users: loadJson(KEY_USERS, []),
        pdm: loadJson(KEY_PDM, { base: 'http://172.16.77.134:9000', key: '' }),
        // url prazdna = subor sa vybera rucne cez tlacidlo "Vybrať Excel"
        excel: loadJson(KEY_EXCEL, { url: '', colOrder: 'H', colDrawing: 'AH', colVersion: 'AI' }),
        // kategorie pracovisk pre uvodny prehlad: Nazov|Podtitul|VZOR,VZOR,...
        groups: loadJson(KEY_GROUPS, [
            'Assembly|Finálna montáž a podzostavy|MONTAZ,ASSEMBLY,PODZOST,MONT',
            'Welding|Zváranie a príprava|ZVAR,TIG,MIG,WELD',
            'Machining|CNC a konvenčné obrábanie|CNC,FREZ,SUSTR,BRUS,LMS,HMS,K-TEC,VRTA,PILA,HOBL,HEDELL',
            'Quality Control|Kontrola a meranie|OTK,KONTROL,MERAN,QC,KVALIT',
        ].join('\n')),
        // pravidla farieb tlacidiel; null = pri prvom spusteni sa nasadia predvolene
        buttons: loadJson(KEY_BUTTONS, null),
        // heslo na ozubene koliesko a prepinac 'nastavovanie tlacidiel pravym klikom'
        admin: loadJson(KEY_ADMIN, { password: '123456', pickMode: false }),
    };

    // "H" -> 7, "AH" -> 33 (vracia 0-based index stlpca ako ho vidi XLSX)
    function colToIndex(letters, fallbackIndex) {
        const s = String(letters || '').trim().toUpperCase();
        if (!/^[A-Z]{1,3}$/.test(s)) return fallbackIndex;
        let n = 0;
        for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64);
        return n - 1;
    }

    function isModuleOn(mod) {
        const stored = settings.modules[mod.id];
        return stored === undefined ? mod.def : !!stored;
    }

    /* ========================================================================
     *  2. ZDIELANE POMOCKY
     * ====================================================================== */

    function sleep(ms) {
        return new Promise((r) => setTimeout(r, ms));
    }

    function onReady(fn) {
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', fn, { once: true });
        } else {
            fn();
        }
    }

    function waitFor(checkFn, { interval = 150, timeout = 10000 } = {}) {
        return new Promise((resolve, reject) => {
            const start = Date.now();
            const timer = setInterval(() => {
                let result = null;
                try { result = checkFn(); } catch (e) { /* ignore */ }
                if (result) {
                    clearInterval(timer);
                    resolve(result);
                } else if (Date.now() - start > timeout) {
                    clearInterval(timer);
                    reject(new Error('waitFor timeout'));
                }
            }, interval);
        });
    }

    // --- pristup k SAP UI5 controlom (aplikacia bezi v kontexte stranky) ---

    function getControl(controlId) {
        try {
            if (W.sap && W.sap.ui && W.sap.ui.getCore) return W.sap.ui.getCore().byId(controlId);
        } catch (e) { /* ignore */ }
        return null;
    }

    function resolveControl(el) {
        while (el && el.nodeType === 1) {
            try {
                if (W.jQuery && typeof W.jQuery(el).control === 'function') {
                    const arr = W.jQuery(el).control();
                    if (arr && arr.length && arr[0]) return arr[0];
                }
                if (W.sap && W.sap.ui && W.sap.ui.core && W.sap.ui.core.Element &&
                    typeof W.sap.ui.core.Element.closestTo === 'function') {
                    const c = W.sap.ui.core.Element.closestTo(el);
                    if (c) return c;
                }
                if (el.id && W.sap && W.sap.ui && W.sap.ui.getCore) {
                    const c = W.sap.ui.getCore().byId(el.id);
                    if (c) return c;
                }
            } catch (e) { /* ignore */ }
            el = el.parentElement;
        }
        return null;
    }

    function pressElement(domEl) {
        const control = resolveControl(domEl);
        if (control && typeof control.firePress === 'function') {
            control.firePress();
            return;
        }
        domEl.click();
    }

    function setControlValue(controlId, value) {
        try {
            const control = getControl(controlId);
            if (control && typeof control.setValue === 'function') {
                control.setValue(value);
                if (typeof control.fireLiveChange === 'function') control.fireLiveChange({ value });
                if (typeof control.fireChange === 'function') control.fireChange({ value, newValue: value });
                return true;
            }
        } catch (e) {
            console.warn(LOG, 'setControlValue zlyhalo pre', controlId, e);
        }
        const inputEl = document.getElementById(controlId + '-inner');
        if (inputEl) {
            const nativeSetter = Object.getOwnPropertyDescriptor(W.HTMLInputElement.prototype, 'value').set;
            nativeSetter.call(inputEl, value);
            inputEl.dispatchEvent(new Event('input', { bubbles: true }));
            inputEl.dispatchEvent(new Event('change', { bubbles: true }));
            return true;
        }
        return false;
    }

    /* ------------------------------------------------------------------
     *  Jeden spolocny sledovac DOM namiesto piatich samostatnych.
     *  Aplikacia je SAP UI5 a prekresluje sa sama, takze moduly musia
     *  svoje prvky dokladat opakovane. Vsetko sa zbiera do jednej davky.
     * ---------------------------------------------------------------- */
    const DomWatch = (function () {
        const callbacks = [];
        let started = false;
        let scheduled = false;

        function flush() {
            scheduled = false;
            for (const fn of callbacks) {
                try { fn(); } catch (e) { console.warn(LOG, 'chyba v DOM callbacku', e); }
            }
        }

        function schedule() {
            if (scheduled) return;
            scheduled = true;
            setTimeout(flush, 120);
        }

        function start() {
            if (started || !document.body) return;
            started = true;
            new MutationObserver(schedule).observe(document.body, {
                childList: true, subtree: true, characterData: true,
            });
        }

        return {
            add(fn) {
                callbacks.push(fn);
                onReady(() => { start(); schedule(); });
            },
            poke: schedule,
        };
    })();

    /* ------------------------------------------------------------------
     *  Jedno spolocne odpocuvanie sietovej komunikacie namiesto troch.
     *  Prepisuje XMLHttpRequest v kontexte stranky a rozposiela vysledky
     *  vsetkym prihlasenym modulom.
     * ---------------------------------------------------------------- */
    const XhrBus = (function () {
        const TARGET = '/client/1.0/executeBO';
        const listeners = [];
        let patched = false;

        function patch() {
            if (patched) return;
            patched = true;

            const proto = W.XMLHttpRequest.prototype;
            const originalOpen = proto.open;
            const originalSend = proto.send;

            proto.open = function (method, url, ...rest) {
                this.__pdaUrl = url;
                this.__pdaMethod = method;
                return originalOpen.call(this, method, url, ...rest);
            };

            proto.send = function (body) {
                const url = this.__pdaUrl || '';
                if (url.indexOf(TARGET) !== -1) {
                    const xhr = this;
                    xhr.addEventListener('load', function () {
                        let requestJson = null;
                        let responseJson = null;
                        try { requestJson = JSON.parse(body); } catch (e) { /* ignore */ }
                        try { responseJson = JSON.parse(xhr.responseText); } catch (e) { /* ignore */ }

                        const event = {
                            url,
                            requestRaw: body,
                            request: requestJson,
                            response: responseJson,
                            responseRaw: xhr.responseText,
                        };
                        for (const fn of listeners) {
                            try { fn(event); } catch (e) { console.warn(LOG, 'chyba v XHR callbacku', e); }
                        }
                    });
                }
                return originalSend.apply(this, arguments);
            };
        }

        return {
            subscribe(fn) {
                patch();
                listeners.push(fn);
            },
        };
    })();

    /* ------------------------------------------------------------------
     *  Dekodovanie vstupu zo skenera.
     *  Skener posiela cisla ako slovenske znaky (SK klavesnica), prvy znak
     *  "J" je marker zariadenia a bodka zastupuje pomlcku.
     * ---------------------------------------------------------------- */
    const SCANNER_CHAR_MAP = {
        '+': '1', 'ľ': '2', 'š': '3', 'č': '4', 'ť': '5',
        'ž': '6', 'ý': '7', 'á': '8', 'í': '9', 'é': '0',
    };

    function decodeScannerInput(raw) {
        // POZOR: prvy znak sa odstrani LEN ak je to naozaj marker "J".
        // (v povodnych skriptoch sa na jednom mieste odrezaval vzdy, co
        //  rucne napisanemu cislu zjedlo prvu cislicu)
        const stripped = raw.length > 0 && raw[0] === 'J' ? raw.slice(1) : raw;

        let out = '';
        for (const ch of stripped) {
            if (ch === '.') out += '-';
            else if (Object.prototype.hasOwnProperty.call(SCANNER_CHAR_MAP, ch)) out += SCANNER_CHAR_MAP[ch];
            else out += ch;
        }
        return out;
    }

    function padOperationPart(value) {
        const i = value.indexOf('-');
        if (i === -1) return value;
        return value.slice(0, i) + '-' + value.slice(i + 1).padStart(4, '0');
    }

    /* ------------------------------------------------------------------
     *  Prepinanie pouzivatela - zdielana sluzba.
     *  Pouziva ju aj bocny panel, aj citacka kariet, takze je tu zvlast
     *  a nie je viazana na to, ci je panel zapnuty.
     * ---------------------------------------------------------------- */
    const UserSwitch = (function () {
        const CHANGE_USER_SUFFIX = 'Button_ChangeUser';
        const USER_COMBO_ID = 'Popups--User_ComboBox';
        const PASSWORD_INPUT_ID = 'Popups--EmployeeNo_Input';
        const CONFIRM_BUTTON_ID = 'Popups--UserDialog_Button_close';
        const DIALOG_WAIT_TIMEOUT = 10000;
        const BEFORE_CONFIRM_DELAY = 100;

        async function selectComboBoxItemByText(controlId, text, { timeout = 5000 } = {}) {
            let control;
            try {
                control = await waitFor(() => {
                    const c = getControl(controlId);
                    return c && typeof c.getItems === 'function' && c.getItems().length > 0 ? c : null;
                }, { timeout });
            } catch (e) {
                console.warn(LOG, 'polozky v zozname pouzivatelov sa nenacitali vcas', e);
                return false;
            }

            const items = control.getItems();
            const match =
                items.find((it) => it.getText() === text) ||
                items.find((it) => it.getText().toLowerCase() === text.toLowerCase());

            if (!match) {
                console.warn(LOG, 'pouzivatel sa v zozname nenasiel:', text,
                    'dostupni:', items.map((i) => i.getText()));
                return false;
            }

            control.setSelectedItem(match);
            if (typeof control.fireSelectionChange === 'function') control.fireSelectionChange({ selectedItem: match });
            if (typeof control.fireChange === 'function') control.fireChange({ value: match.getText() });
            return true;
        }

        async function fillLoginPopup(username, password) {
            try {
                await waitFor(
                    () => document.getElementById(USER_COMBO_ID + '-inner') &&
                          document.getElementById(PASSWORD_INPUT_ID + '-inner'),
                    { timeout: DIALOG_WAIT_TIMEOUT }
                );

                const selected = await selectComboBoxItemByText(USER_COMBO_ID, username);
                if (!selected) setControlValue(USER_COMBO_ID, username);

                setControlValue(PASSWORD_INPUT_ID, password);
                await sleep(BEFORE_CONFIRM_DELAY);

                const confirmBtn = await waitFor(() => document.getElementById(CONFIRM_BUTTON_ID), { timeout: 5000 });
                pressElement(confirmBtn);
            } catch (e) {
                console.warn(LOG, 'dialog na zmenu pouzivatela sa neobjavil vcas', e);
            }
        }

        return {
            to(username, password) {
                const btn = document.querySelector('[id$="' + CHANGE_USER_SUFFIX + '"]');
                if (!btn) {
                    console.warn(LOG, 'tlacidlo na zmenu pouzivatela sa na obrazovke nenaslo');
                    return;
                }
                pressElement(btn);
                fillLoginPopup(username, password);
            },
            findByCardId(cardId) {
                return settings.users.find((u) => u.cardId && u.cardId === cardId) || null;
            },
        };
    })();

    // zdielany stav medzi modulmi (nahradza povodne window.PDA_* premenne)
    const shared = {
        ordersIndex: [],
        ordersIndexUpdatedAt: null,
        currentOperation: null,
        fillSearchInput: null,   // doplni modul vyhladavania, ak bezi
        pdmOpenDialog: null,     // doplni modul vykresu - okno so zoznamom vykresov
        intentionalReload: false, // nastavi panel nastaveni pred location.reload()
    };

    /* ========================================================================
     *  3. MODULY
     * ====================================================================== */

    /* ---------------------- 3.1 Vylepsena hlavicka ---------------------- */

    function modEnhancedHeader() {
        // pozor: na roznych strankach ma element ine ID (napr. Main--Label_Username-bdi
        // vs. WorkcenterDetail--Label_Username2-bdi), preto "obsahuje" + "konci na", nie presny suffix
        const USERNAME_SELECTOR = '[id*="Label_Username"][id$="-bdi"]';
        const BUTTONS = [
            { suffix: 'Button_HomeScreen-img', label: 'Pracoviská' },
            { suffix: 'Button_Reporting-img', label: 'Reporty' },
            { suffix: 'Button_Message-img', label: 'Správy' },
            { suffix: 'Button_Schedule-img', label: 'Rozvrh' },
            { suffix: 'Button_Settings-img', label: 'Admin' },
            { suffix: 'Button_ChangeUser-img', label: 'Zmena používateľa' },
            { suffix: 'Button_Logout-img', label: 'Odhlásenie' },
        ];
        const LOGOUT_BUTTON_SUFFIX = 'Button_Logout-inner';

        function apply() {
            document.querySelectorAll(USERNAME_SELECTOR).forEach((el) => {
                if (el.dataset.pdaUsernameStyled) return;
                el.style.fontSize = '1.8rem';
                el.style.fontWeight = 'bold';
                el.style.color = '#1f1f1f';
                el.dataset.pdaUsernameStyled = '1';
            });

            BUTTONS.forEach(({ suffix, label }) => {
                document.querySelectorAll('[id$="' + suffix + '"]').forEach((img) => {
                    if (img.dataset.pdaTextAdded) return;
                    const span = document.createElement('span');
                    span.className = 'sapMBtnContent';
                    span.style.marginRight = '0.6rem';
                    span.innerHTML = '<bdi>' + label + '</bdi>';
                    img.insertAdjacentElement('afterend', span);
                    img.dataset.pdaTextAdded = '1';
                });
            });

            document.querySelectorAll('[id$="' + LOGOUT_BUTTON_SUFFIX + '"]').forEach((btn) => {
                if (btn.dataset.pdaLogoutStyled) return;
                btn.style.backgroundColor = '#d67a74';
                btn.style.borderRadius = '4px';
                btn.dataset.pdaLogoutStyled = '1';
            });
        }

        DomWatch.add(apply);
    }

    /* ------------------ 3.2 Farebne tlacidla (pravidla) ----------------- */

    /*
     * Portovane z Python appky (SKIN + PICKER v pda_action.py, karta "Farby tlacidiel").
     *  - pravidla { text, id, bg, fg, poradie } su v nastaveniach (ulozisko Tampermonkey),
     *    NIE v kode; pri prvom spusteni sa nasadia predvolene farby stavovych tlacidiel,
     *    takze sa oproti doterajsiemu spravaniu nic nemeni
     *  - matchuje sa podla TEXTU tlacidla ("obsahuje", bez ohladu na velkost pismen);
     *    ID sa pouzije len ked tlacidlo text nema - ID stavovych tlacidiel obsahuje
     *    poradove cislo, ktore sa lisi podla pracoviska
     *  - farba textu sa dopocita z jasu pozadia
     *  - pravy klik na tlacidlo (ked je v nastaveniach zapnute "Nastavovanie tlacidiel")
     *    ukaze paletu 12 farieb + Reset; vyber sa hned ulozi a nanesie
     *  - `poradie` drzi doterajsie zoradenie stavovych tlacidiel (vyroba, prestoj, chyba)
     */
    const DEFAULT_BUTTON_RULES = [
        { text: 'Výroba', id: '', bg: '#4e9041', fg: '#ffffff', poradie: 0 },
        { text: 'Upinanie', id: '', bg: '#4e9041', fg: '#ffffff', poradie: 0 },
        { text: 'Programovanie', id: '', bg: '#d66c37', fg: '#ffffff', poradie: 1 },
        { text: 'Upratovanie stola', id: '', bg: '#d66c37', fg: '#ffffff', poradie: 1 },
        { text: 'Meranie v Procese s OTK', id: '', bg: '#d66c37', fg: '#ffffff', poradie: 1 },
        { text: 'Chyba programu', id: '', bg: '#d04040', fg: '#ffffff', poradie: 2 },
    ];

    // rovnakych 12 farieb ako v Python verzii
    const BUTTON_PALETTE = ['#16a34a', '#65a30d', '#eab308', '#d97706', '#dc2626', '#db2777',
                            '#7c3aed', '#2563eb', '#0891b2', '#7c4a1e', '#64748b', '#1e293b'];

    function contrastColor(hex) {
        const h = String(hex || '').replace('#', '');
        if (h.length !== 6) return '#ffffff';
        const r = parseInt(h.slice(0, 2), 16), g = parseInt(h.slice(2, 4), 16), b = parseInt(h.slice(4, 6), 16);
        return (0.299 * r + 0.587 * g + 0.114 * b) > 150 ? '#1e293b' : '#ffffff';
    }

    function normalizeRule(r) {
        const bg = /^#[0-9a-f]{6}$/i.test(String(r.bg || '')) ? String(r.bg).toLowerCase() : '#2563eb';
        const fg = /^#[0-9a-f]{6}$/i.test(String(r.fg || '')) ? String(r.fg).toLowerCase() : contrastColor(bg);
        const poradie = Number.isFinite(Number(r.poradie)) ? Number(r.poradie) : 1.5;
        return { text: String(r.text || '').trim(), id: String(r.id || '').trim(), bg, fg, poradie };
    }

    function buttonRules() {
        if (!Array.isArray(settings.buttons)) {
            settings.buttons = DEFAULT_BUTTON_RULES.map((r) => Object.assign({}, r));
        }
        return settings.buttons;
    }

    function saveButtonRules(rules) {
        settings.buttons = rules.map(normalizeRule).filter((r) => r.text || r.id);
        saveJson(KEY_BUTTONS, settings.buttons);
        mirrorSettingsToFile(false);
    }

    // rovnake pravidlo ako Python `bezPravidla`: pravidlo "patri" tlacidlu podla textu alebo ID
    function ruleMatches(rule, txt, id) {
        const t = String(rule.text || '').toLowerCase();
        return (t && String(txt || '').toLowerCase().indexOf(t) !== -1) || (rule.id && rule.id === id);
    }

    function modButtonColors() {
        const CONTAINER_ID = 'WorkcenterDetail--Order_Status_Flexbox';
        const OWN_UI = '#__pda_settings_overlay__, #__pda_settings_pass__, #__pda_overview__, #__pda_pdm_overlay__, #__pda_button_menu__';
        const STYLE_ID = '__pda_status_buttons_styles__';

        /*
         * Tvar stavovych tlacidiel (farby nanasa `paint` inline podla pravidiel):
         * povodne boli vysoke cez 100 px a pri viacerych riadkoch bola medzi
         * riadkami velka diera. Teraz su nizsie, s jemnym ramom a tienom a pri
         * prechode mysou sa o 2 px nadvihnu.
         */
        function injectStyles() {
            if (document.getElementById(STYLE_ID)) return;
            const st = document.createElement('style');
            st.id = STYLE_ID;
            st.textContent = `
/* Plati pre VSETKY stavove tlacidla - aj pre osobny stav hore (Stretnutie,
   Prestavka, cakanie), lebo appka im dava tu istu triedu statusBtn.
   V riadku su rovnako vysoke: natiahnu sa na to najvyssie (dvojriadkove
   "Meranie v Procese s OTK"), text ostava zvisle na stred. */
.pda-status-row { align-content:flex-start !important; align-items:stretch !important; row-gap:0 !important; }
.statusBtn { height:auto !important; min-height:0 !important; margin:4px !important;
  align-self:stretch !important; border-radius:12px !important;
  border:2px solid #13315c !important;
  box-shadow:0 2px 6px rgba(16,36,63,.20) !important;
  transition:transform .13s ease, box-shadow .13s ease !important; }
.statusBtn .sapMBtnInner { height:100% !important; width:100% !important; min-height:0 !important;
  padding:11px 16px !important; border-radius:12px !important; box-shadow:none !important;
  display:flex !important; align-items:center !important; justify-content:center !important;
  box-sizing:border-box !important; }
/* sirka podla textu - nazov sa uz nezalomi ("Stretnuti / a", "Prestavk / a") */
.statusBtn { width:auto !important; min-width:130px !important; max-width:none !important; }
.statusBtn .sapMBtnContent, .statusBtn bdi { line-height:1.25 !important; white-space:nowrap !important;
  overflow:visible !important; text-overflow:clip !important; }
.statusBtn:hover { transform:translateY(-3px) !important;
  box-shadow:0 10px 20px rgba(16,36,63,.30) !important; }
.statusBtn:active { transform:translateY(-1px) !important;
  box-shadow:0 3px 8px rgba(16,36,63,.24) !important; }
/* panel "Osobny stav" na uvodnej obrazovke bol vysoky na pol obrazovky,
   hoci v nom su len tri tlacidla - zmensime jeho vnutorne odsadenie */
.pda-panel-tesny, .pda-panel-tesny .sapMPanelContent { min-height:0 !important; height:auto !important; }
.pda-panel-tesny .sapMPanelContent { padding-top:2px !important; padding-bottom:6px !important; }
.pda-panel-tesny .sapMPanelHdr, .pda-panel-tesny .sapMPanelHeaderTB {
  min-height:0 !important; padding-top:2px !important; padding-bottom:0 !important; }
.pda-panel-tesny .sapMFlexBox { min-height:0 !important; }
`;
            document.head.appendChild(st);
        }

        function textOf(btn) {
            return (btn.textContent || '').trim();
        }

        // najkonkretnejsie (najdlhsie) textove pravidlo vyhrava; ID len ak nic textove nesedi
        function ruleFor(btn) {
            const txt = textOf(btn).toLowerCase();
            const id = btn.id || '';
            let best = null, byId = null;
            for (const r of buttonRules()) {
                const t = String(r.text || '').toLowerCase();
                if (t) {
                    if (txt.indexOf(t) !== -1 && (!best || t.length > String(best.text).length)) best = r;
                } else if (r.id && r.id === id) {
                    byId = r;
                }
            }
            return best || byId;
        }

        function paint(btn, rule) {
            const inner = btn.querySelector('.sapMBtnInner') || btn;
            const content = btn.querySelector('.sapMBtnContent');
            if (rule) {
                inner.style.setProperty('background-image', 'none', 'important'); // gradient temy by farbu prekryl
                inner.style.setProperty('background-color', rule.bg, 'important');
                inner.style.setProperty('border-color', rule.bg, 'important');
                inner.style.setProperty('color', rule.fg, 'important');
                if (content) content.style.setProperty('color', rule.fg, 'important');
                btn.dataset.pdaPainted = '1';
            } else if (btn.dataset.pdaPainted) {
                ['background-image', 'background-color', 'border-color', 'color'].forEach((p) => inner.style.removeProperty(p));
                if (content) content.style.removeProperty('color');
                delete btn.dataset.pdaPainted;
            }
        }

        function apply() {
            injectStyles();
            // riadok, v ktorom stavove tlacidla sedia, musi natahovat na rovnaku vysku
            document.querySelectorAll('.statusBtn').forEach((b) => {
                const p = b.parentElement;
                if (p && !p.classList.contains('pda-status-row')) p.classList.add('pda-status-row');

                // panel osobneho stavu (uvodna obrazovka) je zbytocne vysoky;
                // panelov v detaile pracoviska sa nedotykame
                if (b.closest('[id^="WorkcenterDetail--"]')) return;
                const panel = b.closest('.sapMPanel');
                if (panel && !panel.classList.contains('pda-panel-tesny')) panel.classList.add('pda-panel-tesny');
            });
            document.querySelectorAll('.sapMBtn').forEach((btn) => {
                if (btn.closest(OWN_UI)) return;
                paint(btn, ruleFor(btn));
                if (btn.classList.contains('statusBtn')) {
                    [btn, btn.querySelector('.sapMBtnInner')].forEach((el) => el && el.style.setProperty('border-radius', '10px', 'important'));
                }
            });

            // zoradenie stavovych tlacidiel podla `poradie` (0 vyroba, 1 prestoj, 2 chyba, inak 1.5)
            const container = document.getElementById(CONTAINER_ID);
            if (!container) return;
            const buttons = Array.from(container.children).filter((el) => el.classList.contains('statusBtn'));
            if (buttons.length < 2) return;
            const weight = (btn) => {
                const r = ruleFor(btn);
                return r && Number.isFinite(Number(r.poradie)) ? Number(r.poradie) : 1.5;
            };
            const sorted = [...buttons].sort((a, b) => weight(a) - weight(b));
            if (!sorted.every((b, i) => buttons[i] === b)) sorted.forEach((b) => container.appendChild(b));
        }

        DomWatch.add(apply);

        /* ---- pravy klik: paleta farieb (len ked je v nastaveniach zapnute) ---- */

        let menu = null;

        function closeMenu() {
            if (menu) { menu.remove(); menu = null; }
        }

        function openMenu(e, btn) {
            closeMenu();
            const txt = textOf(btn).slice(0, 60);
            const id = btn.id || '';

            menu = document.createElement('div');
            menu.id = '__pda_button_menu__';
            menu.style.cssText =
                'position:fixed;z-index:2147483002;background:#13315c;color:#fff;' +
                'font:12px/1.4 "Segoe UI",system-ui,sans-serif;border-radius:12px;padding:12px 14px;' +
                'box-shadow:0 6px 20px rgba(0,0,0,.35);max-width:440px;';
            menu.style.left = Math.max(4, Math.min(e.clientX, W.innerWidth - 460)) + 'px';
            menu.style.top = Math.max(4, Math.min(e.clientY, W.innerHeight - 170)) + 'px';

            const title = document.createElement('div');
            title.style.cssText = 'font-weight:700;margin-bottom:4px;';
            title.textContent = txt || '(bez textu)';
            const sub = document.createElement('div');
            sub.style.cssText = 'color:#a8c0e0;font-size:10px;word-break:break-all;margin-bottom:8px;';
            sub.textContent = id || '(bez ID)';
            const label = document.createElement('div');
            label.style.cssText = 'color:#a8c0e0;font-size:11px;margin-bottom:4px;';
            label.textContent = 'Zmeniť farbu';

            const swatches = document.createElement('div');
            swatches.style.marginBottom = '10px';
            BUTTON_PALETTE.forEach((c) => {
                const sw = document.createElement('span');
                sw.style.cssText = 'display:inline-block;width:28px;height:28px;border-radius:8px;margin:3px;cursor:pointer;' +
                    'border:1px solid rgba(255,255,255,.28);background:' + c + ';';
                sw.title = c;
                sw.addEventListener('click', (ev) => {
                    ev.stopPropagation();
                    const fg = contrastColor(c);
                    const old = buttonRules().find((r) => ruleMatches(r, txt, id));
                    const rest = buttonRules().filter((r) => !ruleMatches(r, txt, id));
                    rest.push({ text: txt, id: txt ? '' : id, bg: c, fg, poradie: old ? old.poradie : 1.5 });
                    saveButtonRules(rest);
                    paint(btn, { bg: c, fg });
                    DomWatch.poke();
                    closeMenu();
                });
                swatches.appendChild(sw);
            });

            const reset = document.createElement('button');
            reset.type = 'button';
            reset.textContent = 'Reset tlačidla';
            reset.style.cssText = 'background:transparent;color:#a8c0e0;border:1px solid #1c478a;border-radius:9px;' +
                'padding:7px 14px;cursor:pointer;font:inherit;font-size:12px;';
            reset.addEventListener('click', (ev) => {
                ev.stopPropagation();
                saveButtonRules(buttonRules().filter((r) => !ruleMatches(r, txt, id)));
                paint(btn, null);
                DomWatch.poke();
                closeMenu();
            });

            menu.appendChild(title);
            menu.appendChild(sub);
            menu.appendChild(label);
            menu.appendChild(swatches);
            menu.appendChild(reset);
            document.body.appendChild(menu);
        }

        document.addEventListener('click', (e) => { if (menu && !menu.contains(e.target)) closeMenu(); }, true);
        document.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeMenu(); }, true);
        document.addEventListener('contextmenu', (e) => {
            if (!settings.admin || !settings.admin.pickMode) return;
            const btn = e.target && e.target.closest ? e.target.closest('.sapMBtn') : null;
            if (!btn || btn.closest(OWN_UI)) return;
            e.preventDefault();
            e.stopPropagation();
            openMenu(e, btn);
        }, true);
    }

    /* ------------------ 3.3 Blokovanie tlacidla Spat -------------------- */

    function modPreventBack() {
        function lockCurrentState() {
            W.history.pushState(W.history.state, document.title, location.href);
        }

        onReady(() => {
            lockCurrentState();
            W.addEventListener('popstate', lockCurrentState);
            W.addEventListener('beforeunload', (e) => {
                // obnovenie stranky vyvolane panelom nastaveni sa nema pytat
                if (shared.intentionalReload) return;
                e.preventDefault();
                e.returnValue = '';
            });
        });
    }

    /* --------------- 3.4 Vyhladavanie naprieč pracoviskami -------------- */

    function modCrossSearch() {
        const BO_METHOD = 'getOperationListTempForWorkcenters';
        const WORKCENTER_TILE_SELECTOR = '[id^="Main--Workcenter_Toolbar-Main--ui_layout_Grid3-"]';
        const CONTENT_ID = 'Main--Workcenter_Panel-content';
        const HOME_BUTTON_ID = 'Main--Button_HomeScreen';
        const PANEL_ID = 'Main--Workcenter_Panel';
        const PARENT_SECTION_ID = 'Main--MainPage-cont';
        const SIDEBAR_ID = '__pda_search_sidebar__';
        const UI_ID = '__pda_custom_search_ui__';
        const TILE_ID_PREFIX = 'Main--Workcenter_Toolbar-';
        const LIST_ID_PREFIX = 'Main--List2-';
        const TILE_WAIT_TIMEOUT = 15000;
        const SETTLE_DELAY = 350;

        let opening = false;
        let renderFn = null;
        let lastAutoOpenedKey = null;

        XhrBus.subscribe((ev) => {
            if (!ev.request || ev.request.BOMethod !== BO_METHOD) return;
            if (!ev.response || !ev.response.result) return;

            const workcenters = ev.response.result.aOperationListTempForWorkcenter || [];
            const flat = [];
            workcenters.forEach((wc) => {
                (wc.operationList || []).forEach((op) => {
                    flat.push({
                        workcenter: wc.workcenterDescription,
                        workcenterCode: wc.workcenter,
                        salesOrderNo: op.salesOrderNo,
                        salesOrderItem: op.salesOrderItem,
                        productionOrderNo: op.productionOrderNo,
                        operationNo: op.operationNo,
                        sequenceNo: op.sequenceNo,
                        materialNo: op.materialNo,
                        material: op.material,
                        description: op.description,
                        descriptionShort: op.descriptionShort,
                        status: op.status,
                        confirmed: op.confirmed,
                    });
                });
            });

            shared.ordersIndex = flat;
            shared.ordersIndexUpdatedAt = new Date();
            console.log(LOG, 'index zakaziek naplneny:', flat.length);

            if (renderFn) {
                const input = document.querySelector('#' + UI_ID + ' input.sapMSFI');
                renderFn(input ? input.value : '');
            }
        });

        function getWorkcenterName(tile) {
            const suffix = tile.id.replace(TILE_ID_PREFIX, '');
            const nameEl = document.getElementById('Main--WorkcenterDescription_Label-' + suffix + '-bdi');
            return nameEl ? nameEl.textContent.trim() : tile.id;
        }

        function getListIdForTile(tile) {
            if (!tile.id || tile.id.indexOf(TILE_ID_PREFIX) !== 0) return null;
            return LIST_ID_PREFIX + tile.id.slice(TILE_ID_PREFIX.length);
        }

        function extractOrderText(li) {
            const titleEl = li.querySelector('[id*="Title"][id$="-inner"]');
            return titleEl ? titleEl.textContent.trim() : '';
        }

        function formatProductionOrder(item) {
            return item.productionOrderNo + ' - ' + item.operationNo + ' - ' + item.sequenceNo;
        }

        function getItemKey(item) {
            return [item.workcenter, item.productionOrderNo, item.operationNo, item.sequenceNo].join('|');
        }

        function setStatus(msg) {
            const el = document.getElementById('__pda_status__');
            if (el) el.textContent = msg;
        }

        function fireListItemPress(li) {
            const itemControl = resolveControl(li);
            if (!itemControl) return false;

            let listControl = itemControl;
            while (listControl && typeof listControl.fireItemPress !== 'function') {
                listControl = listControl.getParent && listControl.getParent();
            }
            if (!listControl) return false;

            let itemForEvent = itemControl;
            while (itemForEvent && itemForEvent.getParent && itemForEvent.getParent() !== listControl) {
                itemForEvent = itemForEvent.getParent();
            }
            if (!itemForEvent) itemForEvent = itemControl;

            listControl.fireItemPress({ listItem: itemForEvent, srcControl: itemForEvent });
            return true;
        }

        async function openItem(item) {
            if (opening) return;
            opening = true;
            try {
                setStatus('Otváram zákazku...');

                const tiles = Array.from(document.querySelectorAll(WORKCENTER_TILE_SELECTOR));
                const tile = tiles.find((t) => getWorkcenterName(t) === item.workcenter);
                if (!tile) { setStatus('Pracovisko sa nenašlo.'); return; }

                const listId = getListIdForTile(tile);
                let list = listId ? document.getElementById(listId) : null;

                if (!list) {
                    pressElement(tile);
                    list = listId ? await waitFor(() => document.getElementById(listId), { timeout: TILE_WAIT_TIMEOUT }) : null;
                }
                if (!list) { setStatus('Zoznam zákaziek pre toto pracovisko sa nenašiel.'); return; }

                list.scrollIntoView({ block: 'center' });
                await sleep(SETTLE_DELAY);

                const targetText = formatProductionOrder(item);
                let li = Array.from(list.querySelectorAll('li')).find((el) => extractOrderText(el) === targetText);
                if (!li) {
                    await sleep(SETTLE_DELAY);
                    li = Array.from(list.querySelectorAll('li')).find((el) => extractOrderText(el) === targetText);
                }
                if (!li) { setStatus('Konkrétna zákazka sa nenašla, otvorené je aspoň pracovisko.'); return; }

                li.scrollIntoView({ block: 'center' });
                const fired = fireListItemPress(li);
                if (!fired) li.click();
                setStatus(fired ? 'Otvorené: ' + targetText : 'Chyba pri otváraní zákazky.');
            } catch (e) {
                console.warn(LOG, 'chyba pri otvarani polozky', e);
                setStatus('Chyba pri otváraní zákazky.');
            } finally {
                opening = false;
            }
        }

        function matchesTerm(it, term) {
            const dashIndex = term.indexOf('-');
            if (dashIndex !== -1) {
                const orderPart = term.slice(0, dashIndex).trim();
                const opPart = term.slice(dashIndex + 1).trim();
                const orderMatch = !orderPart || (it.productionOrderNo || '').toLowerCase().includes(orderPart);
                const opMatch = !opPart || (it.operationNo || '').toLowerCase().includes(opPart);
                return orderMatch && opMatch;
            }
            return (
                (it.productionOrderNo || '').toLowerCase().includes(term) ||
                (it.salesOrderNo || '').toLowerCase().includes(term) ||
                (it.materialNo || '').toLowerCase().includes(term) ||
                (it.material || '').toLowerCase().includes(term)
            );
        }

        function buildSearchUI(sidebar) {
            const list = document.createElement('div');
            list.id = UI_ID;
            list.className = 'sapMList sapMListBGSolid';
            list.style.width = '100%';

            const header = document.createElement('div');
            header.className = 'sapMIBar sapMTB sapMTBNewFlex sapMTBInactive sapMTBStandard sapMTB-Transparent-CTX sapMListHdr sapMListHdrTBar sapMTBHeader-CTX';
            const headerTitle = document.createElement('div');
            headerTitle.className = 'sapMTitle sapMTitleStyleAuto sapMTitleNoWrap sapUiSelectable sapMTitleMaxWidth sapMTitleTB sapMBarChild sapMTBShrinkItem';
            headerTitle.innerHTML = '<span dir="auto">Vyhľadať zákazku</span>';
            header.appendChild(headerTitle);

            const tbContainer = document.createElement('div');
            tbContainer.className = 'sapMListInfoTBarContainer';
            const toolbar = document.createElement('div');
            toolbar.className = 'sapMIBar sapMTB sapMTBNewFlex sapMTBInactive sapMTBClear sapMTB-Transparent-CTX sapMListInfoTBar';
            const sf = document.createElement('div');
            sf.className = 'sapMSF sapMSFVal sapMBarChild sapMTBShrinkItem';
            sf.style.width = '100%';

            const form = document.createElement('form');
            form.className = 'sapMSFF';
            form.addEventListener('submit', (e) => e.preventDefault());

            const input = document.createElement('input');
            input.type = 'search';
            input.autocomplete = 'off';
            input.placeholder = 'Číslo zákazky / materiál';
            input.className = 'sapMSFI';

            const resetDiv = document.createElement('div');
            resetDiv.title = 'Resetovať';
            resetDiv.className = 'sapMSFR sapMSFB';
            resetDiv.addEventListener('click', () => { input.value = ''; render(''); });

            const searchDiv = document.createElement('div');
            searchDiv.title = 'Hľadať';
            searchDiv.className = 'sapMSFS sapMSFB';

            form.appendChild(input);
            form.appendChild(resetDiv);
            form.appendChild(searchDiv);
            sf.appendChild(form);
            toolbar.appendChild(sf);
            tbContainer.appendChild(toolbar);

            const status = document.createElement('div');
            status.id = '__pda_status__';
            status.style.fontSize = '0.72rem';
            status.style.color = '#888';
            status.style.padding = '0.2rem 0.5rem';

            const resultsUl = document.createElement('ul');
            resultsUl.className = 'sapMListItems sapMListUl sapMListHighlight sapMListShowSeparatorsAll sapMListModeSingleSelectMaster';
            resultsUl.setAttribute('role', 'listbox');
            resultsUl.tabIndex = 0;

            function renderItem(it) {
                const li = document.createElement('li');
                li.tabIndex = 0;
                li.setAttribute('role', 'option');
                li.className = 'sapMLIB sapMLIB-CTX sapMLIBShowSeparator sapMLIBTypeActive sapMLIBActionable sapMLIBHoverable sapMLIBFocusable sapMCLI sapUiTinyMargin';
                li.innerHTML =
                    '<div class="sapMLIBContent">' +
                    '<div class="sapMFlexBoxFit sapMFlexBox sapMHBox sapMFlexBoxJustifyStart sapMFlexBoxAlignItemsStretch sapMFlexBoxWrapNoWrap sapMFlexBoxAlignContentSpaceBetween sapMFlexBoxBGTransparent" style="height:100%;width:100%;">' +
                    '<div class="sapMFlexBox sapMVBox sapMFlexBoxJustifyStart sapMFlexBoxAlignItemsStretch sapMFlexBoxWrapNoWrap sapMFlexBoxAlignContentSpaceBetween sapMFlexBoxBGTransparent sapMFlexItem" style="height:100%;width:100%;">' +
                    '<span class="sapMLabel sapUiSelectable sapMLabelMaxWidth sapUiTinyMargin sapUiNoMarginBottom sapMFlexItem" style="font-weight:bold;text-align:left;"><span class="sapMLabelTextWrapper"><bdi>' + formatProductionOrder(it) + '</bdi></span></span>' +
                    '<span class="sapMLabel sapUiSelectable sapMLabelMaxWidth sapUiTinyMargin sapUiNoMarginBottom sapMFlexItem" style="text-align:left;"><span class="sapMLabelTextWrapper"><bdi>' + (it.workcenter || '') + '</bdi></span></span>' +
                    '<span class="sapMLabel sapUiSelectable sapMLabelMaxWidth sapUiTinyMargin sapUiNoMarginBottom sapMFlexItem" style="text-align:left;"><span class="sapMLabelTextWrapper"><bdi>' + (it.material || '') + '</bdi></span></span>' +
                    '<span class="sapMLabel sapUiSelectable sapMLabelMaxWidth sapMFlexItem" style="text-align:left;"><span class="sapMLabelTextWrapper"><bdi>' + (it.descriptionShort || '') + '</bdi></span></span>' +
                    '</div></div></div>';

                li.addEventListener('click', () => openItem(it));
                li.addEventListener('keydown', (e) => {
                    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); openItem(it); }
                });
                return li;
            }

            function idleStatus(index) {
                return 'Index: ' + index.length + ' zákaziek' +
                    (shared.ordersIndexUpdatedAt
                        ? ' (aktualiz. ' + shared.ordersIndexUpdatedAt.toLocaleTimeString() + ')'
                        : ' (čaká sa na načítanie aplikácie)');
            }

            function render(filterText) {
                resultsUl.innerHTML = '';
                const term = (filterText || '').trim().toLowerCase();
                const index = shared.ordersIndex || [];

                if (!term) {
                    setStatus(idleStatus(index));
                    lastAutoOpenedKey = null;
                    return;
                }

                const matches = index.filter((it) => matchesTerm(it, term));
                setStatus(matches.length + ' výsledok/-ov (z ' + index.length + ' položiek)');
                matches.slice(0, 200).forEach((it) => resultsUl.appendChild(renderItem(it)));

                if (matches.length === 1) {
                    const key = getItemKey(matches[0]);
                    if (key !== lastAutoOpenedKey) {
                        lastAutoOpenedKey = key;
                        openItem(matches[0]);
                        setTimeout(() => {
                            input.value = '';
                            resultsUl.innerHTML = '';
                            setStatus(idleStatus(index));
                            lastAutoOpenedKey = null;
                        }, 500);
                    }
                } else {
                    lastAutoOpenedKey = null;
                }
            }

            input.addEventListener('input', () => render(input.value));
            input.addEventListener('keydown', (e) => {
                if (e.key !== 'Enter') return;
                e.preventDefault();
                const decoded = padOperationPart(decodeScannerInput(input.value));
                input.value = decoded;
                render(decoded);
            });

            list.appendChild(header);
            list.appendChild(tbContainer);
            list.appendChild(status);
            list.appendChild(resultsUl);
            sidebar.appendChild(list);

            render('');
            renderFn = render;

            // sprístupní vyplnenie vyhľadávania pre modul skenera
            shared.fillSearchInput = (value) => {
                input.value = value;
                render(value);
            };
        }

        function ensureLayout() {
            if (opening) return;
            if (!document.getElementById(HOME_BUTTON_ID) || !document.getElementById(CONTENT_ID)) return;

            const parentSection = document.getElementById(PARENT_SECTION_ID);
            const workcenterPanel = document.getElementById(PANEL_ID);
            if (!parentSection || !workcenterPanel) return;

            let container = document.getElementById(SIDEBAR_ID);
            if (!container) {
                container = document.createElement('div');
                container.id = SIDEBAR_ID;
                container.style.width = '100%';
                container.style.boxSizing = 'border-box';
                container.style.maxHeight = '80vh';
                container.style.overflowY = 'auto';
                container.style.margin = '10px 0';
                container.style.backgroundColor = '#ffffff';
                container.style.borderRadius = '10px';
            }

            if (container.nextElementSibling !== workcenterPanel || container.parentElement !== parentSection) {
                parentSection.insertBefore(container, workcenterPanel);
            }

            if (!document.getElementById(UI_ID)) buildSearchUI(container);
        }

        DomWatch.add(ensureLayout);
    }

    /* ------------------ 3.5 Bocny panel pouzivatelov -------------------- */

    function modUsersPanel() {
        const PARENT_ID = 'Main';
        const CONTENT_ID = 'Main--MainPage';
        const WRAPPER_ID = '__pda_main_wrapper__';
        const PANEL_ID = '__pda_user_switch_panel__';
        const STYLE_TAG_ID = '__pda_panel_styles__';
        const PANEL_WIDTH = '280px';

        function injectStyles() {
            if (document.getElementById(STYLE_TAG_ID)) return;
            const style = document.createElement('style');
            style.id = STYLE_TAG_ID;
            style.textContent =
                '.pda-user-btn { border: 2px solid transparent; }' +
                '.pda-user-btn:hover { border: 2px solid #e9e9f0; }';
            document.head.appendChild(style);
        }

        function buildUserButtons(panel) {
            if (settings.users.length === 0) {
                const hint = document.createElement('div');
                hint.textContent = 'Zatiaľ nie sú zadaní žiadni používatelia. Doplň ich v nastaveniach (ozubené koliesko vpravo dole).';
                hint.style.color = '#c9c9dd';
                hint.style.fontSize = '0.8rem';
                hint.style.lineHeight = '1.4';
                panel.appendChild(hint);
                return;
            }

            settings.users.forEach(({ username, password }) => {
                const btn = document.createElement('button');
                btn.textContent = username;
                btn.classList.add('pda-user-btn');
                btn.style.display = 'block';
                btn.style.width = '100%';
                btn.style.padding = '0.5rem 0.8rem';
                btn.style.marginBottom = '0.4rem';
                btn.style.cursor = 'pointer';
                btn.style.textAlign = 'left';
                btn.style.borderRadius = '8px';
                btn.style.color = '#e9e9f0';
                btn.style.backgroundColor = '#222252';
                btn.addEventListener('click', () => UserSwitch.to(username, password));
                panel.appendChild(btn);
            });
        }

        function ensureLayout() {
            const parent = document.getElementById(PARENT_ID);
            const content = document.getElementById(CONTENT_ID);
            if (!parent || !content) return;

            let wrapper = document.getElementById(WRAPPER_ID);
            let panel = document.getElementById(PANEL_ID);
            if (wrapper && wrapper.contains(content) && panel) return;

            if (!wrapper) {
                wrapper = document.createElement('div');
                wrapper.id = WRAPPER_ID;
                wrapper.style.display = 'flex';
                wrapper.style.flexDirection = 'row';
                wrapper.style.alignItems = 'stretch';
                wrapper.style.width = '100%';
                wrapper.style.height = '100%';
            }

            if (!panel) {
                panel = document.createElement('div');
                panel.id = PANEL_ID;
                panel.style.flex = '0 0 ' + PANEL_WIDTH;
                panel.style.maxWidth = PANEL_WIDTH;
                panel.style.marginRight = '4px';
                panel.style.background = '#313175';
                panel.style.overflowY = 'auto';
                panel.style.padding = '0.75rem';
                panel.style.boxSizing = 'border-box';

                const title = document.createElement('div');
                title.textContent = 'Rýchla zmena používateľa';
                title.style.fontWeight = 'bold';
                title.style.marginBottom = '0.5rem';
                title.style.color = '#e9e9f0';
                panel.appendChild(title);

                buildUserButtons(panel);
            }

            if (!wrapper.contains(content)) {
                parent.insertBefore(wrapper, content);
                wrapper.appendChild(panel);
                wrapper.appendChild(content);
                content.style.flex = '1 1 auto';
                content.style.minWidth = '0';
            }
        }

        onReady(injectStyles);
        DomWatch.add(ensureLayout);
    }

    /* ------------------- 3.6 Skener a RFID citacka ---------------------- */

    function modInputListener() {
        const HIDDEN_INPUT_ID = '__pda_hidden_scanner_input__';
        const CARD_ID_LENGTH = 10;
        const REFOCUS_INTERVAL = 1000;

        function handleRawInput(raw) {
            const decoded = decodeScannerInput(raw);

            if (decoded.length === CARD_ID_LENGTH) {
                const user = UserSwitch.findByCardId(decoded);
                if (user) {
                    console.log(LOG, 'karta patrí používateľovi:', user.username);
                    UserSwitch.to(user.username, user.password);
                } else {
                    console.warn(LOG, 'karta nerozpoznaná, ID:', decoded);
                }
                return;
            }

            const finalValue = padOperationPart(decoded);
            console.log(LOG, 'rozpoznané číslo zákazky:', finalValue);
            if (typeof shared.fillSearchInput === 'function') {
                shared.fillSearchInput(finalValue);
            } else {
                console.warn(LOG, 'modul vyhľadávania nebeží, zákazka nebola vložená');
            }
        }

        function createHiddenInput() {
            let input = document.getElementById(HIDDEN_INPUT_ID);
            if (input) return input;

            input = document.createElement('input');
            input.id = HIDDEN_INPUT_ID;
            input.type = 'text';
            input.autocomplete = 'off';
            input.style.position = 'fixed';
            input.style.top = '0';
            input.style.left = '0';
            input.style.width = '1px';
            input.style.height = '1px';
            input.style.opacity = '0';
            input.style.border = 'none';
            input.style.padding = '0';
            input.style.margin = '0';
            input.style.pointerEvents = 'none';

            input.addEventListener('keydown', (e) => {
                if (e.key !== 'Enter') return;
                e.preventDefault();
                const raw = input.value;
                input.value = '';
                if (raw) handleRawInput(raw);
            });

            document.body.appendChild(input);
            return input;
        }

        function ensureFocus(input) {
            const active = document.activeElement;
            const isButton = !!active && (active.tagName === 'BUTTON' || active.getAttribute('role') === 'button');
            // nekradne focus, ked pouzivatel pise do policka alebo je otvoreny panel nastaveni
            if (document.getElementById('__pda_settings_overlay__') || document.getElementById('__pda_settings_pass__')) return;
            if (!active || active === document.body || isButton) input.focus();
        }

        onReady(() => {
            const input = createHiddenInput();
            input.focus();
            setInterval(() => ensureFocus(input), REFOCUS_INTERVAL);
        });
    }

    /* ---------------------- 3.7 Tlacidlo vykresu ------------------------ */

    function modDrawingButton() {
        const DB_NAME = 'pda_drawing_db';
        const STORE_NAME = 'handles';
        const HANDLE_KEY = 'excel_file_handle';

        // stlpce sa daju prestavit v nastaveniach (predvolene H, AH, AI)
        const COL_ORDER_NO = colToIndex(settings.excel.colOrder, 7);
        const COL_DRAWING_NO = colToIndex(settings.excel.colDrawing, 33);
        const COL_VERSION = colToIndex(settings.excel.colVersion, 34);

        const CONTAINER_ID = 'WorkcenterDetail--Order_FlexBox';
        const WRAPPER_ID = '__pda_order_drawing_wrapper__';
        const BUTTON_ID = '__pda_order_drawing_button__';
        const LOAD_BUTTON_ID = '__pda_order_drawing_load_button__';

        let drawingIndex = {};
        let currentDrawingInfo = null;
        let pendingHandle = null;
        let currentLoadState = 'checking';

        const supportsFsAccess = typeof W.showOpenFilePicker === 'function';

        // --- IndexedDB (aby vyber suboru prezil obnovenie stranky) ---

        function openHandleDb() {
            return new Promise((resolve, reject) => {
                const req = W.indexedDB.open(DB_NAME, 1);
                req.onupgradeneeded = () => req.result.createObjectStore(STORE_NAME);
                req.onsuccess = () => resolve(req.result);
                req.onerror = () => reject(req.error);
            });
        }

        async function saveHandle(handle) {
            const db = await openHandleDb();
            return new Promise((resolve, reject) => {
                const tx = db.transaction(STORE_NAME, 'readwrite');
                tx.objectStore(STORE_NAME).put(handle, HANDLE_KEY);
                tx.oncomplete = () => resolve();
                tx.onerror = () => reject(tx.error);
                tx.onabort = () => reject(tx.error);
            });
        }

        async function loadHandle() {
            const db = await openHandleDb();
            return new Promise((resolve, reject) => {
                const tx = db.transaction(STORE_NAME, 'readonly');
                const req = tx.objectStore(STORE_NAME).get(HANDLE_KEY);
                req.onsuccess = () => resolve(req.result || null);
                req.onerror = () => reject(req.error);
            });
        }

        // --- Excel ---

        /*
         * Cislo vyrobnej zakazky sa v Exceli a v PDA nepise rovnako:
         *   PDA:   001600089585
         *   Excel: '1600089585   (apostrof = textova bunka, bez uvodnych nul)
         * Povodny skript to riesil tak, ze natvrdo odrezal prve dva znaky a
         * pridal apostrof - staci mala zmena formatu v Exceli a nenajde nic.
         * Preto sa obe strany prevedu na rovnaky tvar: bez apostrofu, bez
         * medzier a bez uvodnych nul.
         */
        // '001600108008' aj "'1600108008" -> '1600108008' (len cislice, bez uvodnych nul)
        function normalizeOrderKey(value) {
            return String(value === undefined || value === null ? '' : value)
                .replace(/\D/g, '')
                .replace(/^0+/, '');
        }

        // Stlpce sa hladaju podla NAZVU v hlavicke (AutomatedOQ180.xlsx ich ma
        // stale rovnake), pismena z nastaveni su len zaloha pre iny subor.
        const HEADER_NAMES = {
            order: 'production order number',
            drawing: 'document number (material of production order)',
            version: 'document version (material of production order)',
            material: 'material number (production order)',
        };

        // 'SIEHE DIS', 'B.V.', 'O.Z.' nie su cisla vykresov - cislo vykresu ma vzdy cislicu
        function isDrawingNo(v) {
            return /\d/.test(String(v || ''));
        }

        function cellText(row, col) {
            if (col < 0 || row[col] === undefined || row[col] === null) return '';
            return String(row[col]).trim().replace(/^'+/, '');
        }

        function findColumns(headerRow) {
            const hdr = (headerRow || []).map((h) => String(h === undefined || h === null ? '' : h).trim().toLowerCase());
            const cols = {
                order: hdr.indexOf(HEADER_NAMES.order),
                drawing: hdr.indexOf(HEADER_NAMES.drawing),
                version: hdr.indexOf(HEADER_NAMES.version),
                material: hdr.indexOf(HEADER_NAMES.material),
            };
            if (cols.order === -1 || cols.drawing === -1) {
                console.log(LOG, 'Excel: hlavičky stĺpcov sa nenašli, používam písmená z nastavení');
                return { order: COL_ORDER_NO, drawing: COL_DRAWING_NO, version: COL_VERSION, material: -1 };
            }
            return cols;
        }

        function buildDrawingIndex(rows) {
            const cols = findColumns(rows[0]);
            const index = {};
            for (let i = 1; i < rows.length; i++) {
                const row = rows[i];
                if (!row) continue;
                const key = normalizeOrderKey(row[cols.order]);
                if (!key || index[key]) continue;   // prvy vyskyt vyhrava (v Python verzii overene: 0 konfliktov)
                const vyk = cellText(row, cols.drawing);
                const valid = isDrawingNo(vyk);
                index[key] = {
                    drawingNo: valid ? vyk : '',
                    version: valid ? cellText(row, cols.version) : '',
                    materialNo: cellText(row, cols.material).replace(/^0+/, ''),
                };
            }
            return index;
        }

        function parseWorkbook(bytes) {
            const workbook = XLSX.read(bytes, { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            if (!sheet) {
                console.warn(LOG, 'v exceli sa nenašiel žiadny list');
                return false;
            }
            drawingIndex = buildDrawingIndex(XLSX.utils.sheet_to_json(sheet, { header: 1 }));
            console.log(LOG, 'index výkresov vytvorený, záznamov:', Object.keys(drawingIndex).length);
            updateLoadButtonState('loaded');
            if (shared.currentOperation) applyForOperation(shared.currentOperation);
            return true;
        }

        async function loadExcelFromFile(file) {
            parseWorkbook(new Uint8Array(await file.arrayBuffer()));
        }

        /*
         * Zdroj Excelu z nastaveni:
         *   http(s)://...          -> stiahne sa zo servera
         *   C:\...  alebo  \\server\... alebo file:///... -> precita sa z disku cez file://
         *      (rovnako ako v Python appke; Tampermonkey na to potrebuje zapnute
         *       "Povolit pristup k URL adresam suborov" na stranke chrome://extensions)
         *   prazdne                -> rucny vyber suboru, prehliadac si ho zapamata
         */
        function excelSource() {
            let raw = String(settings.excel.url || '').trim().replace(/^"+|"+$/g, '');
            if (!raw) return { kind: 'none', url: '', raw: '' };
            if (/^https?:\/\//i.test(raw)) return { kind: 'http', url: raw, raw };
            if (/^file:\/\//i.test(raw)) return { kind: 'file', url: raw, raw };
            if (/^\\\\/.test(raw)) {                       // \\server\share\subor -> file://server/share/subor
                return { kind: 'file', url: 'file:' + encodeURI(raw.replace(/\\/g, '/')), raw };
            }
            if (/^[a-zA-Z]:[\\/]/.test(raw)) {               // C:\priecinok\subor -> file:///C:/priecinok/subor
                return { kind: 'file', url: 'file:///' + encodeURI(raw.replace(/\\/g, '/')), raw };
            }
            return { kind: 'invalid', url: '', raw };
        }

        // Excel stiahnuty z adresy (http://... alebo file://...). GM_xmlhttpRequest obchadza CORS,
        // takze sa da siahnut aj na interny server; pri file:// vracia Chrome status 0.
        function fetchExcel(url) {
            return new Promise((resolve, reject) => {
                GM_xmlhttpRequest({
                    method: 'GET',
                    url,
                    responseType: 'arraybuffer',
                    onload: (r) => {
                        const hasData = r.response && r.response.byteLength > 0;
                        const ok = (r.status >= 200 && r.status < 300 && hasData) || (r.status === 0 && hasData);
                        if (!ok) {
                            reject(new Error(r.status === 0
                                ? 'prehliadač súbor nevydal — chýba povolenie prístupu k súborom?'
                                : 'HTTP ' + r.status));
                            return;
                        }
                        resolve(r.response);
                    },
                    onerror: () => reject(new Error(/^file:/i.test(url)
                        ? 'súbor sa nedá otvoriť — cesta alebo povolenie prístupu k súborom'
                        : 'spojenie so serverom zlyhalo')),
                    ontimeout: () => reject(new Error('zdroj neodpovedal včas')),
                });
            });
        }

        async function loadExcelFromUrl(url) {
            updateLoadButtonState('loading');
            try {
                const buffer = await fetchExcel(url);
                if (!buffer) throw new Error('prázdna odpoveď');
                parseWorkbook(new Uint8Array(buffer));
            } catch (e) {
                console.warn(LOG, 'Excel sa nepodarilo načítať z', url, e);
                updateLoadButtonState(/^file:/i.test(url) ? 'file-denied' : 'url-error');
            }
        }

        async function pickFileAndRemember() {
            try {
                const [handle] = await W.showOpenFilePicker({
                    types: [{
                        description: 'Excel',
                        accept: {
                            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
                            'application/vnd.ms-excel': ['.xls'],
                        },
                    }],
                    multiple: false,
                });
                await saveHandle(handle);
                updateLoadButtonState('loading');
                await loadExcelFromFile(await handle.getFile());
            } catch (err) {
                if (err && err.name === 'AbortError') return;
                console.warn(LOG, 'chyba pri výbere súboru', err);
                updateLoadButtonState('error');
            }
        }

        async function tryAutoLoad() {
            if (!supportsFsAccess) { updateLoadButtonState('unsupported'); return; }

            let handle = null;
            try { handle = await loadHandle(); } catch (e) { /* ignore */ }
            if (!handle) { updateLoadButtonState('nofile'); return; }

            let permission;
            try {
                permission = await handle.queryPermission({ mode: 'read' });
            } catch (e) {
                updateLoadButtonState('needs-permission', handle);
                return;
            }

            if (permission === 'granted') {
                updateLoadButtonState('loading');
                try { await loadExcelFromFile(await handle.getFile()); }
                catch (e) { updateLoadButtonState('error'); }
                return;
            }
            updateLoadButtonState('needs-permission', handle);
        }

        async function confirmPermissionAndLoad(handle) {
            try {
                const permission = await handle.requestPermission({ mode: 'read' });
                if (permission !== 'granted') { updateLoadButtonState('needs-permission', handle); return; }
                updateLoadButtonState('loading');
                await loadExcelFromFile(await handle.getFile());
            } catch (e) {
                updateLoadButtonState('error');
            }
        }

        // --- napojenie na aktualne otvorenu operaciu ---

        /*
         * Prave otvorena operacia sa cita priamo z premennej appky
         * getGlobals().getVar('oSelectedWorkcenterOperation') - rovnako ako v
         * Python verzii. Povodne odpocuvanie siete cakalo na result.operation,
         * ktore appka neposiela spolahlivo, preto tlacidlo ostavalo prazdne.
         */
        function readSelectedOperation() {
            try {
                const main = W.sap.ui.getCore().byId('Main');
                const op = main && main.getController().getGlobals().getVar('oSelectedWorkcenterOperation');
                if (!op || !op.productionOrderNo) return null;
                return {
                    workcenter: op.workcenterDescription,
                    workcenterCode: op.workcenter,
                    productionOrderNo: String(op.productionOrderNo || ''),
                    salesOrderNo: String(op.salesOrderNo || ''),
                    operationNo: String(op.operationNo || ''),
                    sequenceNo: op.sequenceNo,
                    materialNo: String(op.materialNo || '').replace(/^0+/, ''),
                    material: op.material,
                };
            } catch (e) {
                return null;
            }
        }

        /*
         * Mapa vyrobna zakazka -> { kluc (cislo vykresu alebo material), typ, revizia }.
         * Stiahne sa zo sluzby RAZ pri starte (endpoint /zakazky = unikatne zakazky z Excelu,
         * prvy vyskyt vyhrava) a drzi sa v pamati aj v ulozisku Tampermonkey, aby prezila
         * obnovenie stranky. Znovu sa stahuje len ked sa na serveri zmeni odtlacok mapy.
         * Do Excelu sa z prehliadaca nechodi vobec - cita ho server raz denne.
         */
        const KEY_ORDER_MAP = 'pda_zakazky_cache_v1';
        let orderMap = null;
        let orderMapStamp = '';
        let orderMapLoading = false;

        (function restoreOrderMap() {
            const c = loadJson(KEY_ORDER_MAP, null);
            if (c && c.zakazky) {
                orderMap = c.zakazky;
                orderMapStamp = String(c.mapa_z || '');
                console.log(LOG, 'mapa zákaziek z úložiska:', Object.keys(orderMap).length, 'zákaziek (mapa z ' + orderMapStamp + ')');
            }
        })();

        function pdmGetJson(path, timeout) {
            return new Promise((resolve, reject) => {
                GM_xmlhttpRequest({
                    method: 'GET',
                    url: settings.pdm.base + path,
                    headers: settings.pdm.key ? { 'X-API-Key': settings.pdm.key } : {},
                    timeout: timeout || 15000,
                    onload: (r) => {
                        if (r.status === 404) { resolve(null); return; }
                        if (r.status < 200 || r.status >= 300) { reject(new Error('HTTP ' + r.status + ' ' + path)); return; }
                        try { resolve(JSON.parse(r.responseText)); }
                        catch (e) { reject(new Error('neplatná odpoveď ' + path)); }
                    },
                    onerror: () => reject(new Error('spojenie so službou zlyhalo')),
                    ontimeout: () => reject(new Error('služba neodpovedala včas')),
                });
            });
        }

        async function loadOrderMap() {
            if (orderMapLoading) return;
            orderMapLoading = true;
            try {
                // lacna kontrola odtlacku: /health ma ~400 B, cela mapa stovky kB
                const h = await pdmGetJson('/health', 8000);
                const stamp = h && h.mapa && h.mapa.posledne ? String(h.mapa.posledne) : '';
                if (orderMap && stamp && stamp === orderMapStamp) {
                    console.log(LOG, 'mapa zákaziek je aktuálna (' + stamp + ')');
                    return;
                }
                const d = await pdmGetJson('/zakazky', 30000);
                if (!d || !d.zakazky) {
                    console.log(LOG, 'služba ešte nemá /zakazky — hľadám po jednej zákazke');
                    return;
                }
                orderMap = d.zakazky;
                orderMapStamp = String(d.mapa_z || stamp || '');
                saveJson(KEY_ORDER_MAP, { mapa_z: orderMapStamp, zakazky: orderMap });
                console.log(LOG, 'mapa zákaziek stiahnutá:', d.pocet, 'zákaziek (mapa z ' + orderMapStamp + ')');
                if (shared.currentOperation) applyForOperation(shared.currentOperation);
            } catch (e) {
                console.log(LOG, 'mapa zákaziek sa nestiahla:', e.message);
            } finally {
                orderMapLoading = false;
            }
        }

        let lastOperationKey = null;

        function pollSelectedOperation() {
            const cur = readSelectedOperation();
            const key = cur ? cur.productionOrderNo + '|' + cur.operationNo + '|' + cur.sequenceNo : '';
            if (key === lastOperationKey) return;
            lastOperationKey = key;
            shared.currentOperation = cur;
            if (cur) applyForOperation(cur);
            else updateDisplay(null);
        }

        onReady(loadOrderMap);
        onReady(() => setInterval(pollSelectedOperation, 500));

        function stripLeadingApostrophe(value) {
            if (!value) return value;
            return value.charAt(0) === "'" ? value.slice(1) : value;
        }

        function salesOrderOf(current) {
            return String((current && current.salesOrderNo) || '').trim();
        }

        // "3-55.2-06.74-075_001_A_0.pdf" -> "3-55.2-06.74-075"
        function drawingNoFromResult(v) {
            if (v.cislo_vykresu) return v.cislo_vykresu;
            if (v.kluc && v.kluc_typ === 'vykres') return v.kluc;
            return String(v.nazov || '').replace(/\.(pdf|tiff?)$/i, '').split('_')[0];
        }

        // Ked sluzba vrati viac suborov, najrelevantnejsi je ten, ktory lezi
        // v priecinku tejto zakazky a ktoremu sedi revizia so SAP verziou.
        function bestResult(list) {
            return list.find((v) => v.v_zakazke && v.zhoda_revizie === true) ||
                   list.find((v) => v.v_zakazke) ||
                   list.find((v) => v.zhoda_revizie === true) ||
                   list[0] || null;
        }

        let lookupToken = 0;

        /*
         * Cislo vykresu sa zistuje v dvoch krokoch:
         *   1. z Excelu podla cisla vyrobnej zakazky (ak je Excel nacitany)
         *   2. inak priamo zo sluzby vykresov podla cisla materialu, ktore
         *      PDA uz pozna - Excel teda nie je podmienkou
         * Zakaznicka zakazka ide do parametra "path", vdaka comu sluzba oznaci
         * vykresy lezice priamo v tejto zakazke (v_zakazke = true).
         */
        async function applyForOperation(current) {
            const token = ++lookupToken;

            const fromExcel = drawingIndex[normalizeOrderKey(current.productionOrderNo)];
            const excelDrawingNo = fromExcel ? stripLeadingApostrophe(fromExcel.drawingNo) : '';
            if (excelDrawingNo) {
                updateDisplay({
                    drawingNo: excelDrawingNo,
                    version: stripLeadingApostrophe(fromExcel.version) || '',
                    searchTerm: excelDrawingNo,
                    source: 'excel',
                });
                return;
            }

            updateDisplay({ drawingNo: '…', version: '', searchTerm: '', source: 'hladam' });

            // 2) sluzba vykresov pozna denny export -> kluc priamo pre tuto vyrobnu zakazku
            try {
                const cislo = normalizeOrderKey(current.productionOrderNo);
                const local = orderMap && cislo ? orderMap[cislo] : null;
                const z = local
                    ? { kluc: local.kluc, typ: local.typ, revizia: local.revizia, pocet: 0 }
                    : await pdmOrderLookup(current.productionOrderNo);
                if (token !== lookupToken) return;
                if (z && z.kluc) {
                    const isDrawing = z.typ !== 'material';
                    updateDisplay({
                        drawingNo: isDrawing ? z.kluc : 'SAP ' + z.kluc,
                        version: isDrawing ? (z.revizia || '') : '',
                        searchTerm: z.kluc,
                        source: 'server',
                        count: z.pocet || 0,
                    });
                    return;
                }
            } catch (e) {
                if (token !== lookupToken) return;
                console.log(LOG, 'služba nepozná zákazku (alebo ešte nemá /zakazka), skúšam materiál:', e.message);
            }

            // 3) zaloha: material - najprv z Excelu (ak zakazku pozna, ale vykres tam nema), inak z operacie
            const material = ((fromExcel && fromExcel.materialNo) || String(current.materialNo || '')).trim();
            if (!material) { updateDisplay(null); return; }

            updateDisplay({ drawingNo: '…', version: '', searchTerm: material, source: 'hladam' });

            try {
                const d = await pdmSearch(material, { path: salesOrderOf(current) });
                if (token !== lookupToken) return; // medzitym sa otvorila ina operacia

                const best = bestResult(d.vysledky || []);
                if (!best) { updateDisplay(null); return; }

                updateDisplay({
                    drawingNo: drawingNoFromResult(best),
                    version: best.revizia || '',
                    searchTerm: material,
                    source: 'pdm',
                    count: d.pocet || 0,
                });
            } catch (e) {
                if (token !== lookupToken) return;
                console.warn(LOG, 'hľadanie výkresu zlyhalo', e);
                updateDisplay(null);
            }
        }

        function updateDisplay(info) {
            const valueEl = document.getElementById('__pda_order_drawing_value__');
            const revisionEl = document.getElementById('__pda_order_drawing_revision__');
            if (!valueEl || !revisionEl) return;

            if (!info) {
                currentDrawingInfo = null;
                valueEl.textContent = '—';
                revisionEl.textContent = '';
                return;
            }

            currentDrawingInfo = info;
            valueEl.textContent = info.drawingNo || '—';

            if (info.source === 'hladam') {
                revisionEl.textContent = 'hľadám…';
            } else if (info.version) {
                revisionEl.textContent = 'rev. ' + info.version + (info.count > 1 ? ' · ' + info.count + ' súb.' : '');
            } else {
                revisionEl.textContent = info.count > 1 ? info.count + ' súbory' : '';
            }
        }

        // --- sluzba "Mapa vykresov" (PDM) ---

        /*
         * Vyrobna zakazka -> kluc z denneho exportu. Sluzba vykresov cita AutomatedOQ180.xlsx
         * raz denne a ku kazdemu vykresu drzi zoznam zakaziek; endpoint /zakazka/<cislo>
         * to vrati jednym dopytom. 404 = zakazka v exporte nie je.
         */
        function pdmOrderLookup(productionOrderNo) {
            const cislo = normalizeOrderKey(productionOrderNo);
            if (!cislo) return Promise.resolve(null);
            return new Promise((resolve, reject) => {
                GM_xmlhttpRequest({
                    method: 'GET',
                    url: settings.pdm.base + '/zakazka/' + encodeURIComponent(cislo),
                    headers: settings.pdm.key ? { 'X-API-Key': settings.pdm.key } : {},
                    timeout: 15000,
                    onload: (r) => {
                        if (r.status === 404) { resolve(null); return; }
                        if (r.status < 200 || r.status >= 300) { reject(new Error('HTTP ' + r.status)); return; }
                        try { resolve(JSON.parse(r.responseText)); }
                        catch (e) { reject(new Error('neplatná odpoveď služby')); }
                    },
                    onerror: () => reject(new Error('spojenie so službou zlyhalo')),
                    ontimeout: () => reject(new Error('služba neodpovedala včas')),
                });
            });
        }

        function pdmSearch(cislo, { path = '', live = false } = {}) {
            return new Promise((resolve, reject) => {
                const params = new URLSearchParams({ q: cislo });
                if (path) params.set('path', path);
                if (live) params.set('live', '1');
                GM_xmlhttpRequest({
                    method: 'GET',
                    url: settings.pdm.base + '/search?' + params,
                    headers: settings.pdm.key ? { 'X-API-Key': settings.pdm.key } : {},
                    onload: (response) => {
                        if (response.status < 200 || response.status >= 300) {
                            reject(new Error('Služba výkresov vrátila HTTP ' + response.status));
                            return;
                        }
                        try { resolve(JSON.parse(response.responseText)); }
                        catch (e) { reject(new Error('Neplatná odpoveď zo služby výkresov')); }
                    },
                    onerror: () => reject(new Error('Služba výkresov neodpovedala (chyba spojenia)')),
                    ontimeout: () => reject(new Error('Služba výkresov neodpovedala včas')),
                });
            });
        }

        /*
         * Zhoda revizie z Excelu s reviziou v nazve suboru - rovnake pravidlo ako
         * v sluzbe (Python verzia): pismeno musi sediet, index len ak ho poznaju
         * obaja. 'BC' ~ 'C' ano, 'BC' ~ 'B' nie, '0A' ~ 'A' ano.
         */
        function revisionMatches(label, rev) {
            label = String(label || '').toUpperCase();
            rev = String(rev || '').toUpperCase();
            if (!label || !rev || label === 'BEZ REV.') return false;
            const l1 = label.slice(-1), i1 = label.length > 1 ? label.slice(0, -1) : '';
            const l2 = rev.slice(-1), i2 = rev.length > 1 ? rev.slice(0, -1) : '';
            return l1 === l2 && (!i1 || !i2 || i1 === i2);
        }

        // Hlada sa LEN podla cisla vykresu (bez revizie); revizia z Excelu sa
        // v zozname iba zvyrazni, rozhodnutie ostava na cloveku.
        function pdmOpenDialog(cislo, { path = '', rev = '', titul = '' } = {}) {
            const overlay = document.createElement('div');
            overlay.id = '__pda_pdm_overlay__';
            overlay.style.cssText =
                'position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:99999;' +
                'display:flex;align-items:center;justify-content:center;font-family:inherit;';

            overlay.innerHTML =
                '<div style="background:#fff;border-radius:10px;max-width:900px;width:92%;max-height:80vh;display:flex;flex-direction:column;overflow:hidden;box-shadow:0 8px 30px rgba(0,0,0,.3)">' +
                '<div style="display:flex;justify-content:space-between;align-items:center;padding:12px 16px;background:#5b9bd5;color:#fff">' +
                // v hlavicke je to, co vidi operator v okienku VYKRES (cislo vykresu
                // + revizia); hlada sa stale podla `cislo` (moze to byt aj material)
                '<b>Výkresy — ' + (titul || cislo) + '</b>' +
                '<button data-pda-zavrit type="button" style="background:none;border:0;color:#fff;font-size:22px;line-height:1;cursor:pointer">×</button>' +
                '</div><div data-pda-telo style="padding:14px 16px;overflow:auto">Hľadám…</div></div>';

            const close = () => { overlay.remove(); document.removeEventListener('keydown', onEsc); };
            const onEsc = (e) => { if (e.key === 'Escape') close(); };

            overlay.querySelector('[data-pda-zavrit]').addEventListener('click', close);
            overlay.addEventListener('click', (e) => { if (e.target === overlay) close(); });
            document.addEventListener('keydown', onEsc);
            document.body.appendChild(overlay);

            const telo = overlay.querySelector('[data-pda-telo]');

            const TH = 'style="text-align:left;padding:6px;border-bottom:2px solid #5b9bd5"';
            const TD = 'style="padding:6px;border-bottom:1px solid #eee"';

            function run(live) {
                telo.innerHTML = live
                    ? '<p>Hľadám naživo v PDM… (môže to trvať pár sekúnd)</p>'
                    : '<p>Hľadám…</p>';

                pdmSearch(cislo, { path, live }).then((d) => {
                    if (!d.pocet) {
                        telo.innerHTML =
                            '<p>Pre <b>' + cislo + '</b> sa nenašiel žiadny PDF ani TIFF výkres.</p>' +
                            (live ? '' :
                                '<p style="color:#777;font-size:12px">Denná mapa sa obnovuje raz za deň — čerstvo pridaný výkres v nej ešte nemusí byť.</p>' +
                                '<button data-pda-live type="button" style="background:#5b9bd5;color:#fff;border:0;border-radius:7px;padding:8px 15px;cursor:pointer;font:inherit;font-size:.88rem">Hľadať naživo v PDM</button>');
                        const liveBtn = telo.querySelector('[data-pda-live]');
                        if (liveBtn) liveBtn.addEventListener('click', () => run(true));
                        return;
                    }

                    // zhoda revizie: primarne zo sluzby (pocita ju zo SAP verzie), inak lokalne z Excelu
                    const rows = (d.vysledky || []).map((v) => {
                        const known = v.zhoda_revizie === true || v.zhoda_revizie === false;
                        return Object.assign({}, v, { zhoda: known ? v.zhoda_revizie : (rev ? revisionMatches(v.revizia, rev) : null) });
                    });
                    rows.sort((a, b) => ((b.zhoda === true) - (a.zhoda === true)) || String(a.nazov).localeCompare(String(b.nazov)));

                    telo.innerHTML =
                        '<table style="width:100%;border-collapse:collapse;font-size:14px"><thead><tr>' +
                        '<th ' + TH + '>Názov</th><th ' + TH + '>Revízia</th>' +
                        '<th ' + TH + '>Stav</th><th ' + TH + '>Ver.</th>' +
                        '</tr></thead><tbody>' +
                        rows.map((v) => {
                            const bg = v.zhoda === true ? '#e6f4ea' : (v.v_zakazke ? '#f3f7fd' : 'transparent');
                            const mark = v.v_zakazke ? ' <span style="font-size:10px;color:#2f6fbf;border:1px solid #b9cbe8;border-radius:9px;padding:1px 6px">v zákazke</span>' : '';
                            return '<tr data-id="' + v.id + '" style="cursor:pointer;background:' + bg + '">' +
                                '<td ' + TD + '><b>' + v.nazov + '</b>' + mark + '</td>' +
                                '<td ' + TD + '>' + (v.revizia || '') + '</td>' +
                                '<td ' + TD + '>' + (v.stav || '') + '</td>' +
                                '<td ' + TD + '>' + (v.verzia != null ? v.verzia : '') + '</td></tr>';
                        }).join('') +
                        '</tbody></table>' +
                        '<p data-pda-info style="color:#777;font-size:12px;margin-top:10px">' + d.pocet + ' výkresov · ' +
                        (d.zdroj === 'mapa' ? 'z dennej mapy' : 'naživo z PDM') +
                        (d.mapa_z ? ' (' + d.mapa_z + ')' : '') +
                        ' · zelený riadok = revízia sedí so SAP · kliknutím sa výkres otvorí</p>';

                    telo.querySelectorAll('tr[data-id]').forEach((tr) => {
                        tr.addEventListener('click', () => {
                            const info = telo.querySelector('[data-pda-info]');
                            // vydaj suboru trva 0,4-4 s, treba to dat najavo
                            if (info) info.textContent = 'Otváram výkres… (ťahá sa z PDM, môže to chvíľu trvať)';
                            W.open(settings.pdm.base + '/file/' + tr.dataset.id, '_blank');
                        });
                    });
                }).catch((err) => {
                    telo.innerHTML = '<p style="color:#b00">Služba výkresov neodpovedala.<br>' + err.message +
                        '<br><span style="color:#777;font-size:12px">Adresa: ' + settings.pdm.base + '</span></p>';
                });
            }

            run(false);
        }

        // --- tlacidla ---

        /*
         * Nadpis okna s vykresmi = to iste, co je v okienku VYKRES vpravo hore
         * ("1-60.2-06.66-009 rev. C"). Predtym tam bolo cislo, podla ktoreho sa
         * hlada - pri hladani podla materialu teda cislo materialu, co operatora
         * mylilo. Hladanie sa nemeni, meni sa len text v hlavicke.
         */
        function dialogTitle() {
            if (!currentDrawingInfo) return '';
            const cislo = String(currentDrawingInfo.drawingNo || '').trim();
            if (!cislo || cislo === '—' || cislo === '…') return '';
            const rev = String(currentDrawingInfo.version || '').trim();
            return rev ? cislo + ' rev. ' + rev : cislo;
        }

        /*
         * Hlasenia o Exceli sa uz operatorovi neukazuju. Vykres sa najde aj bez
         * Excelu (cez sluzbu PDM podla materialu), takze cervene "Excel sa
         * nestiahol - skus znova" pri otvorenej zakazke len zavadzalo. Jedina
         * vynimka je stav "needs-permission": tam prehliadac ziada kliknutie
         * cloveka, bez tlacidla by sa Excel nedal nacitat vobec.
         * Ostatne stavy idu uz len do konzoly.
         */
        function renderLoadButtonState(loadButton, state) {
            if (state !== 'needs-permission') {
                loadButton.style.display = 'none';
                if (state !== 'loaded' && state !== 'nofile' && state !== 'loading') {
                    console.log(LOG, 'stav Excelu:', state, '(hlasenie sa operatorovi nezobrazuje)');
                }
                return;
            }
            loadButton.style.display = '';

            const states = {
                loading: ['Načítavam…', '#f5f5f5', '#ccc'],
                'needs-permission': ['Povoliť prístup k Excelu', '#fff4e5', '#f9a825'],
                error: ['Chyba, skús znova', '#fdecea', '#e53935'],
                'url-error': ['Excel sa nestiahol — skús znova', '#fdecea', '#e53935'],
                'url-invalid': ['Adresa Excelu je nezrozumiteľná — klik otvorí nastavenia', '#fff4e5', '#f9a825'],
                'file-denied': ['Excel z disku sa nenačítal — zapni prístup k súborom v Tampermonkey', '#fdecea', '#e53935'],
                unsupported: ['Prehliadač nepodporuje zapamätanie', '#fdecea', '#e53935'],
                // Excel uz nie je podmienkou - vykres sa najde aj podla materialu
                nofile: ['Excel (nepovinné)', '#f5f5f5', '#ccc'],
            };
            const [text, bg, border] = states[state] || states.nofile;
            loadButton.textContent = text;
            loadButton.style.backgroundColor = bg;
            loadButton.style.borderColor = border;
        }

        function updateLoadButtonState(state, handle) {
            currentLoadState = state;
            pendingHandle = state === 'needs-permission' ? handle : null;
            const loadButton = document.getElementById(LOAD_BUTTON_ID);
            if (loadButton) renderLoadButtonState(loadButton, state);
        }

        function buildWrapper() {
            const wrapper = document.createElement('div');
            wrapper.id = WRAPPER_ID;
            wrapper.style.cssText = 'display:flex;justify-content:flex-end;align-items:center;width:100%;box-sizing:border-box;margin-bottom:8px;';

            const loadButton = document.createElement('button');
            loadButton.id = LOAD_BUTTON_ID;
            loadButton.type = 'button';
            loadButton.style.cssText = 'padding:6px 10px;border:1px solid #ccc;border-radius:6px;background:#f5f5f5;cursor:pointer;font-size:0.75rem;margin-right:8px;align-self:center;';
            loadButton.addEventListener('click', () => {
                const src = excelSource();
                if (src.kind === 'invalid') openSettings();
                else if (src.kind === 'http' || src.kind === 'file') loadExcelFromUrl(src.url);
                else if (pendingHandle) confirmPermissionAndLoad(pendingHandle);
                else pickFileAndRemember();
            });

            const button = document.createElement('button');
            button.id = BUTTON_ID;
            button.type = 'button';
            button.style.cssText = 'display:flex;flex-direction:column;align-items:flex-end;gap:2px;padding:6px 12px;border:1px solid #5b9bd5;border-radius:6px;background:#fff;cursor:pointer;width:150px;flex-shrink:0;margin-right:8px;';

            const label = document.createElement('span');
            label.textContent = 'VÝKRES';
            label.style.cssText = 'font-size:0.65rem;color:#888;letter-spacing:0.05em;';

            const value = document.createElement('span');
            value.id = '__pda_order_drawing_value__';
            value.textContent = '—';
            value.style.cssText = 'font-size:0.9rem;font-weight:bold;color:#000;';

            const revision = document.createElement('span');
            revision.id = '__pda_order_drawing_revision__';
            revision.style.cssText = 'font-size:0.75rem;color:#555;';

            button.appendChild(label);
            button.appendChild(value);
            button.appendChild(revision);
            button.addEventListener('click', () => {
                const path = salesOrderOf(shared.currentOperation);

                // ak uz vieme cislo vykresu (alebo aspon material), hladame podla neho
                if (currentDrawingInfo && currentDrawingInfo.searchTerm) {
                    pdmOpenDialog(currentDrawingInfo.searchTerm, {
                        path,
                        rev: (currentDrawingInfo.source === 'excel' || currentDrawingInfo.source === 'server')
                            ? currentDrawingInfo.version : '',
                        titul: dialogTitle(),
                    });
                    return;
                }
                // inak skusime aspon cislo materialu z otvorenej operacie
                const material = shared.currentOperation && String(shared.currentOperation.materialNo || '').trim();
                if (material) pdmOpenDialog(material, { path, titul: dialogTitle() });
                else console.log(LOG, 'pre aktuálnu zákazku nie je známe ani číslo výkresu, ani materiálu');
            });

            wrapper.appendChild(loadButton);
            wrapper.appendChild(button);
            return wrapper;
        }

        function ensureButton() {
            const container = document.getElementById(CONTAINER_ID);
            if (!container || document.getElementById(WRAPPER_ID)) return;

            container.insertBefore(buildWrapper(), container.firstChild);
            const loadButton = document.getElementById(LOAD_BUTTON_ID);
            if (loadButton) renderLoadButtonState(loadButton, currentLoadState);
            if (shared.currentOperation) applyForOperation(shared.currentOperation);
        }

        // okno so zoznamom vykresov sprístupnime aj ostatnym modulom
        // (pouziva ho tlacidlo Vykresy v pravom paneli HF Slovakia)
        shared.pdmOpenDialog = pdmOpenDialog;

        DomWatch.add(ensureButton);
        onReady(() => {
            // ak je v nastaveniach adresa, tahame odtial automaticky;
            // inak sa subor vybera rucne a prehliadac si ho pamata
            const src = excelSource();
            if (src.kind === 'http' || src.kind === 'file') {
                loadExcelFromUrl(src.url);
            } else if (src.kind === 'invalid') {
                console.warn(LOG, 'adresa Excelu je nezrozumiteľná, ignorujem:', src.raw);
                updateLoadButtonState('url-invalid');
            } else {
                tryAutoLoad();
            }
        });
    }

    /* ------------------ 3.8 Prehlad pracovisk (uvodna) ------------------ */

    /*
     * Portovane z Python appky (Temporary/PDA/app/pda_action.py, blok OBLASTI).
     * Rozdiel oproti prvej verzii tohto modulu:
     *  - data sa NEcitaju zo siete ani z DOM, ale priamo z premennej appky
     *    getGlobals().getVar('aWorkcenterOperationLists')
     *  - kategoria = pole `area` (Assembly / Welding / Machining / Quality Control)
     *  - "vyraba" = pole `count` > 0 (kolko ludi je prihlasenych na pracovisku)
     *  - klik na stroj vola ten isty handler, aky vola appka pri kliku na kartu
     *  - povodne karty sa skryju jednym CSS pravidlom, v DOM ostavaju
     */
    function modWorkcenterOverview() {
        const OVERVIEW_ID = '__pda_overview__';
        const STYLE_ID = '__pda_overview_styles__';
        const BODY_CLASS = 'pda-overview-on';
        const HOME_BUTTON_ID = 'Main--Button_HomeScreen';
        const PANEL_ID = 'Main--Workcenter_Panel';
        const ORDER = ['Assembly', 'Welding', 'Machining', 'Quality Control'];
        const SUBTITLES = {
            'Assembly': 'Finálna montáž a podzostavy',
            'Welding': 'Zváranie a príprava',
            'Machining': 'CNC a konvenčné obrábanie',
            'Quality Control': 'Kontrola a meranie',
            'Ostatné': 'Nezaradené pracoviská',
        };
        const ICONS = { 'Assembly': '🔧', 'Welding': '🔥', 'Machining': '⚙', 'Quality Control': '🔎' };

        const expanded = new Set();   // rozbalene kategorie
        const filters = {};           // { q, only } pre kazdu kategoriu
        let lastSignature = null;

        function globals() {
            try {
                const main = W.sap.ui.getCore().byId('Main');
                return main ? main.getController().getGlobals() : null;
            } catch (e) { return null; }
        }

        function workcenterList() {
            const g = globals();
            if (!g) return [];
            try { return g.getVar('aWorkcenterOperationLists') || []; } catch (e) { return []; }
        }

        function userName() {
            const g = globals();
            try {
                const n = g && g.getVar('oUser', 'username');
                if (n) return String(n);
            } catch (e) { /* ignore */ }
            const el = document.querySelector('[id*="Label_Username"][id$="-bdi"]');
            return el ? el.textContent.trim() : '';
        }

        function normalize(s) {
            return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase();
        }

        /*
         * Otvorenie pracoviska presne tak, ako to robi appka po kliku na kartu
         * (handler f_cards_Header3_press). V Python verzii sa zistilo, ze handler
         * po navTo nedobehne - texty casov ostanu prazdne a potvrdzovacie tlacidla
         * zapnute - preto sa to po 1,5 s dorobi rovnako, ako to robi appka.
         */
        function openWorkcenter(wc) {
            try {
                const core = W.sap.ui.getCore();
                const main = core.byId('Main').getController();
                const g = main.getGlobals();
                const app = core.byId('Application');

                if (app && app.getCurrentPage && String(app.getCurrentPage().getId()).indexOf('WorkcenterDetail') === 0) {
                    main.navTo('Main', 'show');
                }
                if (!g.getVar('oSelectedWorkcenterOperation')) g.setVar('oSelectedWorkcenterOperation', {});

                const ctx = { getObject: () => wc };
                const src = { getParent: () => ({ getBindingContext: () => ctx }) };
                main.f_cards_Header3_press({ getSource: () => src });

                setTimeout(() => {
                    try {
                        const op = g.getVar('oSelectedWorkcenterOperation') || {};
                        [['SetupTime_Text', 'setupTimeConf', 'setupTime'],
                         ['MachineTime_Text', 'machineTimeConf', 'machineTime'],
                         ['LaborTime_Text', 'laborTimeConf', 'laborTime']].forEach(([id, conf, plan]) => {
                            const e = core.byId('WorkcenterDetail--' + id);
                            if (e && typeof e.setText === 'function') {
                                e.setText(op.id ? op[conf] + ' / ' + op[plan] + ' min' : '- / - min');
                            }
                        });
                        if (!op.id) {
                            main.setItemValue('WorkcenterDetail', 'Part_Confirm_Button', 'enabled', false);
                            main.setItemValue('WorkcenterDetail', 'Confirm_Button', 'enabled', false);
                        }
                    } catch (e) { /* kozmetika, nie je kriticka */ }
                }, 1500);
            } catch (e) {
                console.warn(LOG, 'otvorenie pracoviska cez handler zlyhalo, skúšam klik na kartu', e);
                const tile = findTile(wc);
                if (tile) pressElement(tile);
            }
        }

        // zaloha: povodna karta v DOM podla nazvu pracoviska
        function findTile(wc) {
            const want = String(wc.workcenterDescription || '').trim();
            const tiles = document.querySelectorAll('[id^="Main--Workcenter_Toolbar-Main--ui_layout_Grid3-"]');
            for (const t of tiles) {
                const suffix = t.id.replace('Main--Workcenter_Toolbar-', '');
                const el = document.getElementById('Main--WorkcenterDescription_Label-' + suffix + '-bdi');
                if (el && el.textContent.trim() === want) return t;
            }
            return null;
        }

        function injectStyles() {
            if (document.getElementById(STYLE_ID)) return;
            const st = document.createElement('style');
            st.id = STYLE_ID;
            st.textContent = `
body.${BODY_CLASS} #${PANEL_ID} .sapMPanelContent > :not(#${OVERVIEW_ID}) { display: none !important; }
#${OVERVIEW_ID} { font: 14px/1.5 -apple-system,"Segoe UI",Roboto,sans-serif; color:#1a2233; padding: 4px 0 10px; }
#${OVERVIEW_ID} .ov-hi { padding: 4px 6px 14px; }
#${OVERVIEW_ID} .ov-hi .k { font-size:.68rem; letter-spacing:.12em; color:#5b7cb5; font-weight:700; }
#${OVERVIEW_ID} .ov-hi .n { font-size:1.7rem; font-weight:700; margin:2px 0 1px; }
#${OVERVIEW_ID} .ov-hi .s { color:#6b7180; font-size:.87rem; }
#${OVERVIEW_ID} .ov-g { border:1px solid #dfe4ec; border-radius:12px; background:#fff; margin-bottom:10px; overflow:hidden; }
#${OVERVIEW_ID} .ov-g.open { border-color:#b9cbe8; box-shadow:0 1px 6px rgba(40,70,130,.08); }
#${OVERVIEW_ID} .ov-gh { display:flex; align-items:center; gap:14px; width:100%; background:none; border:0;
  padding:14px 16px; cursor:pointer; text-align:left; font:inherit; color:inherit; }
#${OVERVIEW_ID} .ov-g.open .ov-gh { background:#f2f6fc; }
#${OVERVIEW_ID} .ov-ic { width:42px; height:42px; border-radius:9px; background:#eaf1fb; flex-shrink:0;
  display:flex; align-items:center; justify-content:center; font-size:21px; }
#${OVERVIEW_ID} .ov-gt { flex:1 1 auto; min-width:0; }
#${OVERVIEW_ID} .ov-gt b { font-size:1.05rem; display:block; }
#${OVERVIEW_ID} .ov-gt span { color:#6b7180; font-size:.85rem; }
#${OVERVIEW_ID} .ov-num { text-align:center; flex-shrink:0; min-width:74px; }
#${OVERVIEW_ID} .ov-num b { display:block; font-size:1.3rem; color:#1d4ed8; line-height:1.1; }
#${OVERVIEW_ID} .ov-num span { font-size:.62rem; letter-spacing:.09em; color:#8b93a3; }
#${OVERVIEW_ID} .ov-num.on b { color:#2f7d43; }
#${OVERVIEW_ID} .ov-ch { color:#9aa3b4; font-size:18px; flex-shrink:0; }
#${OVERVIEW_ID} .ov-gb { padding:4px 16px 16px; border-top:1px solid #e8edf4; }
#${OVERVIEW_ID} .ov-tools { display:flex; gap:8px; justify-content:flex-end; margin:12px 0; flex-wrap:wrap; }
#${OVERVIEW_ID} .ov-tools input, #${OVERVIEW_ID} .ov-tools select {
  padding:6px 10px; border:1px solid #ccd4e0; border-radius:7px; font:inherit; font-size:.85rem; }
#${OVERVIEW_ID} .ov-tools input { width:190px; }
#${OVERVIEW_ID} .ov-chips { display:grid; grid-template-columns:repeat(auto-fill,minmax(215px,1fr)); gap:7px; }
#${OVERVIEW_ID} .ov-chip { display:flex; align-items:center; justify-content:space-between; gap:10px;
  border:1px solid #dfe4ec; border-radius:9px; background:#fff; padding:8px 11px; cursor:pointer;
  font:inherit; text-align:left; transition:border-color .12s, background .12s; }
#${OVERVIEW_ID} .ov-chip:hover { border-color:#5b8def; background:#f6f9ff; }
#${OVERVIEW_ID} .ov-chip b { font-size:.93rem; }
#${OVERVIEW_ID} .ov-txt { display:flex; flex-direction:column; min-width:0; }
#${OVERVIEW_ID} .ov-txt small { font-size:.71rem; color:#6b7180; line-height:1.3;
  white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
#${OVERVIEW_ID} .ov-st { display:flex; align-items:center; gap:5px; font-size:.76rem; color:#1e40af;
  background:#e8f0fe; border-radius:999px; padding:3px 9px; white-space:nowrap; font-weight:600; }
#${OVERVIEW_ID} .ov-dot { width:7px; height:7px; border-radius:50%; background:#2563eb; flex-shrink:0; }
#${OVERVIEW_ID} .ov-chip.run .ov-st { color:#166534; background:#eaf6ee; }
#${OVERVIEW_ID} .ov-chip.run .ov-dot { background:#16a34a; }
#${OVERVIEW_ID} .ov-empty { color:#8b93a3; font-size:.85rem; padding:8px 2px; }
`;
            document.head.appendChild(st);
        }

        function chip(w) {
            const producing = (w.count || 0) > 0;
            const code = String(w.workcenter || '');
            const name = String(w.workcenterDescription || '');

            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'ov-chip' + (producing ? ' run' : '');
            btn.title = code + ' — ' + name;

            const txt = document.createElement('span');
            txt.className = 'ov-txt';
            const label = document.createElement('b');
            label.textContent = code || name;
            txt.appendChild(label);
            if (code && name) {
                const sub = document.createElement('small');
                sub.textContent = name.length > 20 ? name.slice(0, 20).trim() + '…' : name;
                txt.appendChild(sub);
            }

            const st = document.createElement('span');
            st.className = 'ov-st';
            const dot = document.createElement('span');
            dot.className = 'ov-dot';
            st.appendChild(dot);
            st.appendChild(document.createTextNode(producing ? 'Vyrába' : 'Nevyrába'));

            btn.appendChild(txt);
            btn.appendChild(st);
            btn.addEventListener('click', () => openWorkcenter(w));
            return btn;
        }

        function groupBody(groupName, items) {
            const state = filters[groupName] || (filters[groupName] = { q: '', only: 'all' });

            const body = document.createElement('div');
            body.className = 'ov-gb';

            const tools = document.createElement('div');
            tools.className = 'ov-tools';
            const q = document.createElement('input');
            q.type = 'search';
            q.placeholder = 'Hľadať pracovisko…';
            q.value = state.q;
            const sel = document.createElement('select');
            sel.innerHTML = '<option value="all">Všetky</option><option value="run">Vyrába</option><option value="idle">Nevyrába</option>';
            sel.value = state.only;
            tools.appendChild(q);
            tools.appendChild(sel);

            const chips = document.createElement('div');
            chips.className = 'ov-chips';

            function paint() {
                chips.innerHTML = '';
                const needle = normalize(state.q);
                const shown = items.filter((w) => {
                    const producing = (w.count || 0) > 0;
                    if (state.only === 'run' && !producing) return false;
                    if (state.only === 'idle' && producing) return false;
                    if (!needle) return true;
                    return normalize(w.workcenter + ' ' + w.workcenterDescription).indexOf(needle) !== -1;
                });
                if (shown.length === 0) {
                    const empty = document.createElement('div');
                    empty.className = 'ov-empty';
                    empty.textContent = 'Nič nezodpovedá zadaniu.';
                    chips.appendChild(empty);
                    return;
                }
                shown.forEach((w) => chips.appendChild(chip(w)));
            }

            q.addEventListener('input', () => { state.q = q.value; paint(); });
            q.addEventListener('click', (e) => e.stopPropagation());
            sel.addEventListener('change', () => { state.only = sel.value; paint(); });

            body.appendChild(tools);
            body.appendChild(chips);
            paint();
            return body;
        }

        function render(host, list) {
            host.innerHTML = '';

            const hi = document.createElement('div');
            hi.className = 'ov-hi';
            hi.innerHTML = '<div class="k">VITAJTE SPÄŤ</div><div class="n"></div>' +
                '<div class="s">Vyberte pracovisko a začnite pracovať.</div>';
            hi.querySelector('.n').textContent = userName() || 'Pracoviská';
            host.appendChild(hi);

            // zoskupenie podla pola `area`; co ho nema, ide do "Ostatne"
            const groups = {};
            list.forEach((w) => {
                const key = w.area || 'Ostatné';
                (groups[key] = groups[key] || []).push(w);
            });
            const order = ORDER.filter((g) => groups[g])
                .concat(Object.keys(groups).filter((g) => ORDER.indexOf(g) === -1).sort());

            order.forEach((gName) => {
                const items = groups[gName].slice().sort((a, b) => String(a.workcenter).localeCompare(String(b.workcenter)));
                const online = items.reduce((acc, w) => acc + (w.count || 0), 0);
                const isOpen = expanded.has(gName);

                const box = document.createElement('div');
                box.className = 'ov-g' + (isOpen ? ' open' : '');

                const head = document.createElement('button');
                head.type = 'button';
                head.className = 'ov-gh';
                head.innerHTML =
                    '<div class="ov-ic">' + (ICONS[gName] || '🏭') + '</div>' +
                    '<div class="ov-gt"><b></b><span></span></div>' +
                    '<div class="ov-num"><b>' + items.length + '</b><span>PRACOVÍSK</span></div>' +
                    '<div class="ov-num on"><b>' + online + '</b><span>ONLINE</span></div>' +
                    '<div class="ov-ch">' + (isOpen ? '⌄' : '›') + '</div>';
                head.querySelector('.ov-gt b').textContent = gName;
                head.querySelector('.ov-gt span').textContent = SUBTITLES[gName] || '';
                head.addEventListener('click', () => {
                    if (expanded.has(gName)) expanded.delete(gName); else expanded.add(gName);
                    lastSignature = null;
                    ensureOverview();
                });
                box.appendChild(head);
                if (isOpen) box.appendChild(groupBody(gName, items));
                host.appendChild(box);
            });
        }

        function ensureOverview() {
            if (!document.getElementById(HOME_BUTTON_ID)) return;   // len uvodna obrazovka

            const panel = document.getElementById(PANEL_ID);
            const content = panel && panel.querySelector('.sapMPanelContent');
            if (!content) return;

            const list = workcenterList();
            if (!list.length) return;   // appka este taha data zo SAP (~25 s)

            injectStyles();
            document.body.classList.add(BODY_CLASS);

            let host = document.getElementById(OVERVIEW_ID);
            if (!host) {
                host = document.createElement('div');
                host.id = OVERVIEW_ID;
                content.insertBefore(host, content.firstChild);
            } else if (host.parentElement !== content) {
                content.insertBefore(host, content.firstChild);
            }

            // prekreslit len ked sa nieco zmenilo (clovek, pocty, rozbalenie)
            const signature = userName() + '#' +
                list.map((w) => w.workcenter + ':' + (w.count || 0)).join('|') + '#' +
                Array.from(expanded).join(',');
            if (signature === lastSignature) return;
            lastSignature = signature;

            render(host, list);
        }

        DomWatch.add(ensureOverview);
        // `count` sa meni bez zmeny DOM, preto aj casovac (rovnako ako Python verzia)
        onReady(() => setInterval(ensureOverview, 1500));
    }

    /* ---------- 3.9 Pracovny zoznam: zakazky ako kompaktne pilulky ---------- */

    /*
     * Lavy zoznam zakaziek v detaile pracoviska mal kazdu zakazku na styroch
     * riadkoch (~101 px) a okienko pevnu vysku 20em, takze na obrazovke boli
     * vidiet 2-3 zakazky a pod zoznamom ostavalo prazdne miesto az po spodok
     * stranky. Modul robi dve veci:
     *   1. kazdu polozku prekresli na kompaktnu "pilulku" (dva riadky, ~60 px)
     *   2. okienko so zoznamom natiahne az k spodnemu okraju obrazovky
     *
     * Povodny obsah polozky sa iba SKRYJE (CSS), v DOM ostava - klik, oznacenie,
     * prepinac VSET aj vyhladavanie appky funguju dalej bez zmeny. Udaje sa
     * citaju z modelu appky (binding context polozky), nie z textu na obrazovke,
     * takze sa nic neparsuje a nic sa nemoze "netrafit".
     */
    function modOrderListPills() {
        const LIST_ID = 'WorkcenterDetail--Work_List';
        const SCROLL_ID = 'WorkcenterDetail--WorkList_ScrollContainer';
        const LEFT_ID = 'WorkcenterDetail--LeftColumn_FlexBox';
        const STYLE_ID = '__pda_orderlist_styles__';
        const TIP_ID = '__pda_pill_tip__';
        const BOX1_ID = '__pda_left_box_zoznam__';
        const BOX2_ID = '__pda_left_box_graf__';
        const VYSKA_ZOZNAMU = 460;    // pevna vyska posuvneho zoznamu zakaziek (px) - v 3J uz len zaloha
        const MIN_HEIGHT = 260;       // 3J: nizsi zoznam nikdy nebude, ani v malom okne

        let poslednyPodpis = '';
        let poslednyResize = 0;

        function injectStyles() {
            if (document.getElementById(STYLE_ID)) return;
            const st = document.createElement('style');
            st.id = STYLE_ID;
            st.textContent = `
/* Lavy stlpec: uzsi (pilulky maju sirku podla obsahu, siroky box nemal co
   vyplnit) a rozdeleny na dva samostatne zaoblene boxy - zoznam zakaziek
   a graf. Samotny stlpec uz teda ziadny ram nema. */
#${LEFT_ID} { background:transparent !important; border:0 !important; box-shadow:none !important;
  padding:0 !important; box-sizing:border-box;
  width:430px !important; max-width:430px !important; min-width:0 !important; flex:0 0 430px !important; }
.pda-left-box { background:#fff; border:1px solid #dfe4ec; border-radius:14px; padding:8px;
  margin-bottom:10px; box-sizing:border-box; box-shadow:0 1px 4px rgba(16,36,63,.06);
  width:100% !important; max-width:100% !important; align-self:stretch !important; }
.pda-left-box:last-child { margin-bottom:0; }
/* zaoblene hrany hore aj dole - vidno, kde posuvny zoznam konci */
#${SCROLL_ID} { border:0 !important; background:transparent !important; border-radius:14px !important; }
#${LIST_ID} { background:transparent !important; }
/* Zaznam je samostatna pilulka s jemnym modrastym pozadim a malym radiusom,
   bez ramu a bez deliacich liniek - od seba ich oddeluje len medzera. Text je
   v troch riadkoch: vyrobna zakazka (tucne), zakaznicka zakazka, material. */
#${LIST_ID} .sapMLIB.pda-pill-on { min-height:0 !important; height:auto !important; padding:0 !important;
  margin:5px 3px !important; border:0 !important; border-radius:10px !important;
  background:#eef2f9 !important; overflow:hidden;
  transition:background-color .12s; width:auto !important; max-width:100% !important; }
#${LIST_ID} .sapMLIB.pda-pill-on:hover { background:#e2eaf6 !important; }
/* Vybrana zakazka: modra, ale zamerne stlmena (menej sytosti nez cista
   "linkova" modra) - v zozname ma byt zretelna, nie krikliva. */
#${LIST_ID} .sapMLIB.pda-pill-on.sapMLIBSelected { background:#2f5fa8 !important; }
#${LIST_ID} .sapMLIB.pda-pill-on.sapMLIBSelected:hover { background:#3568b6 !important; }
#${LIST_ID} .sapMLIB.pda-pill-on.sapMLIBSelected .pda-pill,
#${LIST_ID} .sapMLIB.pda-pill-on.sapMLIBSelected .t-vyr,
#${LIST_ID} .sapMLIB.pda-pill-on.sapMLIBSelected .t-zak,
#${LIST_ID} .sapMLIB.pda-pill-on.sapMLIBSelected .mat { color:#fff !important; }
#${LIST_ID} .sapMLIB.pda-pill-on.sapMLIBSelected .matn { color:#d3e3f8 !important; }
#${LIST_ID} .sapMLIB.pda-pill-on > *:not([data-pda-pill]) { display:none !important; }
.pda-pill { padding:8px 12px; font:12px/1.3 -apple-system,"Segoe UI",Roboto,sans-serif; color:#1a2233; }
/* prave bezi - zeleny prizvuk pri lavom okraji riadku */
.pda-pill.run { box-shadow: inset 4px 0 0 #2e9e4f; }
/* r1 sa zalamuje: na prvom riadku vyrobna zakazka a vpravo zeleny tag,
   zakaznicka zakazka sa cez flex-basis:100% pretlaci na vlastny riadok
   (v DOM je pritom medzi nimi - preto to poradie riesi "order") */
.pda-pill .r1 { display:flex; align-items:center; gap:8px; flex-wrap:wrap; }
.pda-pill .r2 { display:flex; align-items:baseline; gap:5px; margin-top:2px; overflow:hidden; white-space:nowrap; }
/* zo vsetkych troch textov robime obycajny text - tag ostava len z "vyrába" */
.pda-pill .t { display:inline-block; padding:0; border-radius:0; background:none; border:0;
  white-space:nowrap; letter-spacing:0; line-height:1.35; }
.pda-pill .t-vyr { order:1; font-size:13px; font-weight:700; color:#0f2547;
  flex:0 1 auto; min-width:0; overflow:hidden; text-overflow:ellipsis; }
.pda-pill .t-zak { order:3; flex:0 0 100%; font-size:12px; font-weight:400; color:#5a6b85; }
.pda-pill .t-run { order:2; margin-left:auto; flex:0 0 auto;
  background:#e7f6ec; color:#1d7a3c; border:1px solid #b6e2c5;
  padding:1px 8px; border-radius:999px; font-size:11px; font-weight:700; }
.pda-pill .mat { font-size:11.5px; font-weight:600; color:#6b7c95; white-space:nowrap; }
.pda-pill .matn { font-size:11.5px; color:#6b7c95; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; flex:1 1 auto; min-width:0; }
/* bublina s celym obsahom zakazky */
#${TIP_ID} { display:none; position:fixed; z-index:100001; pointer-events:none;
  background:#fff; border:1px solid #b9cbe8; border-radius:10px; padding:10px 12px;
  box-shadow:0 10px 30px rgba(16,36,63,.22); max-width:520px;
  font:12.5px/1.45 -apple-system,"Segoe UI",Roboto,sans-serif; color:#1a2233; }
#${TIP_ID} .tab { display:grid; grid-template-columns:auto 1fr; gap:3px 12px; }
#${TIP_ID} .r { display:contents; }
#${TIP_ID} .k { color:#6b7c95; font-size:11px; text-transform:uppercase; letter-spacing:.05em;
  white-space:nowrap; padding-top:1px; }
#${TIP_ID} .v { color:#13315c; font-weight:600; }
`;
            document.head.appendChild(st);
        }

        /*
         * Udaje polozky. Najprv z modelu appky - pozor, zoznam moze byt naviazany
         * na POMENOVANY model, vtedy getBindingContext() bez mena vrati nic,
         * preto sa prejdu vsetky kontexty polozky (c.oBindingContexts).
         * Ked sa model neda precitat (v1.13.0 sa to tak aj stalo), pilulka sa
         * poskladá z textu, ktory appka do polozky napisala - tvar tych styroch
         * riadkov je pevny, takze sa to da bezpecne rozobrat.
         */
        function itemData(li) {
            const c = resolveControl(li);
            if (c) {
                const o = zKontextu(c);
                if (o) return o;
            }
            return zTextu(li);
        }

        function pouzitelne(o) {
            return o && (o.productionOrderNo || o.salesOrderNo || o.materialNo) ? o : null;
        }

        function zKontextu(c) {
            try {
                if (typeof c.getBindingContext === 'function') {
                    const ctx = c.getBindingContext();
                    const o = ctx && typeof ctx.getObject === 'function' ? ctx.getObject() : null;
                    if (pouzitelne(o)) return o;
                }
            } catch (e) { /* ignore */ }
            try {
                const mapa = c.oBindingContexts || {};
                for (const meno in mapa) {
                    const ctx = mapa[meno];
                    const o = ctx && typeof ctx.getObject === 'function' ? ctx.getObject() : null;
                    if (pouzitelne(o)) return o;
                }
            } catch (e) { /* ignore */ }
            return null;
        }

        // texty, ktore appka do polozky napisala (aj ked su uz skryte nasou pilulkou)
        const TEXT_SEL = '.sapMText, .sapMLabel, .sapMTitle, .sapMObjectIdentifierTitle,' +
                         '.sapMObjectIdentifierText, .sapMSLITitleOnly, .sapMSLIDescription';

        function riadkyTextu(li) {
            const out = [];
            li.querySelectorAll(TEXT_SEL).forEach((el) => {
                if (el.closest('[data-pda-pill]')) return;   // nase vlastne texty nie
                if (el.querySelector(TEXT_SEL)) return;      // len listy stromu
                const t = (el.textContent || '').replace(/\s+/g, ' ').trim();
                if (t && t !== '-' && out.indexOf(t) === -1) out.push(t);
            });
            return out;
        }

        /*
         * Tvar riadkov v polozke (overene na pracovisku 4836):
         *   7006003722 - 000010                        zakaznicka zakazka - polozka
         *   001600108392 - 0700 - 000000               vyrobna zakazka - operacia - sekvencia
         *   25288306 - S-PLATE CHARG.DOOR IM250E …     material - nazov
         *   WFT11CNC-B Horizontka 110 2 UP             popis operacie
         */
        function zTextu(li) {
            const d = {};
            riadkyTextu(li).forEach((t) => {
                let m;
                if ((m = t.match(/^(\d{6,12})\s*-\s*(\d{3,6})\s*-\s*(\d{3,6})$/))) {
                    d.productionOrderNo = m[1]; d.operationNo = m[2]; d.sequenceNo = m[3];
                } else if ((m = t.match(/^(\d{6,12})\s*-\s*(\d{3,6})$/))) {
                    d.salesOrderNo = m[1]; d.salesOrderItem = m[2];
                } else if ((m = t.match(/^(\d{6,10})\s*-\s*(\S.*)$/))) {
                    d.materialNo = m[1]; d.material = m[2];
                } else if (!d.description) {
                    d.description = t;
                }
            });
            return pouzitelne(d);
        }

        function txt(v) {
            return String(v === undefined || v === null ? '' : v).trim();
        }

        function span(cls, text) {
            const el = document.createElement('span');
            el.className = cls;
            el.textContent = text;
            return el;
        }

        /*
         * V pilulke su prve TRI riadky povodnej polozky (zakaznicka zakazka,
         * vyrobna zakazka s operaciou, material s nazvom). Stvrty riadok - popis
         * operacie - sa uz nezobrazuje, je len v bubline po nabehnuti mysou.
         */
        function buildPill(d) {
            const box = document.createElement('div');
            box.setAttribute('data-pda-pill', '1');
            box.className = 'pda-pill' + (Number(d.count) > 0 ? ' run' : '');

            // bez cisla zakazky nema zmysel ukazovat samotnu polozku ("000000")
            const zak = txt(d.salesOrderNo)
                ? [txt(d.salesOrderNo), txt(d.salesOrderItem)].filter(Boolean).join(' - ')
                : '';
            const vyroba = [txt(d.productionOrderNo), txt(d.operationNo), txt(d.sequenceNo)]
                .filter(Boolean).join(' - ');

            const r1 = document.createElement('div');
            r1.className = 'r1';
            // najprv vyrobna zakazka (tmavomodra), az potom zakaznicka (svetla)
            if (vyroba) r1.appendChild(span('t t-vyr', vyroba));
            if (zak) r1.appendChild(span('t t-zak', zak));
            if (Number(d.count) > 0) r1.appendChild(span('t t-run', '● vyrába'));
            box.appendChild(r1);

            const material = txt(d.materialNo).replace(/^0+/, '');
            const nazov = txt(d.material) || txt(d.descriptionShort);

            const r2 = document.createElement('div');
            r2.className = 'r2';
            if (material) r2.appendChild(span('mat', material));
            if (nazov) r2.appendChild(span('matn', nazov));
            if (r2.childNodes.length) box.appendChild(r2);

            udaje.set(box, d);   // pre bublinu pri nabehnuti mysou
            return box;
        }

        /* ---- bublina s celym obsahom zakazky (nabehnutie mysou) ----
         * Vlastna bublina namiesto systemoveho `title`: ukaze sa hned (nie po
         * sekunde), da sa ostylovat a zmestia sa do nej aj udaje, ktore sa do
         * pilulky nevojdu - hlavne stvrty riadok (popis operacie) a stav.
         */
        const udaje = new WeakMap();
        let tip = null;

        function tipEl() {
            if (tip && tip.isConnected) return tip;
            tip = document.createElement('div');
            tip.id = TIP_ID;
            document.body.appendChild(tip);
            return tip;
        }

        function tipRiadok(tabulka, popisok, hodnota) {
            if (!hodnota) return;
            const r = document.createElement('div');
            r.className = 'r';
            const k = document.createElement('span'); k.className = 'k'; k.textContent = popisok;
            const v = document.createElement('span'); v.className = 'v'; v.textContent = hodnota;
            r.appendChild(k); r.appendChild(v);
            tabulka.appendChild(r);
        }

        function ukazTip(pill, x, y) {
            const d = udaje.get(pill);
            if (!d) return;
            const el = tipEl();
            el.textContent = '';

            const t = document.createElement('div');
            t.className = 'tab';
            tipRiadok(t, 'Zákazka', [txt(d.salesOrderNo), txt(d.salesOrderItem)].filter(Boolean).join(' - '));
            tipRiadok(t, 'Výrobná zákazka', [txt(d.productionOrderNo), txt(d.operationNo), txt(d.sequenceNo)]
                .filter(Boolean).join(' - '));
            tipRiadok(t, 'Materiál', [txt(d.materialNo).replace(/^0+/, ''), txt(d.material) || txt(d.descriptionShort)]
                .filter(Boolean).join(' - '));
            // text operacie sa v bubline zamerne neukazuje (zelanie pouzivatela)
            tipRiadok(t, 'Stav', [txt(d.status), Number(d.count) > 0 ? 'vyrába (' + d.count + ')' : '']
                .filter(Boolean).join(' · '));
            el.appendChild(t);

            el.style.display = 'block';
            posunTip(x, y);
        }

        function posunTip(x, y) {
            if (!tip) return;
            const r = tip.getBoundingClientRect();
            const l = Math.min(Math.max(8, x + 16), W.innerWidth - r.width - 8);
            const t = y + 18 + r.height > W.innerHeight ? Math.max(8, y - r.height - 12) : y + 18;
            tip.style.left = Math.round(l) + 'px';
            tip.style.top = Math.round(t) + 'px';
        }

        function skryTip() {
            if (tip) tip.style.display = 'none';
        }

        // jedno odpocuvanie na celom zozname namiesto listenerov na kazdej pilulke
        function napojTip(list) {
            if (list.dataset.pdaTipOn) return;
            list.dataset.pdaTipOn = '1';
            list.addEventListener('mouseover', (e) => {
                const pill = e.target.closest && e.target.closest('[data-pda-pill]');
                if (pill) ukazTip(pill, e.clientX, e.clientY);
            });
            list.addEventListener('mousemove', (e) => {
                if (tip && tip.style.display === 'block') posunTip(e.clientX, e.clientY);
            });
            list.addEventListener('mouseleave', skryTip);
            list.addEventListener('mouseout', (e) => {
                const kam = e.relatedTarget;
                if (!kam || !kam.closest || !kam.closest('[data-pda-pill]')) skryTip();
            });
            // pri kliku (vybera sa zakazka) bublina prekaza
            list.addEventListener('click', skryTip);
        }

        /*
         * Vyska zoznamu je PEVNA (VYSKA_ZOZNAMU). Predtym sa dopocitavala z volneho
         * miesta pod zoznamom az po spodok stlpca, takze sa menila podla toho, co
         * bolo v stlpci pod nim a ako vysoke prave bolo okno. Appka si vysku pise
         * inline, preto ju tu prepisujeme my - a len ked naozaj sedi ina hodnota,
         * aby sa DOM nesahal pri kazdom tiku.
         */
        /*
         * 3J: zoznam siaha az po spodok okna (nad patku), pod nim ostava len box
         * "Prehlad zdrojov". Produkcia mala pevnych VYSKA_ZOZNAMU px, lebo povodny
         * vypocet poskakoval - bral volne miesto "pod zoznamom", ktore sa menilo
         * s tym, co v stlpci prave bolo. Tu sa rata len z veci, ktore od vysky
         * zoznamu NEZAVISIA: horny okraj zoznamu, odsadenie pod nim vo vlastnom
         * boxe, vyska boxu Prehlad zdrojov a patka - vysledok je preto stabilny.
         * Meni sa len pri rozdiele nad 6 px, aby sa DOM nesahal pri kazdom tiku.
         * Ked sa nieco z toho neda zmerat, plati povodna pevna vyska.
         */
        const PATKA_3J = 34;          // vyska patky noveho dizajnu
        const OKRAJ_3J = 12;          // medzera nad patkou

        function fitHeight() {
            const sc = document.getElementById(SCROLL_ID);
            if (!sc || !sc.offsetParent) return;

            let ciel = VYSKA_ZOZNAMU;
            try {
                const box = document.getElementById(BOX1_ID);
                const scR = sc.getBoundingClientRect();
                if (box && scR.top > 0) {
                    const podZoznamomVBoxe = Math.max(0, box.getBoundingClientRect().bottom - scR.bottom);
                    const graf = document.getElementById(BOX2_ID);
                    const grafR = graf ? graf.getBoundingClientRect() : null;
                    const podBoxom = grafR && grafR.height > 0 ? grafR.height + 10 : 0;
                    const spodok = W.innerHeight - PATKA_3J - OKRAJ_3J;
                    ciel = Math.round(Math.max(MIN_HEIGHT, spodok - podBoxom - podZoznamomVBoxe - scR.top));
                }
            } catch (e) { ciel = VYSKA_ZOZNAMU; }

            const teraz = parseFloat(sc.style.height) || 0;
            if (Math.abs(teraz - ciel) > 6) sc.style.height = ciel + 'px';
        }

        /*
         * Lavy stlpec sa rozdeli na dva zaoblene boxy: v prvom zoznam zakaziek,
         * v druhom graf. Povodne deti stlpca sa presunu do nasich obalov - to,
         * co obsahuje zoznam (a vsetko pred nim), ide do prveho, zvysok do druheho.
         * Ked appka stlpec prekresli, obaly zmiznu a pri dalsom tiku sa spravia
         * znova.
         */
        function rozdelStlpec() {
            const left = document.getElementById(LEFT_ID);
            const sc = document.getElementById(SCROLL_ID);
            if (!left || !sc) return;

            let b1 = document.getElementById(BOX1_ID);
            let b2 = document.getElementById(BOX2_ID);
            if (!b1) { b1 = document.createElement('div'); b1.id = BOX1_ID; b1.className = 'pda-left-box'; }
            if (!b2) { b2 = document.createElement('div'); b2.id = BOX2_ID; b2.className = 'pda-left-box'; }
            if (b1.parentElement !== left) left.insertBefore(b1, left.firstChild);
            if (b2.parentElement !== left) left.appendChild(b2);

            let zaZoznamom = false;
            let presunute = false;
            Array.from(left.children).forEach((ch) => {
                if (ch === b1 || ch === b2) return;
                if (ch.contains(sc)) {
                    if (ch.parentElement !== b1) { b1.appendChild(ch); presunute = true; }
                    zaZoznamom = true;
                    return;
                }
                const ciel = zaZoznamom ? b2 : b1;
                if (ch.parentElement !== ciel) { ciel.appendChild(ch); presunute = true; }
            });

            /*
             * Po presune ma graf novy ramec - Chart.js sa prepocita az po `resize`,
             * inak by ostal neviditelny (vysoky 0 px). `resize` vsak rozhybe cely
             * UI5, takze ho posielame najviac raz za sekundu; bez tejto poistky
             * by sa pri kazdom prekresleni appky mohla rozbehnut spatna vazba
             * (presun -> resize -> prekreslenie -> presun) a appka by zamrzla.
             */
            if (presunute && Date.now() - poslednyResize > 1000) {
                poslednyResize = Date.now();
                setTimeout(() => W.dispatchEvent(new Event('resize')), 60);
            }
        }

        function apply() {
            const list = document.getElementById(LIST_ID);
            if (!list) return;
            injectStyles();
            rozdelStlpec();

            napojTip(list);

            const polozky = list.querySelectorAll('.sapMLIB');
            let hotovych = 0;

            polozky.forEach((li) => {
                const d = itemData(li);
                if (!d) return;
                hotovych++;
                const key = [d.salesOrderNo, d.salesOrderItem, d.productionOrderNo, d.operationNo,
                             d.sequenceNo, d.materialNo, d.count].join('|');
                if (li.dataset.pdaPillKey === key && li.querySelector('[data-pda-pill]')) return;

                const stary = li.querySelector('[data-pda-pill]');
                if (stary) stary.remove();
                li.appendChild(buildPill(d));
                li.dataset.pdaPillKey = key;
                li.classList.add('pda-pill-on');
            });

            // aby bolo v konzole hned vidiet, ci sa udaje polozky daju precitat
            const podpis = hotovych + '/' + polozky.length;
            if (polozky.length && podpis !== poslednyPodpis) {
                poslednyPodpis = podpis;
                console.log(LOG, 'pilulky v pracovnom zozname:', podpis);
            }

            fitHeight();
        }

        DomWatch.add(apply);
        W.addEventListener('resize', fitHeight);
        onReady(apply);
    }

    /* -------- 3.10 Popis operacie: velke tlacidlo + okno cez obrazovku -------- */

    /*
     * Popis operacie bol v detaile vysadzany drobnym pismom do uzkeho stlpca a
     * pri dlhom texte sa nedal precitat. Modul povodny blok skryje a na jeho
     * miesto da velke tlacidlo; po kliknuti sa popis ukaze cez celu obrazovku
     * vo velkom pisme. Zatvara sa klikom mimo okna, krizikom alebo Esc.
     *
     * Prvky appky (overene v prehliadaci uz v Python verzii):
     *   WorkcenterDetail--Description_SimpleForm     - cely blok "Popis:"
     *   WorkcenterDetail--Description_FormattedText  - samotny text
     */
    function modOperationDescription() {
        const FORM_ID = 'WorkcenterDetail--Description_SimpleForm';
        const TEXT_ID = 'WorkcenterDetail--Description_FormattedText';
        const BTN_ID = '__pda_opis_button__';
        const OVERLAY_ID = '__pda_opis_overlay__';
        const STYLE_ID = '__pda_opis_styles__';
        const PRAZDNY = 'k tejto operácii nie je popis';

        function injectStyles() {
            if (document.getElementById(STYLE_ID)) return;
            const st = document.createElement('style');
            st.id = STYLE_ID;
            st.textContent = `
.pda-opis-skryty { display:none !important; }
#${BTN_ID} { display:flex; align-items:center; gap:14px; width:100%; box-sizing:border-box; margin:6px 0 20px;
  padding:12px 18px; border:2px solid #b9cbe8; border-radius:14px; background:#f4f8ff; cursor:pointer;
  text-align:left; font:14px/1.4 -apple-system,"Segoe UI",Roboto,sans-serif; color:#13315c;
  box-shadow:0 2px 6px rgba(16,36,63,.10);
  transition:transform .13s ease, box-shadow .13s ease, border-color .13s, background .13s; }
/* rovnako "plasticke" ako stavove tlacidla - pri prechode mysou sa nadvihne */
#${BTN_ID}:hover { border-color:#2563eb; background:#eaf2ff; transform:translateY(-3px);
  box-shadow:0 10px 20px rgba(16,36,63,.22); }
#${BTN_ID}:active { transform:translateY(-1px); box-shadow:0 3px 8px rgba(16,36,63,.18); }
#${BTN_ID} .ikona { font-size:30px; line-height:1; flex:0 0 auto; }
#${BTN_ID} .stred { flex:1 1 auto; min-width:0; }
#${BTN_ID} .nadpis { display:block; font-size:15px; font-weight:800; letter-spacing:.06em; text-transform:uppercase; }
#${BTN_ID} .ukazka { display:block; font-size:12.5px; color:#5b6b83; margin-top:2px;
  overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
#${BTN_ID} .sipka { flex:0 0 auto; font-size:13px; font-weight:700; color:#2563eb; white-space:nowrap; }
#${BTN_ID}.prazdny { border-color:#e2e6ec; background:#fafbfc; color:#8a93a3; cursor:default; }
#${BTN_ID}.prazdny .sipka { display:none; }
#${OVERLAY_ID} { position:fixed; inset:0; background:rgba(10,20,40,.55); z-index:100000;
  display:flex; align-items:center; justify-content:center; padding:24px; box-sizing:border-box; }
#${OVERLAY_ID} .karta { background:#fff; border-radius:16px; width:min(1150px,94vw); max-height:90vh;
  display:flex; flex-direction:column; overflow:hidden; box-shadow:0 24px 70px rgba(16,36,63,.4);
  font-family:-apple-system,"Segoe UI",Roboto,sans-serif; }
#${OVERLAY_ID} .hl { display:flex; align-items:center; gap:14px; padding:14px 20px; background:#13315c; color:#fff; }
#${OVERLAY_ID} .hl .n { font-size:16px; font-weight:800; letter-spacing:.08em; text-transform:uppercase; }
#${OVERLAY_ID} .hl .z { font-size:13px; opacity:.8; }
#${OVERLAY_ID} .hl .x { margin-left:auto; background:none; border:0; color:#fff; font-size:30px; line-height:1;
  cursor:pointer; padding:0 4px; }
#${OVERLAY_ID} .telo { padding:24px 30px 30px; overflow:auto; font-size:21px; line-height:1.62; color:#17202e;
  white-space:pre-wrap; word-break:break-word; }
`;
            document.head.appendChild(st);
        }

        function textEl() {
            return document.getElementById(TEXT_ID) || document.querySelector('[id$="Description_FormattedText"]');
        }

        function formEl() {
            const f = document.getElementById(FORM_ID);
            if (f) return f;
            const t = textEl();
            return t ? (t.closest('.sapUiForm') || t.parentElement) : null;
        }

        function popisText() {
            const t = textEl();
            if (!t) return '';
            return String(t.innerText || t.textContent || '').replace(/ /g, ' ').trim();
        }

        // cislo zakazky do hlavicky okna (kozmetika - ked sa neda precitat, nic sa nedeje)
        function zakazkaPopis() {
            try {
                const main = W.sap.ui.getCore().byId('Main');
                const op = main && main.getController().getGlobals().getVar('oSelectedWorkcenterOperation');
                if (!op || !op.productionOrderNo) return '';
                return [op.productionOrderNo, op.operationNo].filter(Boolean).join(' / ');
            } catch (e) { return ''; }
        }

        function openOverlay() {
            const text = popisText();
            if (!text) return;
            const stary = document.getElementById(OVERLAY_ID);
            if (stary) stary.remove();

            const overlay = document.createElement('div');
            overlay.id = OVERLAY_ID;

            const karta = document.createElement('div');
            karta.className = 'karta';

            const hl = document.createElement('div');
            hl.className = 'hl';
            const n = document.createElement('span'); n.className = 'n'; n.textContent = 'Popis operácie';
            const z = document.createElement('span'); z.className = 'z'; z.textContent = zakazkaPopis();
            const x = document.createElement('button'); x.className = 'x'; x.type = 'button';
            x.textContent = '×'; x.setAttribute('aria-label', 'Zavrieť');
            hl.appendChild(n); hl.appendChild(z); hl.appendChild(x);

            const telo = document.createElement('div');
            telo.className = 'telo';
            telo.textContent = text;

            karta.appendChild(hl); karta.appendChild(telo);
            overlay.appendChild(karta);

            const zavri = () => { overlay.remove(); document.removeEventListener('keydown', naEsc); };
            const naEsc = (e) => { if (e.key === 'Escape') zavri(); };

            x.addEventListener('click', zavri);
            overlay.addEventListener('click', (e) => { if (e.target === overlay) zavri(); });
            document.addEventListener('keydown', naEsc);

            document.body.appendChild(overlay);
        }

        function apply() {
            const form = formEl();
            if (!form || !form.parentElement) return;
            injectStyles();

            let btn = document.getElementById(BTN_ID);
            if (!btn) {
                btn = document.createElement('button');
                btn.id = BTN_ID;
                btn.type = 'button';
                const ikona = document.createElement('span'); ikona.className = 'ikona'; ikona.textContent = '📄';
                const stred = document.createElement('span'); stred.className = 'stred';
                const nadpis = document.createElement('span'); nadpis.className = 'nadpis';
                nadpis.textContent = 'Popis operácie';
                const ukazka = document.createElement('span'); ukazka.className = 'ukazka';
                stred.appendChild(nadpis); stred.appendChild(ukazka);
                const sipka = document.createElement('span'); sipka.className = 'sipka'; sipka.textContent = 'otvoriť ▸';
                btn.appendChild(ikona); btn.appendChild(stred); btn.appendChild(sipka);
                btn.addEventListener('click', openOverlay);
            }
            if (btn.parentElement !== form.parentElement || btn.nextElementSibling !== form) {
                form.parentElement.insertBefore(btn, form);
            }
            if (!form.classList.contains('pda-opis-skryty')) form.classList.add('pda-opis-skryty');

            const text = popisText();
            const ukazka = btn.querySelector('.ukazka');
            const skratene = text.length > 120 ? text.slice(0, 120) + '…' : text;
            if (ukazka && ukazka.textContent !== (skratene || PRAZDNY)) {
                ukazka.textContent = skratene || PRAZDNY;
            }
            btn.classList.toggle('prazdny', !text);
            btn.disabled = !text;
        }

        DomWatch.add(apply);
        onReady(apply);
    }

    /* ------------- 3.11 Krajsi graf vytazenia (Resource over time) ------------- */

    /*
     * Graf pod pracovnym zoznamom kresli appka cez Chart.js 4.4.2 do platna
     * `ResourceDetails`, detail (lupa) do `DialogChart` v dialogu
     * `Popups--Dialog_Chart` - zistene zo zdrojakov appky (PDA_Chart.displayTimelineChart).
     *
     * Modul graf NEKRESLI odznova, iba prestavi hotovu instanciu, takze udaje
     * ostavaju presne tie, ktore poslala appka.
     *
     * ⚠️ Dolezite (zistene 2026-09-14 pri v1.16.0): v Chart.js v4 NESTACI menit
     * `chart.options` - pri prekresleni sa nastavenia beru z `chart.config.options`
     * (a hotove osi maju este vlastnu kopiu v `chart.scales.x.options`). Preto sa
     * ta ista uprava nanasa na VSETKY tieto miesta; inak sa zmeni len hrubka pasov
     * (tie su v `chart.data.datasets`) a popisky casu ostanu dlhe.
     */
    function modChartStyle() {
        const MALY = 'ResourceDetails';
        const VELKY = 'DialogChart';
        const DIALOG_ID = 'Popups--Dialog_Chart';
        const STYLE_ID = '__pda_chart_styles__';

        // maly graf pod zoznamom: co najnizsi, nech ostane miesto na zakazky
        const M_RIADOK = 30, M_OKRAJE = 30, M_MIN = 84, M_MAX = 220;
        // detail: siroky, ale nie na celu vysku obrazovky
        const V_RIADOK = 80, V_OKRAJE = 165, V_MIN = 260;   // vacsie okraje = miesto na popisky casu

        let chybaKnizniceLogged = false;
        let detailOtvoreny = false;
        let zatvaranieNapojene = false;

        function injectStyles() {
            if (document.getElementById(STYLE_ID)) return;
            const st = document.createElement('style');
            st.id = STYLE_ID;
            st.textContent = `
#WorkcenterDetail--ChartFlexBox { padding-top:0 !important; margin-top:0 !important; }
canvas#${MALY} { max-width:100% !important; }
/* nadpis "Resource over time" zabral dva riadky - staci jemny jednoriadkovy popisok */
.pda-chart-nadpis { font-size:11px !important; font-weight:700 !important; letter-spacing:.1em !important;
  text-transform:uppercase !important; color:#8e9bb0 !important; white-space:nowrap !important;
  line-height:1.2 !important; margin:0 !important; padding:0 !important; }
.pda-chart-nadpis .sapMTitleInner, .pda-chart-nadpis bdi, .pda-chart-nadpis span { font-size:11px !important;
  font-weight:700 !important; color:#8e9bb0 !important; white-space:nowrap !important; }
/* detail grafu (lupa): siroky na celu obrazovku, vysoky len tolko, kolko treba */
#${DIALOG_ID} { width:96vw !important; max-width:96vw !important; left:2vw !important;
  height:auto !important; max-height:82vh !important; top:9vh !important; transform:none !important;
  border-radius:16px !important; box-shadow:0 24px 70px rgba(16,36,63,.4) !important; }
#${DIALOG_ID} .sapMDialogSection, #${DIALOG_ID} .sapMDialogScrollCont { height:auto !important;
  max-height:none !important; padding:6px 12px 10px !important; }
/* poistka: kym kod graf nedoladi, ma okno aspon nejaku vysku - inak by sa
   na okamih ukazalo prazdne (vysku platna nastavuje az JS) */
#${DIALOG_ID} .sapMDialogScrollCont { min-height:300px !important; }
#${DIALOG_ID} canvas#${VELKY} { max-width:100% !important; min-height:260px !important; }
`;
            document.head.appendChild(st);
        }

        function kniznica() {
            const C = W.Chart;
            if (C && (typeof C.getChart === 'function' || C.instances)) return C;
            if (!chybaKnizniceLogged) {
                chybaKnizniceLogged = true;
                console.log(LOG, 'Chart.js sa nenašiel, graf len zmenším cez CSS');
            }
            return null;
        }

        function instancia(C, canvas) {
            try {
                if (typeof C.getChart === 'function') {
                    const ch = C.getChart(canvas);
                    if (ch) return ch;
                }
            } catch (e) { /* ignore */ }
            try {
                const zoznam = C.instances || {};
                for (const k in zoznam) {
                    if (zoznam[k] && zoznam[k].canvas === canvas) return zoznam[k];
                }
            } catch (e) { /* ignore */ }
            return null;
        }

        function dve(n) {
            return (n < 10 ? '0' : '') + n;
        }

        /*
         * Z hocijakeho tvaru casu spravi "07:00". Popisok moze prist ako cislo
         * (timestamp), Date alebo text "2026-09-14 07:00:00" - preto tri cesty.
         */
        function hhmm(v) {
            if (v === null || v === undefined) return v;
            if (v instanceof Date) return dve(v.getHours()) + ':' + dve(v.getMinutes());
            if (typeof v === 'number') {
                const d = new Date(v);
                return isNaN(d.getTime()) ? v : dve(d.getHours()) + ':' + dve(d.getMinutes());
            }
            const s = String(v);
            const m = s.match(/(\d{1,2}):(\d{2})/);
            return m ? dve(Number(m[1])) + ':' + m[2] : s;
        }

        /*
         * Popisky casu su u tohto grafu v `chart.data.labels`. Ukazalo sa (v1.16.1),
         * ze zmeny v `options` sa pri prekresleni nie vzdy prejavia - appka si graf
         * sama aktualizuje - kym zmeny v `data` drzia spolahlivo. Preto sa sekundy
         * odrezavaju priamo v popiskoch.
         * Prisny vzor: prepise sa LEN retazec, ktory je cely cas ("06:10:00",
         * "2026-09-14 06:10:00"). Mena pracovnikov ("4829: Maros Minarik") ostanu.
         */
        const CAS_CELY = /^\s*(?:\d{4}-\d{2}-\d{2}[ T])?(\d{1,2}):(\d{2})(?::\d{2})?\s*$/;

        function skratPopisky(chart) {
            const l = chart.data && chart.data.labels;
            if (!Array.isArray(l)) return false;
            let zmena = false;
            for (let i = 0; i < l.length; i++) {
                if (typeof l[i] !== 'string') continue;
                const m = l[i].match(CAS_CELY);
                if (!m) continue;
                const novy = dve(Number(m[1])) + ':' + m[2];
                if (novy !== l[i]) { l[i] = novy; zmena = true; }
            }
            return zmena;
        }

        function riadkov(chart) {
            try {
                const l = chart.data && chart.data.labels;
                if (l && l.length) return l.length;
            } catch (e) { /* ignore */ }
            return 1;
        }

        function vyskaMaleho(chart) {
            return Math.min(M_MAX, Math.max(M_MIN, riadkov(chart) * M_RIADOK + M_OKRAJE));
        }

        function vyskaVelkeho(chart) {
            const strop = Math.round(W.innerHeight * 0.74);
            return Math.min(strop, Math.max(V_MIN, riadkov(chart) * V_RIADOK + V_OKRAJE));
        }

        // vsetky miesta, odkial Chart.js berie nastavenia (viac o tom v komentari hore)
        function vsetkyOptions(chart) {
            const out = [];
            const pridaj = (o) => { if (o && out.indexOf(o) === -1) out.push(o); };
            try { pridaj(chart.options); } catch (e) { /* ignore */ }
            try { pridaj(chart.config && chart.config.options); } catch (e) { /* ignore */ }
            try { pridaj(chart.config && chart.config._config && chart.config._config.options); } catch (e) { /* ignore */ }
            return out;
        }

        function osX(velky) {
            return {
                maxRotation: 0,
                minRotation: 0,
                autoSkip: true,
                maxTicksLimit: velky ? 20 : 9,
                color: velky ? '#5b6b83' : '#8e9bb0',
                font: { size: velky ? 13 : 11 },
                callback(value) {
                    let raw = value;
                    try {
                        if (typeof this.getLabelForValue === 'function') raw = this.getLabelForValue(value);
                    } catch (e) { /* ignore */ }
                    return hhmm(raw);
                },
            };
        }

        function nastav(o, velky) {
            o.maintainAspectRatio = false;
            o.responsive = true;
            o.animation = false;
            o.layout = Object.assign({}, o.layout, { padding: { top: 2, right: 8, bottom: 0, left: 2 } });

            const sc = o.scales = o.scales || {};
            const x = sc.x = sc.x || {};
            x.ticks = Object.assign({}, x.ticks, osX(velky));
            x.grid = Object.assign({}, x.grid, {
                color: 'rgba(120,140,170,.14)', drawBorder: false, tickLength: 4,
            });
            x.border = Object.assign({}, x.border, { display: false });
            if (x.time) {
                x.time.displayFormats = Object.assign({}, x.time.displayFormats, {
                    millisecond: 'HH:mm', second: 'HH:mm', minute: 'HH:mm', hour: 'HH:mm', day: 'HH:mm',
                });
                x.time.tooltipFormat = 'HH:mm';
            }

            const y = sc.y = sc.y || {};
            y.ticks = Object.assign({}, y.ticks, {
                color: '#5b6b83', font: { size: velky ? 14 : 11.5, weight: '600' },
            });
            y.grid = Object.assign({}, y.grid, { display: false, drawBorder: false });
            y.border = Object.assign({}, y.border, { display: false });

            const p = o.plugins = o.plugins || {};
            // nadpis "Workcenter reports from last 24h: ..." je zbytocny
            p.title = Object.assign({}, p.title, { display: false });
            p.subtitle = Object.assign({}, p.subtitle, { display: false });
            p.tooltip = Object.assign({}, p.tooltip, {
                backgroundColor: 'rgba(19,49,92,.94)',
                titleFont: { size: velky ? 14 : 12.5 },
                bodyFont: { size: velky ? 14 : 12.5 },
                padding: 10,
                cornerRadius: 8,
            });
            if (p.legend) {
                p.legend.labels = Object.assign({}, p.legend.labels, {
                    boxWidth: 12, boxHeight: 12, usePointStyle: true,
                    color: '#5b6b83', font: { size: velky ? 13 : 11.5 },
                });
            }
        }

        function upravChart(chart, canvas) {
            if (!chart) return false;
            const velky = canvas.id === VELKY;

            // pasy: zaoblene, v detaile hrubsie
            (chart.data && chart.data.datasets ? chart.data.datasets : []).forEach((ds) => {
                ds.borderRadius = velky ? 9 : 6;
                ds.borderSkipped = false;
                ds.barThickness = velky ? 34 : 22;
                ds.maxBarThickness = velky ? 44 : 26;
                ds.borderWidth = 0;
            });

            skratPopisky(chart);
            vsetkyOptions(chart).forEach((o) => { try { nastav(o, velky); } catch (e) { /* ignore */ } });

            // hotove osi maju vlastnu kopiu nastaveni - bez toho ostanu dlhe popisky
            try {
                if (chart.scales && chart.scales.x && chart.scales.x.options) {
                    chart.scales.x.options.ticks = Object.assign({}, chart.scales.x.options.ticks, osX(velky));
                }
            } catch (e) { /* ignore */ }

            const wrap = canvas.parentElement;
            if (wrap) {
                wrap.style.position = wrap.style.position || 'relative';
                wrap.style.height = (velky ? vyskaVelkeho(chart) : vyskaMaleho(chart)) + 'px';
            }

            try { chart.resize(); } catch (e) { /* ignore */ }
            try { chart.update(); } catch (e) { /* ignore */ }

            if (!chart.__pdaUpraveny) {
                chart.__pdaUpraveny = true;
                let typOsi = '?';
                try { typOsi = chart.scales && chart.scales.x ? chart.scales.x.type : 'nie je'; } catch (e) { /* ignore */ }
                const ukazka = Array.isArray(chart.data && chart.data.labels)
                    ? chart.data.labels.slice(0, 3) : '(bez labels)';
                console.log(LOG, 'graf upravený:', canvas.id, '· riadkov:', riadkov(chart),
                            '· os X:', typOsi, '· popisky:', ukazka);
            }
            return true;
        }

        // dvojriadkovy nadpis nad malym grafom stlacime na jeden jemny riadok
        function zmensiNadpis() {
            const left = document.getElementById('WorkcenterDetail--LeftColumn_FlexBox');
            if (!left) return;
            left.querySelectorAll('.sapMTitle, .sapMLabel, .sapMText').forEach((el) => {
                if (el.classList.contains('pda-chart-nadpis')) return;
                const t = (el.textContent || '').trim().toLowerCase();
                if (t === 'resource over time' || t === 'zdroj v čase' || t === 'ressource über zeit') {
                    el.classList.add('pda-chart-nadpis');
                }
            });
        }

        function detailJeOtvoreny() {
            const dlg = document.getElementById(DIALOG_ID);
            return !!(dlg && dlg.getBoundingClientRect().width > 0);
        }

        // klik vedla okna detail zavrie (appka na to vlastne tlacidlo nema)
        function napojZatvaranie() {
            if (zatvaranieNapojene) return;
            zatvaranieNapojene = true;
            document.addEventListener('mousedown', (e) => {
                const dlg = document.getElementById(DIALOG_ID);
                if (!dlg || dlg.getBoundingClientRect().width === 0) return;
                if (dlg.contains(e.target)) return;
                try {
                    const ctrl = getControl(DIALOG_ID);
                    if (ctrl && typeof ctrl.close === 'function') ctrl.close();
                } catch (err) { /* ignore */ }
            }, true);
            document.addEventListener('keydown', (e) => {
                if (e.key !== 'Escape' || !detailJeOtvoreny()) return;
                try {
                    const ctrl = getControl(DIALOG_ID);
                    if (ctrl && typeof ctrl.close === 'function') ctrl.close();
                } catch (err) { /* ignore */ }
            });
        }

        /*
         * Pri otvoreni detailu treba grafu poslat `resize` - do novej velkosti
         * dialogu sa sam neprekresli (overene uz v Python verzii).
         */
        function dopasujDetail(C) {
            const otvoreny = detailJeOtvoreny();
            if (otvoreny === detailOtvoreny) return;
            detailOtvoreny = otvoreny;
            if (!otvoreny) return;

            /*
             * Graf v okne appka vytvara az po otvoreni dialogu. Predtym sa tu
             * cakalo pevnych 150 ms a ked este nebol hotovy, doladil ho az
             * pomaly tik (1,2 s) - okno preto dlho vyzeralo prazdne. Teraz sa
             * skusa kazdych 60 ms, takze sa chyti hned, ako vznikne.
             * Globalny `resize` (prekresli vsetky grafy na stranke) ostava uz
             * len ako zaloha, ked sa instancia vobec nenajde.
             */
            let pokus = 0;
            const skus = () => {
                if (!detailJeOtvoreny()) return;
                try {
                    const canvas = document.getElementById(VELKY);
                    const ch = C && canvas ? instancia(C, canvas) : null;
                    if (ch) { upravChart(ch, canvas); return; }
                } catch (e) { /* ignore */ }
                if (++pokus < 40) setTimeout(skus, 60);
                else { try { W.dispatchEvent(new Event('resize')); } catch (e) { /* ignore */ } }
            };
            setTimeout(skus, 30);
        }

        function apply() {
            injectStyles();
            napojZatvaranie();
            zmensiNadpis();

            const C = kniznica();
            dopasujDetail(C);

            [MALY, VELKY].forEach((id) => {
                const canvas = document.getElementById(id);
                if (!canvas) return;
                if (!C) {
                    // zaloha bez kniznice: aspon nizsie platno, Chart.js sa prekresli sam
                    const wrap = canvas.parentElement;
                    if (wrap && id === MALY && wrap.style.height !== M_MIN + 'px') {
                        wrap.style.height = M_MIN + 'px';
                    }
                    return;
                }
                const ch = instancia(C, canvas);
                if (!ch) return;
                if (!ch.__pdaUpraveny) {
                    upravChart(ch, canvas);
                } else if (skratPopisky(ch)) {
                    // appka graf prekreslila s dlhymi popiskami - znovu ich skratime
                    try { ch.update('none'); } catch (e) { /* ignore */ }
                }
            });
        }

        DomWatch.add(apply);
        // graf vznika az po odpovedi zo servera, preto este pomaly tik
        setInterval(apply, 1200);
        onReady(apply);
    }

    /* ------------- 3.12 Krajsia tabulka stavov (mriezka pri grafe) ------------- */

    /*
     * Tlacidlo s mriezkou pri grafe (WorkcenterDetail--Status_Overview_Button)
     * otvara dialog `Popups--TableSelectDialog_Overview` - holu SAP tabulku
     * so zaznamami stavov. Modul ju nechava tak, ako je (data aj klikanie),
     * len ju prekresli:
     *   - dialog siroky, zaobleny, tmavomodra hlavicka
     *   - datum a cas namiesto "2026-09-14 07:00:00" ako "14.09." + "07:00"
     *   - stav ako farebna pilulka podla tych istych pravidiel, ako maju
     *     stavove tlacidla (nastavenia -> Tlacidla - farby)
     *   - riadky vyssie, striedavo podfarbene, bez tvrdych ciar
     *
     * Povodny text bunky sa pamata v `data-pda-orig`, takze po prekresleni
     * tabulky sa prevod spravi znova a nic sa nestrati.
     */
    function modStatusTable() {
        const DIALOG_SEL = '[id^="Popups--TableSelectDialog_Overview"]';
        const STYLE_ID = '__pda_stable_styles__';

        function injectStyles() {
            if (document.getElementById(STYLE_ID)) return;
            const st = document.createElement('style');
            st.id = STYLE_ID;
            st.textContent = `
${DIALOG_SEL}.sapMDialog { border-radius:16px !important; overflow:hidden !important;
  width:min(1450px,96vw) !important; max-width:96vw !important; max-height:86vh !important;
  box-shadow:0 24px 70px rgba(16,36,63,.38) !important; }
${DIALOG_SEL} .sapMDialogTitle, ${DIALOG_SEL} .sapMIBar.sapMHeader-CTX, ${DIALOG_SEL} .sapMDialogTitleGroup {
  background:#13315c !important; color:#fff !important; }
${DIALOG_SEL} .sapMDialogTitle .sapMTitle, ${DIALOG_SEL} .sapMIBar.sapMHeader-CTX .sapMTitle {
  color:#fff !important; font-size:15px !important; font-weight:800 !important; letter-spacing:.06em; }
${DIALOG_SEL} .sapMListTblHeader { background:#f3f6fb !important; }
${DIALOG_SEL} .sapMListTblHeaderCell, ${DIALOG_SEL} .sapMListTblHeaderCell .sapMLabel {
  font-size:11px !important; font-weight:800 !important; letter-spacing:.08em !important;
  text-transform:uppercase !important; color:#6b7c95 !important; }
${DIALOG_SEL} .sapMListTbl, ${DIALOG_SEL} .sapMList { background:#fff !important; }
${DIALOG_SEL} .sapMListTblRow { border-bottom:1px solid #eef2f7 !important; }
${DIALOG_SEL} .sapMListTblRow:nth-child(even) { background:#f7f9fd !important; }
${DIALOG_SEL} .sapMListTblRow:hover { background:#eaf1fc !important; }
${DIALOG_SEL} .sapMListTblCell { padding-top:9px !important; padding-bottom:9px !important;
  font-size:13px !important; color:#1a2233 !important; vertical-align:middle !important; }
${DIALOG_SEL} .sapMListTblRow.sapMLIBSelected { background:#13315c !important; }
${DIALOG_SEL} .sapMListTblRow.sapMLIBSelected .sapMListTblCell,
${DIALOG_SEL} .sapMListTblRow.sapMLIBSelected .sapMText { color:#fff !important; }
${DIALOG_SEL} .sapMSF, ${DIALOG_SEL} .sapMSFB { border-radius:10px !important; }
/* datum a cas */
.pda-cas { display:inline-flex; align-items:baseline; gap:6px; white-space:nowrap; }
.pda-cas .d { font-size:11px; color:#8e9bb0; }
.pda-cas .c { font-size:13.5px; font-weight:700; color:#13315c; }
.sapMLIBSelected .pda-cas .d { color:#b9cbe8; }
.sapMLIBSelected .pda-cas .c { color:#fff; }
/* stav ako pilulka */
.pda-stav { display:inline-block; padding:3px 11px; border-radius:999px; font-size:12px; font-weight:700;
  line-height:1.4; white-space:nowrap; background:#eef2f7; color:#41506a; }
/* ziadne lamanie cisel a nazvov cez pol stlpca */
${DIALOG_SEL} .sapMListTblCell, ${DIALOG_SEL} .sapMListTblHeaderCell { white-space:nowrap !important; }
${DIALOG_SEL} .sapMListTblCell .sapMText, ${DIALOG_SEL} .sapMListTblCell .sapMLabel {
  white-space:nowrap !important; word-break:normal !important; }
/* stlpec s materialom je najsirsi - nazov sa ma zmestit na jeden riadok */
${DIALOG_SEL} .pda-col-material { min-width:330px !important; }
.pda-sap { display:inline-block; margin-right:8px; padding:1px 8px; border-radius:999px;
  background:#13315c; color:#fff; font-size:11.5px; font-weight:700; }
.sapMLIBSelected .pda-sap { background:#fff; color:#13315c; }
`;
            document.head.appendChild(st);
        }

        // rovnake pravidlo ako pri farebnych tlacidlach: vyhrava najdlhsi sediaci text
        function pravidloPreStav(text) {
            const t = String(text || '').toLowerCase();
            let best = null;
            for (const r of buttonRules()) {
                const rt = String(r.text || '').toLowerCase();
                if (rt && t.indexOf(rt) !== -1 && (!best || rt.length > String(best.text || '').length)) best = r;
            }
            return best;
        }

        // "2026-09-14 07:00:00" / "2026-09-14T07:00" -> { den:"14.09.", cas:"07:00" }
        function rozlozCas(s) {
            const m = String(s || '').trim()
                .match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})(?::\d{2})?$/);
            if (!m) return null;
            return { den: m[3] + '.' + m[2] + '.', cas: m[4] + ':' + m[5] };
        }

        function jeStav(text) {
            const t = String(text || '').trim();
            return t.length > 2 && t.length < 60 && !/^\d/.test(t) && !!pravidloPreStav(t);
        }

        function prekresliBunku(el) {
            const raw = el.dataset.pdaOrig !== undefined ? el.dataset.pdaOrig : (el.textContent || '').trim();
            if (!raw) return;
            // uz prevedene a od vtedy sa nic nezmenilo
            if (el.dataset.pdaOrig !== undefined && el.dataset.pdaDone === '1' &&
                el.querySelector('.pda-cas, .pda-stav')) return;

            const cas = rozlozCas(raw);
            if (cas) {
                el.dataset.pdaOrig = raw;
                el.dataset.pdaDone = '1';
                el.textContent = '';
                const w = document.createElement('span');
                w.className = 'pda-cas';
                const d = document.createElement('span'); d.className = 'd'; d.textContent = cas.den;
                const c = document.createElement('span'); c.className = 'c'; c.textContent = cas.cas;
                w.appendChild(d); w.appendChild(c);
                el.appendChild(w);
                return;
            }

            if (jeStav(raw)) {
                const r = pravidloPreStav(raw);
                el.dataset.pdaOrig = raw;
                el.dataset.pdaDone = '1';
                el.textContent = '';
                const p = document.createElement('span');
                p.className = 'pda-stav';
                p.textContent = raw;
                if (r) { p.style.background = r.bg; p.style.color = r.fg; }
                el.appendChild(p);
            }
        }

        /*
         * Stlpec "Materiál" je v tabulke najsirsi obsah (nazov vyrobku), pri
         * povodnej sirke sa lamal na styri riadky. Index stlpca sa najde podla
         * textu v hlavicke, takze to funguje aj keby appka stlpce prehodila.
         */
        function oznacStlpecMaterialu(dialog) {
            const hlavicky = dialog.querySelectorAll('.sapMListTblHeaderCell');
            if (!hlavicky.length) return;
            let idx = -1;
            hlavicky.forEach((h, i) => {
                if (idx < 0 && /materi/i.test(h.textContent || '')) idx = i;
            });
            if (idx < 0) return;
            if (!hlavicky[idx].classList.contains('pda-col-material')) {
                hlavicky[idx].classList.add('pda-col-material');
            }
            dialog.querySelectorAll('.sapMListTblRow').forEach((r) => {
                const bunky = r.querySelectorAll('.sapMListTblCell');
                if (bunky.length !== hlavicky.length || !bunky[idx]) return;   // stlpce nesedia, radsej nic
                if (!bunky[idx].classList.contains('pda-col-material')) {
                    bunky[idx].classList.add('pda-col-material');
                }
                doplnSapCislo(r, bunky[idx]);
            });
        }

        /*
         * V stlpci "Materiál" je nazov vyrobku, nie SAP cislo. Ak ho riadok
         * v datach ma (pole s "material" v nazve a hodnotou ako 25286609),
         * predradi sa pred nazov ako tmava pilulka. Ked ho data nemaju,
         * nestane sa nic.
         */
        function doplnSapCislo(row, bunka) {
            if (bunka.querySelector('.pda-sap')) return;
            let cislo = '';
            try {
                const c = resolveControl(row);
                const ctx = c && typeof c.getBindingContext === 'function' ? c.getBindingContext() : null;
                let o = ctx && typeof ctx.getObject === 'function' ? ctx.getObject() : null;
                if (!o && c && c.oBindingContexts) {
                    for (const meno in c.oBindingContexts) {
                        const x = c.oBindingContexts[meno];
                        const v = x && typeof x.getObject === 'function' ? x.getObject() : null;
                        if (v) { o = v; break; }
                    }
                }
                if (o) {
                    for (const k in o) {
                        if (!/material/i.test(k)) continue;
                        const v = String(o[k] === undefined || o[k] === null ? '' : o[k]).trim().replace(/^0+/, '');
                        if (/^\d{6,10}$/.test(v)) { cislo = v; break; }
                    }
                }
            } catch (e) { /* ignore */ }
            if (!cislo) return;

            const ciel = bunka.querySelector('.sapMText, .sapMLabel') || bunka;
            const p = document.createElement('span');
            p.className = 'pda-sap';
            p.textContent = cislo;
            ciel.insertBefore(p, ciel.firstChild);
        }

        function apply() {
            const dialog = document.querySelector(DIALOG_SEL + '.sapMDialog');
            if (!dialog) return;
            injectStyles();
            dialog.querySelectorAll('.sapMListTblCell .sapMText, .sapMListTblCell .sapMLabel')
                .forEach((el) => {
                    if (el.querySelector('.sapMText, .sapMLabel')) return;   // len listy stromu
                    try { prekresliBunku(el); } catch (e) { /* kozmetika, nikdy nesmie zhodit dialog */ }
                });
            try { oznacStlpecMaterialu(dialog); } catch (e) { /* ignore */ }
        }

        DomWatch.add(apply);
        onReady(apply);
    }

    /* ---------- 3.13 Hlavicka detailu: vsetko do jedneho riadku ---------- */

    /*
     * Hlavicka nad stavovymi tlacidlami zabrala skoro tretinu obrazovky:
     * tri riadky formulara (Zakaznicka zakazka / Material / Production Order)
     * boli od seba na 50 px, tlacidlo "Operation Complete" bolo vysoke na dva
     * riadky a okienko VYKRES malo vlastny riadok nad tym vsetkym.
     *
     * Modul nic nepresklada v appke - len:
     *   1. nase okienko VYKRES presunie do toho isteho riadku, kde uz je
     *      prepinac Machine a tlacidlo Operation Complete (poradie: prepinac,
     *      vykres, tlacidlo)
     *   2. formular stlaci do kompaktneho zaobleneho boxu (riadky tesne pod sebou)
     *   3. tlacidlo Operation Complete da na jeden riadok, nizsie a sirsie
     *
     * Prvky appky: WorkcenterDetail--Order_FlexBox (VBox) -> OrderHeader_FlexBox
     * (HBox: formular, Machine_Switch, Confirm_Button).
     */
    function modDetailHeader() {
        const HEADER_ID = 'WorkcenterDetail--OrderHeader_FlexBox';
        const CONFIRM_ID = 'WorkcenterDetail--Confirm_Button';
        const SWITCH_ID = 'WorkcenterDetail--Machine_Switch';
        const DRAWING_ID = '__pda_order_drawing_wrapper__';
        const STYLE_ID = '__pda_detail_header_styles__';
        const COL_ID = '__pda_detail_rightcol__';
        const MACHINE_ID = '__pda_detail_machine__';
        const STATUS_ID = 'WorkcenterDetail--Order_Status_Flexbox';
        const ORDER_ID = 'WorkcenterDetail--Order_FlexBox';

        function injectStyles() {
            if (document.getElementById(STYLE_ID)) return;
            const st = document.createElement('style');
            st.id = STYLE_ID;
            st.textContent = `
/* box so zakazkou: SAP mu kresli vlastny ram zlava - nepotrebujeme ho.
   Vnutorne odsadenie drzi obsah dalej od okrajov karty. */
#${ORDER_ID} { border-left:0 !important; padding:0 14px !important; }

/* cely riadok hlavicky: prepinac, vykres a tlacidlo vedla seba, zvisle na stred */
#${HEADER_ID} { align-items:center !important; gap:12px !important; padding:0 !important;
  margin:0 !important; height:auto !important; min-height:0 !important; }
/*
 * Obal hlavicky drzal pod kartou "Dokumentácia" zhruba dvojnasobnu medzeru
 * nez je medzi ostatnymi kartami. Su v nom totiz este tri zvysky po povodnom
 * rozlozeni - prazdny FlexBox2, pruzna medzera ToolbarSpacer1 a uz prazdny
 * Machine_FlexBox (prepinac stroja sme presunuli do #__pda_detail_machine__).
 * Nic nezobrazuju, ale roztahuju vysku obalu, a ten rozdiel vyzeral ako
 * vacsia medzera. Skryvame ich, takze obal uz presne obopina kartu.
 */
#WorkcenterDetail--FlexBox2,
#${HEADER_ID} > #WorkcenterDetail--ToolbarSpacer1,
#WorkcenterDetail--Machine_FlexBox:empty { display:none !important; }

/* formular (zakazka / material / production order) ako kompaktny box */
#${HEADER_ID} .sapUiForm { background:#f7f9fd !important; border:1px solid #e3e9f1 !important;
  border-radius:12px !important; padding:7px 12px !important; margin:0 !important; flex:1 1 auto; min-width:0; }
#${HEADER_ID} .sapUiForm .sapUiFormElement,
#${HEADER_ID} .sapUiForm .sapUiRespGridRow { margin:0 !important; padding:0 !important; }
#${HEADER_ID} .sapUiForm [class*="sapUiRespGridSpan"],
#${HEADER_ID} .sapUiForm [class*="sapUiRespGridHSpace"] { padding-top:1px !important; padding-bottom:1px !important;
  margin-top:0 !important; margin-bottom:0 !important; }
/* popisok aj hodnota musia sediet na tom istom riadku - rovnaka vyska riadku
   v px (nie nasobok), inak ich rozne velke pisma posunu voci sebe */
#${HEADER_ID} .sapUiForm .sapMLabel { font-size:11px !important; line-height:22px !important;
  color:#6b7c95 !important; text-transform:uppercase; letter-spacing:.04em; font-weight:700 !important;
  display:inline-block !important; vertical-align:middle !important; }
#${HEADER_ID} .sapUiForm .sapMText, #${HEADER_ID} .sapUiForm .sapMTextMaxLine {
  font-size:13px !important; line-height:22px !important; color:#13315c !important; font-weight:600 !important;
  display:inline-block !important; vertical-align:middle !important; }
#${HEADER_ID} .sapUiForm .sapMLabel .sapMLabelTextWrapper,
#${HEADER_ID} .sapUiForm .sapMText .sapMTextMaxLine { line-height:22px !important; }
#${HEADER_ID} .sapUiForm .sapUiFormTitle, #${HEADER_ID} .sapUiForm .sapUiFormTitleH5 { display:none !important; }

/* prepinac Machine */
#${SWITCH_ID} { flex:0 0 auto; margin:0 !important; }

/* okienko VYKRES presunute do riadku */
#${DRAWING_ID}.pda-v-riadku { width:auto !important; margin:0 !important; flex:0 0 auto; }
#${DRAWING_ID}.pda-v-riadku > button { margin-right:0 !important; }

/* pravy stlpec: zhora VYKRES, pod nim Operation Complete, pod tym prepinac */
#${COL_ID} { display:flex; flex-direction:column; align-items:stretch; gap:8px;
  flex:0 0 auto; margin-left:auto; padding-left:12px; }
#${COL_ID} .pda-machine { display:flex; align-items:center; justify-content:flex-end; gap:8px; }
#${COL_ID} .pda-machine .sapMLabel { font-size:12px !important; color:#5b6b83 !important; }

/* Stavove tlacidla su hore nad boxom so zakazkou a sedia vo vlastnom farebnom
   banery. Zarovnanie s boxom ide cez margin-left a max-width (inline styl),
   vnutorne odsadenie ziadne - riadi si ho obsah sam. */
#${STATUS_ID} { background:linear-gradient(180deg,#eef3fa 0%,#dde7f4 100%) !important;
  border:1px solid #c9d7ea !important; border-radius:14px !important;
  padding:0 !important;
  box-shadow:0 2px 8px rgba(16,36,63,.10) !important; margin-bottom:10px !important;
  box-sizing:border-box !important; }

/* bezici cinnost vpravo (Vyroba - Vyroba / meno / cas / Zastavit) ako jemna pilulka */
.pda-aktivita { border-radius:14px !important; background:#e9f6ed !important;
  border:1px solid #bfe3ca !important; box-shadow:0 2px 8px rgba(16,36,63,.10) !important;
  padding:8px 12px !important; margin:0 !important; box-sizing:border-box; }
/*
 * Zastavit: tien aj pozadie musia byt na TOM ISTOM prvku. Predtym bol tien
 * na <button> a farebna plocha az na jeho vnutri (.sapMBtnInner), takze tien
 * obkresloval ine miesto nez viditelna pilulka a pri bielom pozadi to bilo
 * do oci. Teraz je oboje na .sapMBtnInner a tien je zladeny s cervenou.
 * Cervena #e53935 je rovnaka, aka sa pouziva inde v skripte.
 */
.pda-aktivita .sapMBtn { border-radius:10px !important; box-shadow:none !important;
  margin:0 !important; }
.pda-aktivita .sapMBtn .sapMBtnInner { border-radius:10px !important; padding:7px 18px !important;
  background:#e53935 !important; background-image:none !important;
  border:1px solid #e53935 !important;
  box-shadow:0 2px 6px rgba(229,57,53,.35) !important;
  /* bez flexu sedi ikona s textom pri lavom okraji, nie v strede */
  display:flex !important; align-items:center !important; justify-content:center !important;
  box-sizing:border-box !important; }
.pda-aktivita .sapMBtn .sapMBtnContent { justify-content:center !important; text-align:center !important; }
.pda-aktivita .sapMBtn .sapMBtnContent,
.pda-aktivita .sapMBtn bdi,
.pda-aktivita .sapMBtn .sapUiIcon { color:#fff !important; }

/* Operation Complete: na jeden riadok, nizsie a sirsie */
#${CONFIRM_ID} { height:auto !important; min-height:0 !important; max-height:none !important;
  width:auto !important; flex:0 0 auto; margin:0 !important; }
#${CONFIRM_ID} .sapMBtnInner { height:auto !important; min-height:0 !important;
  padding:10px 26px !important; white-space:nowrap !important; }
#${CONFIRM_ID} .sapMBtnContent, #${CONFIRM_ID} bdi { white-space:nowrap !important;
  font-size:14px !important; line-height:1.2 !important; }
`;
            document.head.appendChild(st);
        }

        /*
         * Vpravo v hlavicke drzime vlastny stlpec a do neho presuvame (zhora dole):
         *   okienko VYKRES  ->  tlacidlo Operation Complete  ->  prepinac Machine
         * Prepinac ma popisok "Machine" ako samostatny prvok pred sebou, preto sa
         * oba davaju do maleho riadku, nech drzia spolu.
         */
        function pravyStlpec(header) {
            let col = document.getElementById(COL_ID);
            if (!col) {
                col = document.createElement('div');
                col.id = COL_ID;
            }
            if (col.parentElement !== header) header.appendChild(col);
            return col;
        }

        function riadokPrepinaca(col) {
            let row = document.getElementById(MACHINE_ID);
            if (!row) {
                row = document.createElement('div');
                row.id = MACHINE_ID;
                row.className = 'pda-machine';
            }
            if (row.parentElement !== col) col.appendChild(row);

            const sw = document.getElementById(SWITCH_ID);
            if (!sw) return;
            const popisok = sw.previousElementSibling &&
                            /machine/i.test(sw.previousElementSibling.textContent || '')
                ? sw.previousElementSibling : null;
            if (popisok && popisok.parentElement !== row) row.appendChild(popisok);
            if (sw.parentElement !== row) row.appendChild(sw);
        }

        /*
         * Lavy okraj stavovych tlacidiel a pilulky POPIS OPERACIE zarovnany
         * s boxom hore (zakazka / material / production order). Odsadenie sa
         * nehada v pixeloch - odmeria sa priamo na stranke, takze to sedi aj
         * pri inej sirke okna. Zmena `style` nespusti DomWatch (sleduje len
         * pridavanie a mazanie prvkov), takze sa to nemoze rozkmitat.
         */
        function zarovnajVlavo() {
            const form = document.querySelector('#' + HEADER_ID + ' .sapUiForm');
            if (!form) return;
            const ciel = form.getBoundingClientRect().left;
            if (!ciel) return;

            /*
             * Baner s tlacidlami zarovnavame CELY (jeho vlastny lavy aj pravy
             * okraj) s boxom pod nim - nie prve tlacidlo v nom. Odkedy ma baner
             * viditelne pozadie, vycnieval by vlavo, aj keby tlacidlo sedelo.
             */
            const stav = document.getElementById(STATUS_ID);
            if (stav) {
                const sr = stav.getBoundingClientRect();
                posun(stav, 'marginLeft', ciel - sr.left);
                const sirkaBanera = Math.round(form.getBoundingClientRect().right - Math.max(sr.left, ciel));
                if (sirkaBanera > 200 && Math.abs(sirkaBanera - sr.width) > 2) {
                    stav.style.maxWidth = sirkaBanera + 'px';
                }
            }

            const opis = document.getElementById('__pda_opis_button__');
            if (!opis) return;
            posun(opis, 'marginLeft', ciel - opis.getBoundingClientRect().left);

            /*
             * Pilulka POPIS bola siroka na celu plochu a podliezala pravy stlpec
             * (Components / BOM, Parallel Process Handling). Stavove tlacidla
             * siahaju az popod ten stlpec (len su vyssie, takze sa nebiju), preto
             * nestaci zarovnat podla nich - hlada sa aj lavy okraj praveho stlpca
             * a berie sa to, co je viac vlavo.
             */
            const r = opis.getBoundingClientRect();
            let koniec = Infinity;

            const tlacidla = stav ? stav.querySelectorAll('.statusBtn') : null;
            if (tlacidla && tlacidla.length) {
                koniec = tlacidla[tlacidla.length - 1].getBoundingClientRect().right;
            }
            const stlpec = lavyOkrajPravehoStlpca(r.left);
            if (stlpec) koniec = Math.min(koniec, stlpec - 14);

            if (!isFinite(koniec)) return;
            const sirka = Math.round(koniec - r.left);
            if (sirka > 200 && Math.abs(sirka - r.width) > 2) opis.style.maxWidth = sirka + 'px';
        }

        /*
         * Zelene okienko s beziacou cinnostou vpravo (Vyroba - Vyroba, meno,
         * cas, Zastavit). Najde sa podla tlacidla "Zastavit" - jeho najblizsi
         * obal je to okienko.
         */
        function oznacAktivitu() {
            document.querySelectorAll('.sapMBtn').forEach((btn) => {
                const t = (btn.textContent || '').trim();
                if (!/^(zastavi|stop)/i.test(t)) return;
                const box = btn.closest('.sapMLIB') || btn.closest('.sapMVBox') ||
                            btn.closest('.sapMFlexBox') || btn.parentElement;
                if (box && !box.classList.contains('pda-aktivita')) box.classList.add('pda-aktivita');
            });
        }

        // lavy okraj praveho stlpca detailu (Components / BOM, Parallel Process Handling)
        function lavyOkrajPravehoStlpca(odX) {
            let x = Infinity;
            document.querySelectorAll('.sapMBtn, .sapMTitle, .sapMPanel').forEach((el) => {
                const t = (el.textContent || '').trim();
                if (!/^components\b|parallel process/i.test(t)) return;
                const er = el.getBoundingClientRect();
                if (er.width && er.left > odX + 150) x = Math.min(x, er.left);
            });
            return isFinite(x) ? x : null;
        }

        /*
         * Popisok a hodnota v hornom boxe nesedeli na jednej linke - popisok bol
         * vyssie. Presnu hodnotu odsadenia sa neda spolahlivo uhadnut (zavisi od
         * temy a velkosti pisma), preto sa rozdiel ich zvislych stredov odmeria
         * priamo na stranke a popisok sa o nho posunie. Po posune je rozdiel
         * nulovy, takze sa to samo ustali.
         */
        function zarovnajRiadky() {
            const form = document.querySelector('#' + HEADER_ID + ' .sapUiForm');
            if (!form) return;
            form.querySelectorAll('.sapMLabel').forEach((lab) => {
                const row = lab.closest('.sapUiFormElement') || lab.closest('.sapUiRespGridRow');
                if (!row) return;
                const val = row.querySelector('.sapMText, .sapMTextMaxLine, .sapMObjectNumberText');
                if (!val) return;

                const lr = lab.getBoundingClientRect();
                const vr = val.getBoundingClientRect();
                if (!lr.height || !vr.height) return;

                const rozdiel = (vr.top + vr.height / 2) - (lr.top + lr.height / 2);
                const teraz = parseFloat(lab.style.top) || 0;
                const nove = Math.round(teraz + rozdiel);
                if (Math.abs(nove - teraz) > 1) {
                    lab.style.position = 'relative';
                    lab.style.top = nove + 'px';
                }
            });
        }

        function posun(el, vlastnost, rozdiel) {
            if (!el || !isFinite(rozdiel)) return;
            const teraz = parseFloat(el.style[vlastnost]) || 0;
            const nove = Math.max(0, Math.round(teraz + rozdiel));
            if (Math.abs(nove - teraz) > 1) el.style[vlastnost] = nove + 'px';
        }

        /*
         * Riadok stavovych tlacidiel patri nad box so zakazkou, materialom
         * a vyrobnou zakazkou - operator ich ma mat hned pod nazvom pracoviska.
         * V appke su pod nim, takze ich posunieme na zaciatok Order_FlexBox.
         */
        function tlacidlaHore() {
            const stav = document.getElementById(STATUS_ID);
            const order = document.getElementById(ORDER_ID);
            if (!stav || !order || stav.parentElement !== order) return;
            if (order.firstElementChild === stav) return;
            order.insertBefore(stav, order.firstElementChild);
        }

        function apply() {
            const header = document.getElementById(HEADER_ID);
            if (!header) return;
            injectStyles();
            tlacidlaHore();

            const col = pravyStlpec(header);

            // 1) VYKRES uplne hore, cely vidno
            const wrap = document.getElementById(DRAWING_ID);
            if (wrap) {
                if (wrap.parentElement !== col) col.appendChild(wrap);
                if (!wrap.classList.contains('pda-v-riadku')) wrap.classList.add('pda-v-riadku');
            }
            // 2) tlacidlo Operation Complete
            const confirm = document.getElementById(CONFIRM_ID);
            if (confirm && confirm.parentElement !== col) col.appendChild(confirm);
            // 3) prepinac Machine uplne dole
            riadokPrepinaca(col);

            zarovnajVlavo();
            zarovnajRiadky();
            oznacAktivitu();
        }

        DomWatch.add(apply);
        onReady(apply);
    }

    /* ---------- 3.15 Priestorove kolacove grafy (casy SAP) ---------- */

    /*
     * Tri kolace pod popisom operacie (SAP Setup / Machine / Labor Time) su ploche.
     * Modul ich nechava presne take, ake su - nekresli ich odznova, nemeni data -
     * len ich vizualne nakloni (pohlad zboku) a prida tien, takze posobia
     * priestorovo. Pri prechode mysou sa naklon zmensi, akoby sa graf otocil k tebe.
     *
     * Kolace sa hladaju cez texty casov, ktore maju stabilne ID
     * (WorkcenterDetail--SetupTime_Text / MachineTime_Text / LaborTime_Text) -
     * kolac je v tom istom boxe. Vdaka tomu sa netrafi do casoveho grafu vlavo.
     */
    function modDonut3D() {
        const KOTVY = ['WorkcenterDetail--SetupTime_Text',
                       'WorkcenterDetail--MachineTime_Text',
                       'WorkcenterDetail--LaborTime_Text'];
        const MIMO = ['ResourceDetails', 'DialogChart'];
        const STYLE_ID = '__pda_donut3d_styles__';
        const TRIEDA = 'pda-3d';

        /*
         * Vyska (bocna stena) kolaca - lacno.
         *
         * Stena je znovu z `drop-shadow`, ale uz len z TROCH po 4 px namiesto
         * dvadsiatich po 1 px. Kazdy tien sa pocita na vysledok predchadzajuceho,
         * takze 20 tienov = 20 prekresleni obrazka (to mrazilo aplikaciu),
         * kym 3 x 4 px daju rovnako vysoku (12 px) a rovnako plnu stenu za tri.
         * Stena je plna preto, ze posun 4 px je oproti hrubke prstenca maly,
         * takze sa kopie prekryvaju.
         *
         * Druha - dolezitejsia - uspora: `filter` je v pokoji aj pod mysou
         * ROVNAKY a meni sa len `transform`. Prehliadac si tak filtrovany obrazok
         * odlozi a pri naklone uz len posuva hotovu tabulku (robi to graficka
         * karta). Predtym sa pri prechode mysou menil aj filter, takze sa cely
         * obrazok prefiltrovaval znovu v kazdom snimku.
         */
        const STENA = 3;      // kolko tienov (pozor: kazdy dalsi stoji prekreslenie)
        const KROK = 4;       // px na jeden tien -> vyska steny je STENA * KROK

        function injectStyles() {
            if (document.getElementById(STYLE_ID)) return;
            const stena = new Array(STENA).fill('drop-shadow(0 ' + KROK + 'px 0 rgba(12,28,55,.40))').join(' ');
            const st = document.createElement('style');
            st.id = STYLE_ID;
            st.textContent = `
/*
 * Stred otacania a zvacsovania je na HORNOM okraji (50% 0).
 * Predtym bol v strede, takze zvacseny kolac rastol aj nahor a pri pretoceni
 * do 2D (kde je najvyssi) siahal na cisla casov nad sebou - posuvanie nadol
 * to riesilo len cast. Takto rastie vylucne nadol a horny okraj sa nehne,
 * takze medzera nad kolacom ostava rovnaka pri kazdom zvacseni.
 */
.${TRIEDA} { transform:perspective(720px) rotateX(38deg) scale(var(--pda3d,1)) !important;
  filter:${stena} drop-shadow(0 12px 9px rgba(16,36,63,.28)) !important;
  transform-origin:50% 0% !important; margin-top:22px !important;
  transition:transform .24s ease !important;
  will-change:transform !important; }
/* pri prechode mysou sa kolac pretoci do skutocneho tvaru - filter ostava rovnaky */
.${TRIEDA}:hover { transform:perspective(720px) rotateX(0deg)
    scale(calc(var(--pda3d,1) * 1.06)) !important; }
/* aby naklonený kolac ani jeho stena neboli orezane okrajom boxu */
.pda-3d-box { overflow:visible !important; }
`;
            document.head.appendChild(st);
        }

        function box(el) {
            return el.closest('.sapMVBox') || el.closest('.sapMFlexBox') || el.parentElement;
        }

        /*
         * Na sirokej obrazovke su kolace zbytocne male - appka ich kresli stale
         * rovnako velke. Zvacsujeme ich zvacsenim (transform), takze sa nemeni
         * ziadny rozmer v rozlozeni a nic sa nikam neposunie; miesta okolo nich
         * je dost.
         */
        function zvacsenie() {
            const w = W.innerWidth;
            if (w >= 2200) return 1.9;
            if (w >= 1900) return 1.65;
            if (w >= 1650) return 1.4;
            if (w >= 1400) return 1.2;
            return 1;
        }

        function apply() {
            injectStyles();
            KOTVY.forEach((id) => {
                const text = document.getElementById(id);
                if (!text) return;
                const b = box(text);
                if (!b) return;
                if (!b.classList.contains('pda-3d-box')) b.classList.add('pda-3d-box');

                b.querySelectorAll('canvas, svg').forEach((g) => {
                    if (MIMO.indexOf(g.id) !== -1) return;
                    if (g.classList.contains(TRIEDA)) return;
                    // ikonky a drobnosti nechame tak, kolac je velky
                    const r = g.getBoundingClientRect();
                    if (r.width < 60 || r.height < 60) return;
                    g.classList.add(TRIEDA);
                });

                const z = String(zvacsenie());
                b.querySelectorAll('.' + TRIEDA).forEach((g) => {
                    if (g.style.getPropertyValue('--pda3d') !== z) g.style.setProperty('--pda3d', z);
                });
            });
        }

        DomWatch.add(apply);
        W.addEventListener('resize', apply);
        onReady(apply);
    }

    /* ---------- 3.16 Lavy panel na celu vysku obrazovky ---------- */

    /*
     * Povodne rozlozenie: cez celu sirku ide panel "Osobny stav", pod nim panel
     * pracoviska a az v nom vlavo zoznam zakaziek. Zoznam tak zacina az v polovici
     * obrazovky a hore vlavo je velka prazdna plocha.
     *
     * Modul to preskladá bez zasahu do appky - iba polohovanim:
     *   - lavy stlpec (zoznam zakaziek + graf) sa pripne nalavo od horneho
     *     baneru (HF logo, meno) az po spodok obrazovky
     *   - panel "Osobny stav" aj panel pracoviska dostanu zlava odsadenie,
     *     takze uz nezacinaju pri lavom okraji, ale az za tym stlpcom
     *
     * Nic sa nepresuva v DOM, len sa nastavuje poloha - ked sa modul vypne,
     * appka je presne taka, aka bola.
     */
    function modFullHeightLayout() {
        const LEFT_ID = 'WorkcenterDetail--LeftColumn_FlexBox';
        const STYLE_ID = '__pda_layout_styles__';
        const SIRKA = 430;      // sirka laveho stlpca (rovnaka ako v module zoznamu)
        const MEDZERA = 14;     // medzera medzi stlpcom a obsahom vpravo
        const SPODOK = 10;      // medzera pod stlpcom

        function injectStyles() {
            if (document.getElementById(STYLE_ID)) return;
            const st = document.createElement('style');
            st.id = STYLE_ID;
            st.textContent = `
.pda-layout-left { position:fixed !important; z-index:3 !important;
  width:${SIRKA}px !important; max-width:${SIRKA}px !important; flex:0 0 ${SIRKA}px !important;
  display:flex !important; flex-direction:column !important; box-sizing:border-box !important; }
.pda-layout-odsad { margin-left:${SIRKA + MEDZERA}px !important; }
/* box so zoznamom si vezme vsetko volne miesto, graf ostane taky, aky je */
.pda-layout-left > .pda-left-box:first-of-type { flex:1 1 auto !important; min-height:0 !important; }
.pda-layout-left > .pda-left-box:last-of-type { flex:0 0 auto !important; }
`;
            document.head.appendChild(st);
        }

        // panel, v ktorom su tlacidla osobneho stavu (Stretnutia / Prestavka / cakanie)
        function osobnyPanel(left) {
            const btn = Array.from(document.querySelectorAll('.statusBtn'))
                .find((b) => !left.contains(b) && b.closest('.sapMPanel'));
            return btn ? btn.closest('.sapMPanel') : null;
        }

        function apply() {
            const left = document.getElementById(LEFT_ID);
            if (!left || !left.offsetParent && left.style.position !== 'fixed') {
                // detail pracoviska nie je otvoreny - nic neriesime
                if (!left) return;
            }
            injectStyles();

            const osobny = osobnyPanel(left);
            const detail = left.closest('.sapMPanel');
            if (!osobny || !osobny.parentElement) return;

            // odsadenie oboch panelov vpravo od stlpca
            [osobny, detail].forEach((p) => {
                if (p && !p.classList.contains('pda-layout-odsad')) p.classList.add('pda-layout-odsad');
            });

            /*
             * Horny okraj stlpca = horny okraj panela "Osobny stav", teda presne
             * pod banerom s logom HF. Lavy okraj sa berie z rodica panela, nie
             * zo samotneho panela - ten sme prave posunuli doprava, takze by
             * sme merali uz posunutu hodnotu.
             */
            const rodic = osobny.parentElement.getBoundingClientRect();
            const hore = Math.round(osobny.getBoundingClientRect().top);
            const vlavo = Math.round(rodic.left + 8);
            if (!isFinite(hore) || hore < 0) return;

            if (!left.classList.contains('pda-layout-left')) left.classList.add('pda-layout-left');
            if (left.style.top !== hore + 'px') left.style.top = hore + 'px';
            if (left.style.left !== vlavo + 'px') left.style.left = vlavo + 'px';
            if (left.style.bottom !== SPODOK + 'px') left.style.bottom = SPODOK + 'px';
        }

        DomWatch.add(apply);
        W.addEventListener('resize', apply);
        onReady(apply);
    }

    /* ---------- 3.17 Menu HF Slovakia (pravy bocny panel) ---------- */

    /*
     * Zvisly panel pri pravom okraji s pripravovanymi funkciami HF Slovakia.
     * Vpravo od obsahu appky (za blokom Components / BOM) je volne miesto,
     * takze panel nic neprekryva a hlavne: je to NAS vlastny prvok pripnuty
     * k oknu - z rozlozenia appky sa nic nevybera a nic sa nepresuva.
     * (Pokus pripnut lavy stlpec appky v 1.22.0 rozlozenie rozhodil, preto
     * sa tu appky nedotykame vobec.)
     *
     * Tlacidla zatial nic nerobia - po kliknuti sa ukaze okno "vo vyvoji".
     * Funkcie sa budu doplnat postupne, kazda do svojej vetvy v `spusti`.
     */
    function modHfMenu() {
        const PANEL_ID = '__pda_hf_menu__';
        const OVERLAY_ID = '__pda_hf_overlay__';
        const STYLE_ID = '__pda_hf_menu_styles__';
        const SIRKA = 250;
        const OKRAJ = 10;
        const TAB_ID = '__pda_hf_tab__';
        const KEY_OTVORENY = 'pda_hfmenu_open_v1';
        const PRAH = 1600;      // od akej sirky okna je panel predvolene otvoreny

        /*
         * Na uzkej obrazovke panel prekryval stavove tlacidla, preto sa da
         * schovat: vpravo ostane uzky pasik a klik na neho panel vysunie.
         * Predvolene je otvoreny len na sirokej obrazovke; kedze si volbu
         * pamatame, na dielni sa nastavi raz a drzi aj po obnoveni stranky.
         */
        let otvoreny = null;

        function nacitajStav() {
            try {
                const v = GM_getValue(KEY_OTVORENY, null);
                if (v === '1') return true;
                if (v === '0') return false;
            } catch (e) { /* ignore */ }
            return null;
        }

        function ulozStav(v) {
            try { GM_setValue(KEY_OTVORENY, v ? '1' : '0'); } catch (e) { /* ignore */ }
        }

        const LOGO_HF = 'data:image/jpeg;base64,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';
        const LOGO_FLEXUS = 'data:image/jpeg;base64,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';
        const FOTO_ELA = 'data:image/jpeg;base64,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';

        const POLOZKY = [
            // nazvy a podnadpisy podla navrhu dizajnu z 2026-09-16
            { ikona: '📄', nazov: 'Výkresy', popis: 'archív', akcia: 'vykresy' },
            { ikona: '⚙', nazov: 'CHIPS', popis: 'nástroje' },
            { ikona: '👤', nazov: 'Majster', popis: 'prihlasovanie' },
            { ikona: '📦', nazov: 'Materiál', popis: 'objednať' },
            { ikona: '🔧', nazov: 'TOOLSHOP', popis: 'výdaj' },
            { foto: FOTO_ELA, nazov: 'Ela', popis: 'dokumentácia' },
            { obrazok: LOGO_FLEXUS, nazov: 'Flexus', lenObrazok: true },
        ];

        function injectStyles() {
            if (document.getElementById(STYLE_ID)) return;
            const st = document.createElement('style');
            st.id = STYLE_ID;
            st.textContent = `
/* Nizke z-index je zamerne: panel ma byt nad obsahom stranky, ale POD
   oknami appky (detail grafu, tabulka stavov). S vysokym cislom prekryval
   otvorene okno a nestmavil sa spolu so zvyskom stranky. */
#${PANEL_ID} { position:fixed; right:${OKRAJ}px; width:${SIRKA}px; z-index:5;
  display:flex; flex-direction:column; gap:8px; box-sizing:border-box;
  font:13px/1.35 -apple-system,"Segoe UI",Roboto,sans-serif; color:#13315c;
  padding:10px; background:#fff; border:1px solid #dfe4ec; border-radius:16px;
  box-shadow:0 4px 18px rgba(16,36,63,.12); overflow-y:auto; justify-content:flex-start; }
#${PANEL_ID} .hf-logo { width:100%; border-radius:10px; display:block; }
#${PANEL_ID} .hf-nadpis { font-size:10.5px; font-weight:800; letter-spacing:.14em;
  text-transform:uppercase; color:#8e9bb0; text-align:center; margin:2px 0 4px; }
#${PANEL_ID} .hf-btn { display:flex; align-items:center; gap:10px; width:100%; box-sizing:border-box;
  padding:10px 12px; border:2px solid #c9d7ea; border-radius:12px;
  background:linear-gradient(180deg,#f7faff 0%,#eaf1fb 100%); cursor:pointer; text-align:left;
  font:inherit; color:#13315c; box-shadow:0 2px 6px rgba(16,36,63,.10);
  transition:transform .13s ease, box-shadow .13s ease, border-color .13s; }
#${PANEL_ID} .hf-btn:hover { transform:translateY(-3px); border-color:#2563eb;
  box-shadow:0 10px 20px rgba(16,36,63,.22); }
#${PANEL_ID} .hf-btn:active { transform:translateY(-1px); box-shadow:0 3px 8px rgba(16,36,63,.18); }
#${PANEL_ID} .hf-btn .ik { font-size:22px; line-height:1; flex:0 0 auto; width:26px; text-align:center; }
#${PANEL_ID} .hf-btn .ik img { width:100%; display:block; border-radius:3px; }
/* fotka Ely je okruhla a o nieco vacsia ako ikonky */
#${PANEL_ID} .hf-btn .ik.foto { width:34px; }
#${PANEL_ID} .hf-btn .ik.foto img { border-radius:50%; border:2px solid #c9d7ea; }
#${PANEL_ID} .hf-btn .tx { flex:1 1 auto; min-width:0; }
#${PANEL_ID} .hf-btn .n { display:block; font-size:13px; font-weight:700; }
#${PANEL_ID} .hf-btn .p { display:block; font-size:11px; color:#6b7c95; margin-top:1px;
  overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
/* tlacidlo, na ktorom je len logo - bez textu */
#${PANEL_ID} .hf-btn.hf-obr { justify-content:center; padding:10px 14px; }
#${PANEL_ID} .hf-btn.hf-obr img { width:100%; max-width:175px; display:block; border-radius:4px; }

/* Zatvaranie a otvaranie je posunom za pravy okraj, nie zmiznutim - je tak
   vidiet, kam sa panel podel. Viditelnost sa prepne az na konci prechodu,
   aby schovany panel nechytal kliknutia. */
#${PANEL_ID} { transition:transform .28s ease, opacity .28s ease, visibility 0s linear .28s; }
#${PANEL_ID}.skryty { transform:translateX(calc(100% + ${OKRAJ + 6}px)); opacity:0;
  visibility:hidden; pointer-events:none; }
/* uzky pasik na pravom okraji - klikom sa panel vysunie alebo schova */
#${TAB_ID} { position:fixed; z-index:5; width:26px; padding:14px 0; cursor:pointer;
  transition:right .28s ease;
  display:flex; flex-direction:column; align-items:center; gap:8px;
  background:#13315c; color:#fff; border:0; border-radius:10px 0 0 10px;
  box-shadow:-2px 2px 10px rgba(16,36,63,.25); font:inherit; }
#${TAB_ID}:hover { background:#1c478a; }
#${TAB_ID} .sip { font-size:13px; line-height:1; }
#${TAB_ID} .txt { writing-mode:vertical-rl; text-orientation:mixed; font-size:10.5px;
  font-weight:800; letter-spacing:.10em; white-space:nowrap; }

#${OVERLAY_ID} { position:fixed; inset:0; background:rgba(10,20,40,.5); z-index:100002;
  display:flex; align-items:center; justify-content:center; padding:24px; box-sizing:border-box; }
#${OVERLAY_ID} .karta { background:#fff; border-radius:16px; width:min(460px,92vw); overflow:hidden;
  box-shadow:0 24px 70px rgba(16,36,63,.4); font-family:-apple-system,"Segoe UI",Roboto,sans-serif; }
#${OVERLAY_ID} .hl { display:flex; align-items:center; gap:12px; padding:14px 18px;
  background:#13315c; color:#fff; }
#${OVERLAY_ID} .hl .n { font-size:15px; font-weight:800; letter-spacing:.06em; text-transform:uppercase; }
#${OVERLAY_ID} .hl .x { margin-left:auto; background:none; border:0; color:#fff; font-size:28px;
  line-height:1; cursor:pointer; padding:0 4px; }
#${OVERLAY_ID} .telo { padding:22px 20px 24px; font-size:15px; line-height:1.6; color:#17202e; text-align:center; }
#${OVERLAY_ID} .telo .velke { font-size:38px; display:block; margin-bottom:10px; }
#${OVERLAY_ID} .telo .co { font-weight:700; color:#13315c; }
#${OVERLAY_ID} .telo.hladat { text-align:left; }
#${OVERLAY_ID} .riadok { display:flex; gap:8px; margin-top:12px; }
#${OVERLAY_ID} input.pole { flex:1 1 auto; min-width:0; padding:10px 12px; font-size:16px;
  border:2px solid #c9d7ea; border-radius:10px; color:#13315c; box-sizing:border-box; }
#${OVERLAY_ID} input.pole:focus { outline:none; border-color:#2563eb; box-shadow:0 0 0 3px rgba(37,99,235,.18); }
#${OVERLAY_ID} button.hladaj { flex:0 0 auto; padding:10px 20px; font-size:15px; font-weight:700;
  border:0; border-radius:10px; background:#13315c; color:#fff; cursor:pointer;
  box-shadow:0 2px 8px rgba(16,36,63,.22); }
#${OVERLAY_ID} button.hladaj:hover { background:#1c478a; }
#${OVERLAY_ID} .napoveda { font-size:12.5px; color:#6b7c95; margin-top:10px; }
`;
            document.head.appendChild(st);
        }

        function okno(nazov) {
            const stary = document.getElementById(OVERLAY_ID);
            if (stary) stary.remove();

            const overlay = document.createElement('div');
            overlay.id = OVERLAY_ID;

            const karta = document.createElement('div');
            karta.className = 'karta';

            const hl = document.createElement('div');
            hl.className = 'hl';
            const n = document.createElement('span'); n.className = 'n'; n.textContent = 'HF Slovakia';
            const x = document.createElement('button'); x.className = 'x'; x.type = 'button';
            x.textContent = '×'; x.setAttribute('aria-label', 'Zavrieť');
            hl.appendChild(n); hl.appendChild(x);

            const telo = document.createElement('div');
            telo.className = 'telo';
            const ik = document.createElement('span'); ik.className = 'velke'; ik.textContent = '🚧';
            const co = document.createElement('div'); co.className = 'co'; co.textContent = nazov;
            const txt = document.createElement('div');
            txt.textContent = 'Táto funkcia je vo vývoji. Zatiaľ nie je funkčná — pripravujeme ju.';
            telo.appendChild(ik); telo.appendChild(co); telo.appendChild(txt);

            karta.appendChild(hl); karta.appendChild(telo);
            overlay.appendChild(karta);

            const zavri = () => { overlay.remove(); document.removeEventListener('keydown', naEsc); };
            const naEsc = (e) => { if (e.key === 'Escape') zavri(); };
            x.addEventListener('click', zavri);
            overlay.addEventListener('click', (e) => { if (e.target === overlay) zavri(); });
            document.addEventListener('keydown', naEsc);
            document.body.appendChild(overlay);
        }

        /*
         * Zadanie cisla vykresu. Samotne hladanie a zoznam najdenych suborov
         * uz robi okno z modulu vykresu (`shared.pdmOpenDialog`) - je to to iste
         * okno, ake sa otvori kliknutim na okienko VYKRES pri zakazke, takze
         * sa spravanie nikde nerozchadza.
         */
        function oknoVykresy() {
            if (typeof shared.pdmOpenDialog !== 'function') {
                okno('Výkresy');
                return;
            }

            const stary = document.getElementById(OVERLAY_ID);
            if (stary) stary.remove();

            const overlay = document.createElement('div');
            overlay.id = OVERLAY_ID;

            const karta = document.createElement('div');
            karta.className = 'karta';

            const hl = document.createElement('div');
            hl.className = 'hl';
            const n = document.createElement('span'); n.className = 'n'; n.textContent = 'Výkresy';
            const x = document.createElement('button'); x.className = 'x'; x.type = 'button';
            x.textContent = '×'; x.setAttribute('aria-label', 'Zavrieť');
            hl.appendChild(n); hl.appendChild(x);

            const telo = document.createElement('div');
            telo.className = 'telo hladat';
            const popis = document.createElement('div');
            popis.textContent = 'Zadaj číslo výkresu alebo číslo materiálu:';

            const riadok = document.createElement('div');
            riadok.className = 'riadok';
            const pole = document.createElement('input');
            pole.className = 'pole';
            pole.type = 'text';
            pole.placeholder = 'napr. 1-60.2-06.66-009 alebo 25217930';
            const hladaj = document.createElement('button');
            hladaj.className = 'hladaj';
            hladaj.type = 'button';
            hladaj.textContent = 'Hľadať';
            riadok.appendChild(pole); riadok.appendChild(hladaj);

            const napoveda = document.createElement('div');
            napoveda.className = 'napoveda';
            napoveda.textContent = 'Zobrazí sa zoznam dostupných výkresov — kliknutím na riadok sa výkres otvorí.';

            telo.appendChild(popis); telo.appendChild(riadok); telo.appendChild(napoveda);
            karta.appendChild(hl); karta.appendChild(telo);
            overlay.appendChild(karta);

            const zavri = () => { overlay.remove(); document.removeEventListener('keydown', naEsc); };
            const naEsc = (e) => { if (e.key === 'Escape') zavri(); };
            const odosli = () => {
                const cislo = pole.value.trim();
                if (!cislo) { pole.focus(); return; }
                zavri();
                shared.pdmOpenDialog(cislo, { titul: cislo });
            };

            x.addEventListener('click', zavri);
            overlay.addEventListener('click', (e) => { if (e.target === overlay) zavri(); });
            hladaj.addEventListener('click', odosli);
            pole.addEventListener('keydown', (e) => { if (e.key === 'Enter') odosli(); });
            document.addEventListener('keydown', naEsc);

            document.body.appendChild(overlay);
            setTimeout(() => pole.focus(), 30);
        }

        // sem budu pribudat skutocne funkcie - kazda si najde svoju vetvu podla `akcia`
        function spusti(polozka) {
            if (polozka.akcia === 'vykresy') { oknoVykresy(); return; }
            okno(polozka.nazov);
        }

        function tlacidlo(p) {
            const b = document.createElement('button');
            b.type = 'button';
            b.className = 'hf-btn';

            // Flexus ma na tlacidle len svoje logo, ziadny text
            if (p.lenObrazok && p.obrazok) {
                b.classList.add('hf-obr');
                const img = document.createElement('img');
                img.src = p.obrazok;
                img.alt = p.nazov;
                b.appendChild(img);
                b.title = p.nazov;
                b.addEventListener('click', () => spusti(p));
                return b;
            }

            const ik = document.createElement('span');
            ik.className = 'ik';
            if (p.foto) {
                ik.classList.add('foto');
                const img = document.createElement('img');
                img.src = p.foto;
                img.alt = p.nazov;
                ik.appendChild(img);
            } else if (p.obrazok) {
                const img = document.createElement('img');
                img.src = p.obrazok;
                img.alt = p.nazov;
                ik.appendChild(img);
            } else {
                ik.textContent = p.ikona;
            }

            const tx = document.createElement('span');
            tx.className = 'tx';
            const n = document.createElement('span'); n.className = 'n'; n.textContent = p.nazov;
            tx.appendChild(n);
            if (p.popis) {
                const s = document.createElement('span'); s.className = 'p'; s.textContent = p.popis;
                tx.appendChild(s);
            }

            b.appendChild(ik); b.appendChild(tx);
            b.addEventListener('click', () => spusti(p));
            return b;
        }

        function postav() {
            const panel = document.createElement('div');
            panel.id = PANEL_ID;

            const logo = document.createElement('img');
            logo.className = 'hf-logo';
            logo.src = LOGO_HF;
            logo.alt = 'PDA App Extension — developed by HF Slovakia';
            panel.appendChild(logo);

            const nad = document.createElement('div');
            nad.className = 'hf-nadpis';
            nad.textContent = 'PDA';
            panel.appendChild(nad);

            POLOZKY.forEach((p) => panel.appendChild(tlacidlo(p)));
            document.body.appendChild(panel);
            return panel;
        }

        /*
         * Panel patri LEN na obrazovku otvoreneho pracoviska - na uvodnej
         * obrazovke (prehlad pracovisk) nema co robit. Pozname ju podla toho,
         * ci je vidiet panel s pracovnym zoznamom.
         */
        function panelPracoviska() {
            const left = document.getElementById('WorkcenterDetail--LeftColumn_FlexBox');
            const p = left && left.closest('.sapMPanel');
            if (!p) return null;
            const r = p.getBoundingClientRect();
            return r.height > 0 && r.width > 0 ? p : null;
        }

        function pasik() {
            let t = document.getElementById(TAB_ID);
            if (t) return t;
            t = document.createElement('button');
            t.id = TAB_ID;
            t.type = 'button';
            const sip = document.createElement('span'); sip.className = 'sip';
            const txt = document.createElement('span'); txt.className = 'txt'; txt.textContent = 'HF SLOVAKIA';
            t.appendChild(sip); t.appendChild(txt);
            t.addEventListener('click', () => {
                otvoreny = !otvoreny;
                ulozStav(otvoreny);
                apply();
            });
            document.body.appendChild(t);
            return t;
        }

        function apply() {
            injectStyles();
            const detail = panelPracoviska();
            const panel = document.getElementById(PANEL_ID);
            const t = document.getElementById(TAB_ID);

            // mimo detailu pracoviska nie je vidiet ani panel, ani pasik
            if (!detail) {
                if (panel) panel.style.display = 'none';
                if (t) t.style.display = 'none';
                return;
            }

            if (otvoreny === null) {
                const ulozene = nacitajStav();
                otvoreny = ulozene === null ? W.innerWidth >= PRAH : ulozene;
            }

            const p = panel || postav();
            p.style.display = '';
            p.classList.toggle('skryty', !otvoreny);

            const tab = pasik();
            tab.style.display = '';
            tab.title = otvoreny ? 'Skryť panel HF Slovakia' : 'Zobraziť panel HF Slovakia';
            tab.querySelector('.sip').textContent = otvoreny ? '▶' : '◀';
            tab.style.right = otvoreny ? (SIRKA + 2 * OKRAJ) + 'px' : '0px';

            /*
             * Zaciatok je zarovnany s panelom pracoviska (teda pod tlacidlami
             * Stretnutia / Prestavka) a dole siaha az k spodku okna, takze
             * vyplni celu volnu plochu vpravo.
             */
            const hore = Math.round(detail.getBoundingClientRect().top);
            if (hore > 0 && Math.abs((parseFloat(p.style.top) || 0) - hore) > 4) {
                p.style.top = hore + 'px';
                p.style.bottom = OKRAJ + 'px';
            }
            if (hore > 0) tab.style.top = (hore + 12) + 'px';
        }

        DomWatch.add(apply);
        W.addEventListener('resize', apply);
        onReady(apply);
    }

    /* ---------- 3.18 NOVY DIZAJN (vetva ver.2) ---------- */

    /*
     * Prezlecenie PDA do dizajnu HF Slovakia podla navrhu z 2026-09-16.
     *
     * Zasada, na ktorej je to postavene: NIC z appky sa nemaze ani nevybera
     * z toku stranky. Vsetko su len styly a par vlastnych prvkov vlozenych
     * medzi povodne. Tlacidla, prepinace a zoznamy ostavaju tie ISTE prvky
     * appky, takze vsetky volania a funkcie fungujú presne ako predtym -
     * menia sa iba farby, tvary a popisy.
     * (Pokus pripnut lavy stlpec cez position:fixed vo vetve 1.22 rozlozenie
     * rozhodil - preto sa tu nic nepolohuje natvrdo.)
     *
     * Cast 1: pozadie a karty, horny pruh s logom, stavove tlacidla ako karty
     * s ikonou a podnadpisom, nadpisy sekcii, patka.
     */
    function modNewDesign() {
        const STYLE_ID = '__pda_nd_styles__';
        const BODY_CLASS = 'pda-nd';
        const FOOTER_ID = '__pda_nd_footer__';
        const LOGO_HF_MALE = 'data:image/jpeg;base64,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';
        const POZADIE = 'data:image/jpeg;base64,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BRRRQAUUUGgApKKKAFopKWgAooooAKKKKAClpKKAFooooGFLSUUALSUUtABRSUUAFLSUUgFooooAKKKKAFopKKAFopKKAFooooAKWkooAWlzSUZoAWikooAWiikzSAWiiigYUtFFABRRRQAUtJRQAtFFFAC0UlLQAUUUUgFopKWgYUUUUCClpKKAFopKKAFopKWgAooooAKKKKBhRRRQAUUUUAFFFFABRRRQAUUUUALSUUUAFFFFAC0UUUAFFFFABRQaKADNFFFABRRRQACiij8aACiiigAooooAKKKKAFopKWgAooopAFFHWigAqG5k2JsB5b9BUjsEQse1UWcuxZupqoq4mxKKKK0MwooooAKKKKQwoxS0dKAEpaKSgBaKMUUAFLikpaAFxRikooAWiikoAWiigUAJRRRQAUUUUAFJilooASilooASilopiEoxS0GkMSiiigApaKKACiiimImpKKKksKKKSgAoopKAFpKKKBBRRSUAFFHekoAXvRSd6KBBSUUUwCkoooASiiimIKSiigYGkoNFABSUtJQIKQ0tJTASiiigBKKKQ0wCq8rb246CpJHwuB1NQ1SJY0iiloqhDelFLSUAIaQ0tJTEJSHrTqkt4PPmCnhByx9BRsBNap5MJmP35BhPZe5pafI298gYHQD0FMqChKKKKYhKSlpKAEpDSmkpgFFFLigQlaNjF5UJuGHzNwg/rVW3gNxOqDp1Y+grQmYM2F4VeAKib6FxXUZ1pKWkJrMoTNNILsET7zHApScVJANkZlP3n4T2HrT2Ac+FAjT7qDAptFFSMKKKKACiikoAKKKKADNFFFAC0UlFABRRRQAUtJRQAtLSUUALRSUUAFLSUUDFopKWgAooooAKKKKAFopBS0gCiiigAoo70UAFFFFABS0lFAC5opKWgBaKSloAKKKKACiiikAtFJS0wClpBRSGLRSUUCFooooGLSUUUALS0gpaADNFFFABRRRSGLRRRQIWikpaACkxS0lABRRRQAUtJS0DA0UUUAJS4oo7UAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFAC5pKKKAFoFFJQAtBoooAKKTvS0AFFFJQAtFFH0oAKSlooABRRRQAtFJRQAtFFFABRRUU8nlpgfeagCG4k3vtH3V/U1DS0VqtDMSilpKACilopAJS4oxRQAUUUUAGKKWig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lFAC0UUUAFFFFAwooooAKKKKACiiloAKKKKAFpKWmSvsTI6npQBFcSZOwdB1qGj+dFaJWICiiigQUUUtAwooopAJS4oooAKSlpKYBRRS0AJRS0UAIKWiigApaKKACiiikAUlLSUwFooooAKKSigBaTFFFIAFGKBRTAKKSloAKKSloAMUUUUABp5cW8BkP324WiJN7c8KOSaq3E3nTFh90cKPahK7sLYZn1ooorUgWikooAuUlFFYmgUUUUAFFFFAAKKKKAEopaKAEpKWigBKDRQaYhKKKKAEoooNACUUUUAFJS0lACGkpT0pKYgpKKSmAtJRTWOBgdaAGucn2plLSVRIlFFBpgJSGlpDQIbRS000wCpraMFjK4+RO3qewqFVZ2CqMknAFXHARViT7q9/U9zSfYaGlizEk8nk0lFJQAGkopKACkpaSmIKSlopgIOtLSUqqXYKoyScAUCLNlB5su9x+7j5PufSrLuZHLHvTiqwQrAnblj6mmVi3d3NUrKwUhpaQ0gEzjrUqDZFu/ifp7CmIgd8H7o5b6UrvvYmh9gQlFFFABRRRQMKKKTNAC0UlLSAQ0UUUwCiiigApaTNLQAUUUUgA0UUUAFFLSUALSUUUAHWlpKKAFopM0uaACiiigAooooGLRSUUALRRRQAUUUtABRRRSAKKKKAClpKM0AFLRRQAUUUUAFAxRijFAC0lLSUAGaWkxS0AGKKKKACiiigYtGKSloAKWkooAWikpaQBS5pKKAFooooAKWkpaAFopKWgAoo7UUgFpKKKACiiigAooooAKKKKACiiimAUUUUgCiiimAUtJRSAKKKKACiiigAo70UUAFLSUtACUUUUDCiiigQtJS0lAC0UlLQAUUUUDClpKKACloooAM4GT0FVJHMj57dqknk52A/WoapIlsKKKKokKKKWkMBRRRQAUtFFABRRRQAUmKXFFACYopaKAE5opaKACikpaACiiigAooooAKKKKACiiigAoopKAFooooASiiigAooopgFFFFAgpfpSVKm2NGmfovT3NJjRHdP5MQgU/M3Ln+lUhSuxkcsx5JzSCtUrIhu4tLRRTELRRRQBaooorI0CiiikAUUtFACUUtJQAUUUUAFJS0maBBSGjNITTAKKTNGaACg0ZpCaAA0UmaTNMBaKTNJuoAU02gtTS1MQ6kpN1JuoAXOKjJyc0rH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AooooAKKKKAClpKWgAoozRQAUUUUAFFFFAC0UlLmgApaSigAooooAKWiigYUtJSikAZooooAKKKKAFopKKAFooooAKKKKACiiigAopKWgAooooAKWkooAKWikoAKKMUtACUUUUAFLSUUAFFFFABRRmigApc0lFAC0tJRQAtFFNdti579qAGTP/APxqKk6nJ60tWSFLSUtIAooooASloooAKSlooAKKKWgAoopaBhmikooAWiiikAUUUUAFFFFABSUUUwCiijFABRRRigAooooAKKKKAEooooEFFFFABRRRQAUqqWYAd6Snu/2aAv/G/CigCC9lGRCn3U6+5qpQeuTRWyVlYhu4UUUZpiFopM0ZpAL0opuaRmwMetMAJyc0U3NG6mSLRTS1FAGnRRRWBsFFFBoASlpKWkAlFFFMAoopKBBSUtIaAEooopgJRRQaAEpKWkoEBpM0tIaYCE4ptKetJQAUlFJTAM0hpaSmISig0lAAaaadTokDNub7q/qfSgRKi+XFz95+vsKTFKx3Eknk0lIYlFGaKYCZozRSGgApKXNJTASijFTW9q05znag6tRe24rXGwQNPJtXgDq3pV75UQRxcIO/8AePrRlUTy4htQfmabWTdy0rBRRRSGFFFFAB7UlFFMAooooAKSlpKACiiigAooooAMUUUUCCkpaKAEopaKAEpaSigYUtJmloEJS0UUDCiiigApaSloAKKKKACikpaACiiigAooooAKM0UUgCiiigApaSigBaKSloAKKKKAFopKWgYUUlLQAUUUUALRSUUgFopKWgAopKWgAooooAKXNJRQAtJRRQAUtJRQAuaKKKACiiigApaSloAKKKKAFpKKKBhSikpRQIM0ZoopDClpKKAFzRSZpaACiiigBaKSigAooooAWikpaACiiigApaSigBc0lFFABRSUUALRSUUALRSZooAWikFLQAlFLSUALRSUtAC0UUUAHQZqu772z27U+Z/4R+NRAVSQmLRRRTEFLRRSAKKKKACiiigAoopaBhRRRQAUtFFABRRRSAKKKKACiiigAooooAKSlooASiiimAUUUUAFFFFABRRRQAlFFFAgooooAKKKBycCgB8SBmy3CjkmqVzOZ5i38I4Ue1WbyXyYhAp+ZuXqhmrgupMn0FopM0VoSGaM0maaTQIdRmmFqTdQA8tgZ7VCXyc0kj9h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AiaNBGgUfjTqKKzLENFFFACUUtJQAlFKaSgApKWkpgGKKKKACiikoAKKKKBBRRRQAUUUUAJRS0lABRRRQAUhpaKYCUUtJQAUUd6KACiiigAxRRRigAoopKAFooooEFFFLQMKKKSgAooooAWkoooAKWiigBKKKKADFFLSdKAFpKKWgBKMUtJQIBRRRQMKKKWgBKWikpAFAooxTAWiiigAooooAWiiikAUUlFAC0UUUAFFFFAwooooELSUUUAFLSUUALSUtFAwooooAKKKKACiiigApaSlpAFFFHSgBaKSigAooooAWiiigAoFFFAC0UmaKAFopKWgAooooAM0UUUAJS0UUAFFFJQA6kozRQAUUUlAC0ZoooAKKKKAFopKWkAUlLTWbauaYDZG42io6OScmiqJClpKKAFpaSigYUtJRQAtLSUUgFooooAKKKKAFooooGFFFFIBRRSZozQAtJSZozTAdRmkozQAtFJRSAKKKKAEooopgFFFBoENaomqQmomFMCI9aSnEUlMQlIaXpSUCCgUlKMkgAZJOAKAJ7eNXcvKcRxjcxrKvLpru6aVuB0Uegq7qc4ghWyjOT96Uj9BWYBVwXUUn0AU4UAUuKsgcKWmiloGOopM0ZoAWkpc0UwEppFPNMY0xEbUwmnMaYaBCZpM0hpuaAHg08VGtSjpQAUUGimI6eiiiuQ6QooooASig0UAIaKKKBCUlOpDTAbSGlpDQIKSikpgFJS0lAhKKKKYBSGikoATNIaWkNMQ6KMySAdu5rQAAGBwBUcEflR8/ePWpKzbuaJBSUtJSATiiiigApKWigBtFLSUwCkpaSgAooooAKKSigAooooEFFFFABRRRQAlBoooAKKKKACiiigAooooAKKKKAEooopgFFFFABRRRQAUYpaKAEpaQ0UAFFLSUAFFFFABRRRQAUtFFACUUUUAFFFFABRRS0CEooooGFFLRQAlFLRQAlLRRSAKKKKADtRS/SkoAKKKKAFopKKACloooAKKKKACikooAWiiigAooooAKKKKAClpKKACloooGFFFFABRRRQAtFFFIAooooAKKKKAClpKKYC0YpKWkAUUUUAFJRRQAtGaKKACjvRRQAClpKWgAoopKAFpKSloAKWkpaACiikoADRS0lAC0UlLQAUtFFABULNuPtTpG4wKjpoTFpKKKYgooooAKWijFIYUuKKKADFFLRQAUUUUAFFFGaACikJpM0AOozTC1JuoAfupC1M3U0tQBJmjdUW+jdTAm3U4GoQakU0gH0UmaWgYUUtJSAKSlopgIabTjTaBCGmEU800igCIrTTUxFMK1QiKkpzLim4oAQ1LG62lu93IM7eIwe7U2KIzShF79T6CqeqXKzziKL/Uw/Kvue5ppXdhbalNnaSRnc5ZjkmnrTAKeorYzHgZpcUDilzSGJSGlJqNmpiF3UbsVEz03fQBYDU7dVUPinCSgLk5amk5pgbNLmgQhGaaVqQCgrTAgIppHNTlaYUoAYoqUChUqTZQBGaKeVoouFjpKWkorlOgKSlpKACkpaSgQUUUUAJSUppDTEIaQ0tIaAEpKU0lMApDQaQ0CA0lFJTAKSlooASpraLc29hwOn1qJULuFXvWgqhECqOBSkxpBRS0lZlhSUtJTEJRS0lABSUtBoASkpaSgApKWimAlFFFABRRRQAlFLSUCCiiigAoopKAFopO9LQAlFGaKACiiigAooooAKKKKAEopaMUAJRS0lMAoopaACkpaKQCUtFFMBKWiikAlFFFMAoFFFABRRRQAUUUUCCloooGFFFFABRRRQAUUhooAKWkpaQBRRRQAUUUUALSUZozQAUUUUAFFFFABS0lAoAWkoxRTAMUtJS0gCiiigAooooAKKKKACiiigBaKSloGFFFJQAtFJS0AFFFFABS0lFAC0UUUAFFJS0gCiiigAoopKAFooozTAKKKKQBS0lLmgAozRSUALRSZooAKWkooAWij60fWgAoopKAFoo7UUAFLSUtABSM21c0VE7bjx0FCQhucnJpaSiqELSUUUALRSUtAC0tJRSGLRRS0AFFFIaADNGaQ00mgB2aTNNLUhamA4mm7qYWppagRIWFMLUzNJTAeXppammkoAdupQ1MozzQImDVKrVVBqRWpWGWAacDUG/inb6BkwPpS5qEPTg9ICSkpAaXNAB2ptLSYoAXFIRS0UANIpjCpCKYRQBEwph4qYrT4YlLF5eI4xljTuKxXupjZWW1eJ5xj/dWsard3I1zctK3fgD0FVyuK1irIzbuIKcDzTM0hfFWIlzRuqDzKPMoAlZqid/ekL5qNqBAzUwtSMabmmIkDUu+os0ZoAnD81KrZqoGqZCaBlpadimJUqikMaVpNtSYoxQA0LTsUuKKAGkUUMeKKYjfopKWuU6AopKKACkpaSgQUUUUAJSGlpDTASkNBpKYgpKWkoASkNLRQA2kpaSmIKKKkgi8yTn7o60AT20exdx6n+VTUtJWb1LQUlLSUhhSUtJQIKSiimAUhpaKAEpKXFJQAUlLRTASiiigAooooAKSiigQUUUUAFJS0UAFJS0mKACiiigAoo70UAFFFFABRSUUAFLSUtMApDS0lABRRRQAUUUUALRRRQAUUUUAJRS0UAJRRQKACiiigApaSl6UAJS0ZooASloooAKKO9FACUUtJQAUUUUgFooopgFFJS0AFFJRQAUtJS0gClpM0UAFFFFABS0UUAGKKKKACkoooAKWkooAWiiigAooooAM0UUUAFFFFABRRRQAtFFFAwooooAWikopAFFFLTAKKKKACik7UtABSUUUAFLSUtIAooooAKKKKACiiigApRSUUALSUUUAFFFFABRRS0AFGaKazBQSaAEkfAwOtRUE5OTRVEhRRRQAUtJS0AFKKBS0AKKKBS0hhRRRQAUhNBppNACE0wtihjUTNTEOL5ppam5pM0wHFqSkzRQIM0UmaTNAC0UmaTNABSim5pRQA4UuaSkzQAu/FKJM1GTTc0wLAkpweq26lD0WC5cD08NVRXqVWzUjJwadUStUgNIY6kNFIaACjbQKdQAzbk4A5NRahL5cYtkPu59T6VaLrbwtO30UeprJZi7FmOSTkmqiru5LdkQsKiZasFaaUrUgqMDUTA1cdKgdKoRWJpM09lwajK0xDs0hNABp4QmgCA0mKs+RmnLBQBWCE07yjV1YB6U/yeKLhYoCKpFUirJixTCuKAFQ1MpqrvANOEtIC1mioRLTt+aBj80E8UwGlzQAE0UxjRTEdFRRRXKdAUUUUAJRRQaBCUUUUwEpCaWkNAhKSlNJQAUlLTaYBSGloNMQ00lLSUAABJAHJNaEUYjjC/magtYv+Wh/CrNRJ9CkgpKWkqShKKWkoASilpMUxBRiiigBKKWkoASiiigBKKWimAlJS0UAJRRRQAlFLSUCCiiigAoNFFABSUtJQAUUUUAFFFFABSUtFMAooopAJS0UUwEopaSgAooooAKWkpaACijpRSAKKKKYBRRRQAUUUYoAKSlooASilpM0ALRSUUALRmkpaACkoooAKKWkoAKKWigApKWkoAKWkpaACkoooABS0UUgCjFFFMAozRRSAWikpc0AFIaWkoAKKKKYBRRRSAWikooAKWiigAooooAKKKKACiiigAooooAKWkooGLRRRQAUUUUALRSUUAFFFFIAoopaYBRRRSAKKKDQAlFApaAEopaKYBSUtJQAtFJS0AFFJS0AFFFFAB0qF23H2FPdsDHeo6EJiUUUUxBRRmkoAWnCmilFADqdTQacKQwpaKKACikzRmgAJqNjTyajY0xETGoiealaompiDNJmkzSZpgOzRmm5oJoAUmm5pCaQmgQ7NITTc0ZoGOBpwNMBpwNAh+aQmkzTWNAxCaTNIxpuaYh+aTdTc0ZpASK3NTI1Vl61YQUDLCHNSjpUSCpR0qRi0ZopKQxc0+NfMcKKjpbqX7La7Qf3sv6Ci3QCrfXImn2If3cfA9/eqwNNApwrZK2hm3cWginAUuKYiIrUTR5qyRzSFaLgUWhzTPJ9q0NgpPLFO4rFIQ+1SLFVnyhTgmKVx2K/lUojxVjaKaaLgIEFLtGKUGkZgBQBGwqCQCpJJOOtVZZapCZBKcGojLimzSVVaX3qiS8s/vU6TAiskSmp4pc0AayNuqXbVKGSravx1qSgZaKdmii4WN6iiiuY3CiiigBKKKDQISilpKYCUlLSUABpKWkNAhKQ0tJTASkpaKBCUscZkcKOnekq7BF5ac/ePWhuw0rjwAoAHAFHalorMsSiiigBMUlOpKAEopaSmAlFLRigBKKKKBCUmKWigBKKWkoASilNJTAKSlooASiiigApKWigBKKWigQlFFFABRRRQAUUUUAJRS0UAJRS0nWgAoxRS0AJQaWkoASjFLRTASloopAFBoooAKKKKACijFFMAopKWgBKKWigAopKWgBBS0UUAFFJQKAFpKWigBKKKKAClopKAFooooAKKKKACiiikAUUUUAFFFFAC0UlLQAlLRSd6YC0lLSd6ACiiloASlpKWkAlLRRQAUUUUAFFFFABRRSUALRRRQAUUUUAFFJmloAKWkooAWikpaBhRRRQAUUUUAFFFFAC0UlFAC0lLRQAlFFFABRRRQAUUUUAFFFFABRRRQAtITtGaWo2OTx0oEMPJyaKWimIbQaWm0AFFFJQAuadmmZpc0APBp2aj3Uu6gCTNJmmFqTdQA/NGaZmjNAC5phNBOKaTTAaxqM1IaZimIjI5pCKkxSFaAI6TNSFaYRQIaTTc0pFMIpgLuozTKM0ASg04GolNSCgB1NJpaQikBGxpKUqaTaaYBQAaeEqRUoAai81ZRaREqZVxSGPUVIBgU1afUjGkUmKdQASQByTSGOiVRmSThEGTWbcSm4naRu/Qegq5fShALZD05c+pqiauK6kyfQbwKQOM01+KiLYNaEFkOKdvqqHpTKB3osBZDU4YNUvP96mjmHrSsBZ20YpiyA08HNAxKQ0ppKAENNNPppoENqKU4FTYqGVcimBnzzFTVJ7jNW7mLOazXQoTVoga8hPWoi2aVmqPNMRIDU0dVQTU8bUAXkYgVYjl5qojcVPGpJpFF1HyKKSNOKKQzpKO1FFcxuFFFFABSUtJQIKSlptABSEUvWkpgJSUtJQIKSiimAlFBoVS7hV6mgRNbx723HoP51bpEQIgUdqWs27lpCUUtJQMSilpKACkpaSgApKWigBKKKKBBSUtFMBKSlpKACkpaKAEooooASilopgNopaTFABRRRQAUUUUAFJS0UAJRS0lABRRRQIKKKKACiiigAxRRRQAUUUUAJS0UUAFFFFACUtFFABSUtFABSUtJQAYpKWimAgpaKKACiilpAJRQaSgBaKKKAEopaSmAYopaSgAopaSgApaKKACiijNABSUtFIAFFFFABRRRQAtFJS0wCiiigApKWigBKKWkoAKWkpaAEopaSgBaKSlpAJRS0UAJRRRQAClpKKAFooooAKKKKACiiigApaSigBaKSloAKKKKBhRRRQAUUUYoAKWkooAWikooAWiijFACUUUUAFFFFABRRSEgDNACO2OBUeaCcmkzTJFopKM0ABpM0E00mgBTSE0lJmmAuaM0maM0ALmlLGmE0ZoEOLUm6mk00mmBLupQaiBp4NIY4000+jFAiPFGKdiigBu2kK1JSGmBERUbCpWqNjQBEaYaexqEmmIWkpu6jNMCValVagRqsRmkA4LRtp6jNKVpDItlJsqQ8UUANC09RihRT8UAOQVKKiU4p4akwJBgUucUzdTS1IY/fUgkFvbtO3U/Kg9TVeNTLKEBx6n0FUb2+We42xn91H8q+/vTUbuwXsOYlmLE5J5NNLVEJQe9I0la2Mwd6rPJSyPVV2NVYVybzsVE8/vUDMaiZiepoET+ec8VYin96oA08MRTA1o5verUcmaxEmINXYZs45qWhpmmGBoqssme9TK+akofSGlBzSGgBKRlyKWigRTmizVGa2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';

        /*
         * Jemne obrazky do pozadia troch tlacidiel osobneho stavu. Su to male
         * ciernobiele SVG vlozene priamo v skripte (rovnako ako POZADIE vyssie),
         * takze nezavisia od siete ani od ziadneho externeho suboru.
         */
        const OBR_STRETNUTIA = 'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyNjAgMjQwIiB3aWR0aD0iMjYwIiBoZWlnaHQ9IjI0MCI+CiAgPGRlZnM+CiAgICA8bGluZWFyR3JhZGllbnQgaWQ9ImdUZWxvIiB4MT0iMCIgeTE9IjAiIHgyPSIxIiB5Mj0iMSI+CiAgICAgIDxzdG9wIG9mZnNldD0iMCIgc3RvcC1jb2xvcj0iI2ZmZiIgc3RvcC1vcGFjaXR5PSIuOTUiLz4KICAgICAgPHN0b3Agb2Zmc2V0PSIxIiBzdG9wLWNvbG9yPSIjMDAwIiBzdG9wLW9wYWNpdHk9Ii4zMCIvPgogICAgPC9saW5lYXJHcmFkaWVudD4KICAgIDxsaW5lYXJHcmFkaWVudCBpZD0iZ1RlbG9UbWF2eSIgeDE9IjAiIHkxPSIwIiB4Mj0iMSIgeTI9IjEiPgogICAgICA8c3RvcCBvZmZzZXQ9IjAiIHN0b3AtY29sb3I9IiNmZmYiIHN0b3Atb3BhY2l0eT0iLjU1Ii8+CiAgICAgIDxzdG9wIG9mZnNldD0iMSIgc3RvcC1jb2xvcj0iIzAwMCIgc3RvcC1vcGFjaXR5PSIuNDUiLz4KICAgIDwvbGluZWFyR3JhZGllbnQ+CiAgICA8bGluZWFyR3JhZGllbnQgaWQ9ImdEb3NrYSIgeDE9IjAiIHkxPSIwIiB4Mj0iMCIgeTI9IjEiPgogICAgICA8c3RvcCBvZmZzZXQ9IjAiIHN0b3AtY29sb3I9IiNmZmYiIHN0b3Atb3BhY2l0eT0iLjk1Ii8+CiAgICAgIDxzdG9wIG9mZnNldD0iMSIgc3RvcC1jb2xvcj0iI2ZmZiIgc3RvcC1vcGFjaXR5PSIuNjAiLz4KICAgIDwvbGluZWFyR3JhZGllbnQ+CiAgICA8bGluZWFyR3JhZGllbnQgaWQ9ImdUYWJ1bGEiIHgxPSIwIiB5MT0iMCIgeDI9IjEiIHkyPSIxIj4KICAgICAgPHN0b3Agb2Zmc2V0PSIwIiBzdG9wLWNvbG9yPSIjZmZmIiBzdG9wLW9wYWNpdHk9Ii4zNCIvPgogICAgICA8c3RvcCBvZmZzZXQ9IjEiIHN0b3AtY29sb3I9IiNmZmYiIHN0b3Atb3BhY2l0eT0iLjA4Ii8+CiAgICA8L2xpbmVhckdyYWRpZW50PgogICAgPHJhZGlhbEdyYWRpZW50IGlkPSJnVGllbiIgY3g9IjUwJSIgY3k9IjUwJSIgcj0iNTAlIj4KICAgICAgPHN0b3Agb2Zmc2V0PSIwIiBzdG9wLWNvbG9yPSIjMDAwIiBzdG9wLW9wYWNpdHk9Ii41MCIvPgogICAgICA8c3RvcCBvZmZzZXQ9IjEiIHN0b3AtY29sb3I9IiMwMDAiIHN0b3Atb3BhY2l0eT0iMCIvPgogICAgPC9yYWRpYWxHcmFkaWVudD4KICA8L2RlZnM+CgogIDwhLS0gdGFidWxhIHMgZ3JhZm9tIC0tPgogIDxyZWN0IHg9IjE1NiIgeT0iMTQiIHdpZHRoPSI5NCIgaGVpZ2h0PSI2NiIgcng9IjYiIGZpbGw9InVybCgjZ1RhYnVsYSkiLz4KICA8cmVjdCB4PSIxNTYiIHk9IjE0IiB3aWR0aD0iOTQiIGhlaWdodD0iNjYiIHJ4PSI2IiBmaWxsPSJub25lIiBzdHJva2U9IiNmZmYiIHN0cm9rZS1vcGFjaXR5PSIuNSIgc3Ryb2tlLXdpZHRoPSIzIi8+CiAgPGcgZmlsbD0iI2ZmZiIgZmlsbC1vcGFjaXR5PSIuNzIiPgogICAgPHJlY3QgeD0iMTY3IiB5PSI1MiIgd2lkdGg9IjExIiBoZWlnaHQ9IjE4IiByeD0iMiIvPgogICAgPHJlY3QgeD0iMTg0IiB5PSI0MSIgd2lkdGg9IjExIiBoZWlnaHQ9IjI5IiByeD0iMiIvPgogICAgPHJlY3QgeD0iMjAxIiB5PSIzMiIgd2lkdGg9IjExIiBoZWlnaHQ9IjM4IiByeD0iMiIvPgogICAgPHJlY3QgeD0iMjE4IiB5PSI0NSIgd2lkdGg9IjExIiBoZWlnaHQ9IjI1IiByeD0iMiIvPgogIDwvZz4KCiAgPCEtLSB6YWRueSByYWQgcG9zdGF2IC0tPgogIDxnPgogICAgPGNpcmNsZSBjeD0iNDYiIGN5PSI3NCIgcj0iMTYiIGZpbGw9InVybCgjZ1RlbG8pIi8+CiAgICA8cGF0aCBkPSJNMjIgMTIwYzMtMTUgMTItMjIgMjQtMjJzMjEgNyAyNCAyMnoiIGZpbGw9InVybCgjZ1RlbG8pIi8+CiAgPC9nPgogIDxnPgogICAgPGNpcmNsZSBjeD0iMTA4IiBjeT0iNjIiIHI9IjE4IiBmaWxsPSJ1cmwoI2dUZWxvKSIvPgogICAgPHBhdGggZD0iTTgxIDExNGMzLTE3IDEzLTI1IDI3LTI1czI0IDggMjcgMjV6IiBmaWxsPSJ1cmwoI2dUZWxvKSIvPgogIDwvZz4KICA8Zz4KICAgIDxjaXJjbGUgY3g9IjE3MCIgY3k9Ijc4IiByPSIxNSIgZmlsbD0idXJsKCNnVGVsbykiLz4KICAgIDxwYXRoIGQ9Ik0xNDggMTIxYzMtMTQgMTEtMjEgMjItMjFzMTkgNyAyMiAyMXoiIGZpbGw9InVybCgjZ1RlbG8pIi8+CiAgPC9nPgoKICA8IS0tIHRpZW4gcG9kIHN0b2xvbSAtLT4KICA8ZWxsaXBzZSBjeD0iMTIwIiBjeT0iMTc4IiByeD0iMTAyIiByeT0iMjIiIGZpbGw9InVybCgjZ1RpZW4pIi8+CgogIDwhLS0gc3RvbDogaHJhbmEgKyBkb3NrYSAtLT4KICA8ZWxsaXBzZSBjeD0iMTIwIiBjeT0iMTU2IiByeD0iOTQiIHJ5PSIzMCIgZmlsbD0iIzAwMCIgZmlsbC1vcGFjaXR5PSIuMzgiLz4KICA8ZWxsaXBzZSBjeD0iMTIwIiBjeT0iMTUwIiByeD0iOTQiIHJ5PSIzMCIgZmlsbD0idXJsKCNnRG9za2EpIi8+CiAgPGVsbGlwc2UgY3g9IjEyMCIgY3k9IjE1MCIgcng9IjY2IiByeT0iMTciIGZpbGw9IiMwMDAiIGZpbGwtb3BhY2l0eT0iLjA5Ii8+CgogIDwhLS0gbm90ZWJvb2sgLS0+CiAgPHBhdGggZD0iTTk2IDE0MGg0NGwxMiAxNGgtNjh6IiBmaWxsPSIjZmZmIiBmaWxsLW9wYWNpdHk9Ii45NSIvPgogIDxwYXRoIGQ9Ik0xMDQgMTE4aDMwdjIyaC0zMHoiIGZpbGw9IiNmZmYiIGZpbGwtb3BhY2l0eT0iLjQ1Ii8+CiAgPHBhdGggZD0iTTEwNCAxMThoMzB2MjJoLTMweiIgZmlsbD0ibm9uZSIgc3Ryb2tlPSIjZmZmIiBzdHJva2Utb3BhY2l0eT0iLjg1IiBzdHJva2Utd2lkdGg9IjMiLz4KCiAgPCEtLSBzYWxreSBuYSBzdG9sZSAtLT4KICA8ZyBmaWxsPSIjZmZmIiBmaWxsLW9wYWNpdHk9Ii44Ij4KICAgIDxyZWN0IHg9IjU4IiB5PSIxNDAiIHdpZHRoPSIxMyIgaGVpZ2h0PSIxNCIgcng9IjMiLz4KICAgIDxyZWN0IHg9IjE4MiIgeT0iMTQyIiB3aWR0aD0iMTMiIGhlaWdodD0iMTQiIHJ4PSIzIi8+CiAgPC9nPgoKICA8IS0tIHByZWRueSByYWQgKGNocmJ0b20sIG9yZXphbnkgc3BvZGtvbSkgLS0+CiAgPGc+CiAgICA8Y2lyY2xlIGN4PSIzNiIgY3k9IjIwNCIgcj0iMjQiIGZpbGw9InVybCgjZ1RlbG9UbWF2eSkiLz4KICAgIDxwYXRoIGQ9Ik0tNiAyNjhjNC0yNyAxOS00MCA0Mi00MHMzOCAxMyA0MiA0MHoiIGZpbGw9InVybCgjZ1RlbG9UbWF2eSkiLz4KICA8L2c+CiAgPGc+CiAgICA8Y2lyY2xlIGN4PSIyMDYiIGN5PSIyMDAiIHI9IjIyIiBmaWxsPSJ1cmwoI2dUZWxvVG1hdnkpIi8+CiAgICA8cGF0aCBkPSJNMTY4IDI2MmM0LTI1IDE3LTM3IDM4LTM3czM0IDEyIDM4IDM3eiIgZmlsbD0idXJsKCNnVGVsb1RtYXZ5KSIvPgogIDwvZz4KPC9zdmc+Cg==';
        const OBR_PRESTAVKA = 'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyNjAgMjQwIiB3aWR0aD0iMjYwIiBoZWlnaHQ9IjI0MCI+CiAgPGRlZnM+CiAgICA8bGluZWFyR3JhZGllbnQgaWQ9ImNUZWxvIiB4MT0iMCIgeTE9IjAiIHgyPSIxIiB5Mj0iMCI+CiAgICAgIDxzdG9wIG9mZnNldD0iMCIgICBzdG9wLWNvbG9yPSIjZmZmIiBzdG9wLW9wYWNpdHk9Ii41NSIvPgogICAgICA8c3RvcCBvZmZzZXQ9Ii4yMiIgc3RvcC1jb2xvcj0iI2ZmZiIgc3RvcC1vcGFjaXR5PSIuOTgiLz4KICAgICAgPHN0b3Agb2Zmc2V0PSIuNjIiIHN0b3AtY29sb3I9IiNmZmYiIHN0b3Atb3BhY2l0eT0iLjgwIi8+CiAgICAgIDxzdG9wIG9mZnNldD0iMSIgICBzdG9wLWNvbG9yPSIjMDAwIiBzdG9wLW9wYWNpdHk9Ii4zOCIvPgogICAgPC9saW5lYXJHcmFkaWVudD4KICAgIDxsaW5lYXJHcmFkaWVudCBpZD0iY0thdmEiIHgxPSIwIiB5MT0iMCIgeDI9IjEiIHkyPSIxIj4KICAgICAgPHN0b3Agb2Zmc2V0PSIwIiBzdG9wLWNvbG9yPSIjMDAwIiBzdG9wLW9wYWNpdHk9Ii41MiIvPgogICAgICA8c3RvcCBvZmZzZXQ9IjEiIHN0b3AtY29sb3I9IiMwMDAiIHN0b3Atb3BhY2l0eT0iLjIyIi8+CiAgICA8L2xpbmVhckdyYWRpZW50PgogICAgPGxpbmVhckdyYWRpZW50IGlkPSJjVGFuaWVyIiB4MT0iMCIgeTE9IjAiIHgyPSIxIiB5Mj0iMCI+CiAgICAgIDxzdG9wIG9mZnNldD0iMCIgICBzdG9wLWNvbG9yPSIjZmZmIiBzdG9wLW9wYWNpdHk9Ii42MiIvPgogICAgICA8c3RvcCBvZmZzZXQ9Ii4zNSIgc3RvcC1jb2xvcj0iI2ZmZiIgc3RvcC1vcGFjaXR5PSIuOTUiLz4KICAgICAgPHN0b3Agb2Zmc2V0PSIxIiAgIHN0b3AtY29sb3I9IiMwMDAiIHN0b3Atb3BhY2l0eT0iLjMwIi8+CiAgICA8L2xpbmVhckdyYWRpZW50PgogICAgPGxpbmVhckdyYWRpZW50IGlkPSJjVWNobyIgeDE9IjAiIHkxPSIwIiB4Mj0iMSIgeTI9IjEiPgogICAgICA8c3RvcCBvZmZzZXQ9IjAiIHN0b3AtY29sb3I9IiNmZmYiIHN0b3Atb3BhY2l0eT0iLjkiLz4KICAgICAgPHN0b3Agb2Zmc2V0PSIxIiBzdG9wLWNvbG9yPSIjMDAwIiBzdG9wLW9wYWNpdHk9Ii4zMCIvPgogICAgPC9saW5lYXJHcmFkaWVudD4KICAgIDxyYWRpYWxHcmFkaWVudCBpZD0iY1RpZW4iIGN4PSI1MCUiIGN5PSI1MCUiIHI9IjUwJSI+CiAgICAgIDxzdG9wIG9mZnNldD0iMCIgc3RvcC1jb2xvcj0iIzAwMCIgc3RvcC1vcGFjaXR5PSIuNDUiLz4KICAgICAgPHN0b3Agb2Zmc2V0PSIxIiBzdG9wLWNvbG9yPSIjMDAwIiBzdG9wLW9wYWNpdHk9IjAiLz4KICAgIDwvcmFkaWFsR3JhZGllbnQ+CiAgPC9kZWZzPgoKICA8IS0tIHBhcmEgLS0+CiAgPGcgZmlsbD0ibm9uZSIgc3Ryb2tlPSIjZmZmIiBzdHJva2UtbGluZWNhcD0icm91bmQiIHN0cm9rZS13aWR0aD0iNyI+CiAgICA8cGF0aCBkPSJNOTIgNjJjMTItMTEtMTItMjIgMC0zMyIgc3Ryb2tlLW9wYWNpdHk9Ii4zOCIvPgogICAgPHBhdGggZD0iTTEyNiA1NmMxMy0xMi0xMy0yNCAwLTM2IiBzdHJva2Utb3BhY2l0eT0iLjU1Ii8+CiAgICA8cGF0aCBkPSJNMTYwIDYyYzEyLTExLTEyLTIyIDAtMzMiIHN0cm9rZS1vcGFjaXR5PSIuMzgiLz4KICA8L2c+CgogIDwhLS0gdGllbiAtLT4KICA8ZWxsaXBzZSBjeD0iMTI2IiBjeT0iMjA2IiByeD0iOTYiIHJ5PSIxOCIgZmlsbD0idXJsKCNjVGllbikiLz4KCiAgPCEtLSB1Y2hvIC0tPgogIDxwYXRoIGQ9Ik0xODggMTEyaDE0YzIwIDAgMzYgMTUgMzYgMzNzLTE2IDMzLTM2IDMzaC0xNCIgZmlsbD0ibm9uZSIgc3Ryb2tlPSJ1cmwoI2NVY2hvKSIgc3Ryb2tlLXdpZHRoPSIxNyIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIi8+CgogIDwhLS0gdGVsbyBzYWxreSAtLT4KICA8cGF0aCBkPSJNNDQgOTZoMTQ2djUyYzAgMzQtMjQgNjAtNTYgNjBIMTAwYy0zMiAwLTU2LTI2LTU2LTYweiIgZmlsbD0idXJsKCNjVGVsbykiLz4KICA8IS0tIGhyZGxvIGEga2F2YSAtLT4KICA8ZWxsaXBzZSBjeD0iMTE3IiBjeT0iOTYiIHJ4PSI3MyIgcnk9IjE3IiBmaWxsPSIjZmZmIiBmaWxsLW9wYWNpdHk9Ii45NSIvPgogIDxlbGxpcHNlIGN4PSIxMTciIGN5PSI5NyIgcng9IjYxIiByeT0iMTIiIGZpbGw9InVybCgjY0thdmEpIi8+CiAgPGVsbGlwc2UgY3g9Ijk2IiBjeT0iOTQiIHJ4PSIyMCIgcnk9IjUiIGZpbGw9IiNmZmYiIGZpbGwtb3BhY2l0eT0iLjIyIi8+CiAgPCEtLSBvZGxlc2sgbmEgdGVsZSAtLT4KICA8cGF0aCBkPSJNNjIgMTEyYzAgMzQgNCA1OCAxNCA3Ni0xNC0xMi0yMi0zNC0yMi02MHoiIGZpbGw9IiNmZmYiIGZpbGwtb3BhY2l0eT0iLjQ1Ii8+CgogIDwhLS0gcG9kbm9zIC0tPgogIDxlbGxpcHNlIGN4PSIxMjYiIGN5PSIyMDAiIHJ4PSIxMDQiIHJ5PSIyMCIgZmlsbD0idXJsKCNjVGFuaWVyKSIvPgogIDxlbGxpcHNlIGN4PSIxMjYiIGN5PSIxOTYiIHJ4PSI3NiIgcnk9IjEzIiBmaWxsPSIjMDAwIiBmaWxsLW9wYWNpdHk9Ii4xNCIvPgoKICA8IS0tIHN1c2llbmthIC0tPgogIDxnPgogICAgPGNpcmNsZSBjeD0iMjE0IiBjeT0iMTkyIiByPSIxOSIgZmlsbD0iI2ZmZiIgZmlsbC1vcGFjaXR5PSIuODUiLz4KICAgIDxjaXJjbGUgY3g9IjIxNCIgY3k9IjE5MiIgcj0iMTkiIGZpbGw9IiMwMDAiIGZpbGwtb3BhY2l0eT0iLjEyIi8+CiAgICA8ZyBmaWxsPSIjMDAwIiBmaWxsLW9wYWNpdHk9Ii4zNSI+CiAgICAgIDxjaXJjbGUgY3g9IjIwOCIgY3k9IjE4NiIgcj0iMy40Ii8+PGNpcmNsZSBjeD0iMjIwIiBjeT0iMTkwIiByPSIzIi8+CiAgICAgIDxjaXJjbGUgY3g9IjIxMSIgY3k9IjE5OSIgcj0iMy4yIi8+PGNpcmNsZSBjeD0iMjIxIiBjeT0iMjAwIiByPSIyLjYiLz4KICAgIDwvZz4KICA8L2c+Cjwvc3ZnPgo=';
        const OBR_CAKANIE = 'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyNjAgMjQwIiB3aWR0aD0iMjYwIiBoZWlnaHQ9IjI0MCI+CiAgPGRlZnM+CiAgICA8bGluZWFyR3JhZGllbnQgaWQ9ImhSYW0iIHgxPSIwIiB5MT0iMCIgeDI9IjEiIHkyPSIwIj4KICAgICAgPHN0b3Agb2Zmc2V0PSIwIiAgIHN0b3AtY29sb3I9IiNmZmYiIHN0b3Atb3BhY2l0eT0iLjU1Ii8+CiAgICAgIDxzdG9wIG9mZnNldD0iLjMiICBzdG9wLWNvbG9yPSIjZmZmIiBzdG9wLW9wYWNpdHk9Ii45OCIvPgogICAgICA8c3RvcCBvZmZzZXQ9IjEiICAgc3RvcC1jb2xvcj0iIzAwMCIgc3RvcC1vcGFjaXR5PSIuMzUiLz4KICAgIDwvbGluZWFyR3JhZGllbnQ+CiAgICA8bGluZWFyR3JhZGllbnQgaWQ9ImhTa2xvIiB4MT0iMCIgeTE9IjAiIHgyPSIxIiB5Mj0iMCI+CiAgICAgIDxzdG9wIG9mZnNldD0iMCIgICBzdG9wLWNvbG9yPSIjZmZmIiBzdG9wLW9wYWNpdHk9Ii40MCIvPgogICAgICA8c3RvcCBvZmZzZXQ9Ii4yOCIgc3RvcC1jb2xvcj0iI2ZmZiIgc3RvcC1vcGFjaXR5PSIuMTIiLz4KICAgICAgPHN0b3Agb2Zmc2V0PSIuNzIiIHN0b3AtY29sb3I9IiNmZmYiIHN0b3Atb3BhY2l0eT0iLjA2Ii8+CiAgICAgIDxzdG9wIG9mZnNldD0iMSIgICBzdG9wLWNvbG9yPSIjMDAwIiBzdG9wLW9wYWNpdHk9Ii4yNiIvPgogICAgPC9saW5lYXJHcmFkaWVudD4KICAgIDxsaW5lYXJHcmFkaWVudCBpZD0iaFBpZXNvayIgeDE9IjAiIHkxPSIwIiB4Mj0iMCIgeTI9IjEiPgogICAgICA8c3RvcCBvZmZzZXQ9IjAiIHN0b3AtY29sb3I9IiNmZmYiIHN0b3Atb3BhY2l0eT0iLjk1Ii8+CiAgICAgIDxzdG9wIG9mZnNldD0iMSIgc3RvcC1jb2xvcj0iIzAwMCIgc3RvcC1vcGFjaXR5PSIuMzAiLz4KICAgIDwvbGluZWFyR3JhZGllbnQ+CiAgICA8cmFkaWFsR3JhZGllbnQgaWQ9ImhUaWVuIiBjeD0iNTAlIiBjeT0iNTAlIiByPSI1MCUiPgogICAgICA8c3RvcCBvZmZzZXQ9IjAiIHN0b3AtY29sb3I9IiMwMDAiIHN0b3Atb3BhY2l0eT0iLjQ4Ii8+CiAgICAgIDxzdG9wIG9mZnNldD0iMSIgc3RvcC1jb2xvcj0iIzAwMCIgc3RvcC1vcGFjaXR5PSIwIi8+CiAgICA8L3JhZGlhbEdyYWRpZW50PgogIDwvZGVmcz4KCiAgPCEtLSB0aWVuIC0tPgogIDxlbGxpcHNlIGN4PSIxMzAiIGN5PSIyMjIiIHJ4PSI4NiIgcnk9IjE1IiBmaWxsPSJ1cmwoI2hUaWVuKSIvPgoKICA8IS0tIHNrbG8gLS0+CiAgPHBhdGggZD0iTTc0IDM0aDExMnYxNmMwIDM0LTQ0IDU2LTQ0IDcwczQ0IDM2IDQ0IDcwdjE2SDc0di0xNmMwLTM0IDQ0LTU2IDQ0LTcwcy00NC0zNi00NC03MHoiIGZpbGw9InVybCgjaFNrbG8pIi8+CiAgPHBhdGggZD0iTTc0IDM0aDExMnYxNmMwIDM0LTQ0IDU2LTQ0IDcwczQ0IDM2IDQ0IDcwdjE2SDc0di0xNmMwLTM0IDQ0LTU2IDQ0LTcwcy00NC0zNi00NC03MHoiCiAgICAgICAgZmlsbD0ibm9uZSIgc3Ryb2tlPSIjZmZmIiBzdHJva2Utb3BhY2l0eT0iLjY1IiBzdHJva2Utd2lkdGg9IjUiIHN0cm9rZS1saW5lam9pbj0icm91bmQiLz4KICA8IS0tIG9kbGVzayBuYSBza2xlIC0tPgogIDxwYXRoIGQ9Ik04OCA0NGMwIDI2IDI4IDQ2IDI4IDYwLTEwLTYtMzQtMjgtMzQtNDh6IiBmaWxsPSIjZmZmIiBmaWxsLW9wYWNpdHk9Ii40MCIvPgogIDxwYXRoIGQ9Ik05MiAxOTJjMC0yNCAyNC00MiAyNC01Mi05IDUtMzIgMjQtMzIgNDJ6IiBmaWxsPSIjZmZmIiBmaWxsLW9wYWNpdHk9Ii4yMiIvPgoKICA8IS0tIGhvcm55IHBpZXNvayAtLT4KICA8cGF0aCBkPSJNODYgNDhoODhjLTIgMjYtMzYgNDQtNDQgNTItOC04LTQyLTI2LTQ0LTUyeiIgZmlsbD0idXJsKCNoUGllc29rKSIvPgogIDxwYXRoIGQ9Ik04NiA0OGg4OGMtMSA3LTQgMTMtOCAxOUg5NGMtNC02LTctMTItOC0xOXoiIGZpbGw9IiNmZmYiIGZpbGwtb3BhY2l0eT0iLjM1Ii8+CgogIDwhLS0gcGFkYWp1Y2kgcHJhbWVuIGEgenJua2EgLS0+CiAgPHJlY3QgeD0iMTI2IiB5PSIxMTYiIHdpZHRoPSI4IiBoZWlnaHQ9IjU4IiByeD0iNCIgZmlsbD0iI2ZmZiIgZmlsbC1vcGFjaXR5PSIuODUiLz4KICA8ZyBmaWxsPSIjZmZmIiBmaWxsLW9wYWNpdHk9Ii43Ij4KICAgIDxjaXJjbGUgY3g9IjEyMSIgY3k9IjE0MCIgcj0iMyIvPjxjaXJjbGUgY3g9IjEzOSIgY3k9IjE1NiIgcj0iMi42Ii8+PGNpcmNsZSBjeD0iMTIwIiBjeT0iMTY0IiByPSIyLjIiLz4KICA8L2c+CgogIDwhLS0gZG9sbnkga29wY2VrIC0tPgogIDxwYXRoIGQ9Ik04OCAxOTJjNi0yNiAzNC0zOCA0Mi0zOHMzNiAxMiA0MiAzOHoiIGZpbGw9InVybCgjaFBpZXNvaykiLz4KICA8cGF0aCBkPSJNMTA0IDE5MmM1LTE0IDE5LTIyIDI2LTIyczIxIDggMjYgMjJ6IiBmaWxsPSIjZmZmIiBmaWxsLW9wYWNpdHk9Ii4zMCIvPgoKICA8IS0tIGhvcm5lIGEgZG9sbmUgZG9za3kgLS0+CiAgPHJlY3QgeD0iNTgiIHk9IjE2IiB3aWR0aD0iMTQ0IiBoZWlnaHQ9IjIwIiByeD0iOSIgZmlsbD0idXJsKCNoUmFtKSIvPgogIDxyZWN0IHg9IjU4IiB5PSIyMDQiIHdpZHRoPSIxNDQiIGhlaWdodD0iMjAiIHJ4PSI5IiBmaWxsPSJ1cmwoI2hSYW0pIi8+CiAgPHJlY3QgeD0iNTgiIHk9IjE2IiB3aWR0aD0iMTQ0IiBoZWlnaHQ9IjciIHJ4PSIzLjUiIGZpbGw9IiNmZmYiIGZpbGwtb3BhY2l0eT0iLjQ1Ii8+CiAgPHJlY3QgeD0iNTgiIHk9IjIwNCIgd2lkdGg9IjE0NCIgaGVpZ2h0PSI3IiByeD0iMy41IiBmaWxsPSIjZmZmIiBmaWxsLW9wYWNpdHk9Ii40NSIvPgo8L3N2Zz4K';
        const OBR_VYROBA = 'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyNjAgMjQwIiB3aWR0aD0iMjYwIiBoZWlnaHQ9IjI0MCI+CiAgPGRlZnM+CiAgICA8bGluZWFyR3JhZGllbnQgaWQ9InZLb3YiIHgxPSIwIiB5MT0iMCIgeDI9IjEiIHkyPSIwIj4KICAgICAgPHN0b3Agb2Zmc2V0PSIwIiBzdG9wLWNvbG9yPSIjZmZmIiBzdG9wLW9wYWNpdHk9Ii41MCIvPgogICAgICA8c3RvcCBvZmZzZXQ9Ii4yOCIgc3RvcC1jb2xvcj0iI2ZmZiIgc3RvcC1vcGFjaXR5PSIuOTgiLz4KICAgICAgPHN0b3Agb2Zmc2V0PSIuNyIgc3RvcC1jb2xvcj0iI2ZmZiIgc3RvcC1vcGFjaXR5PSIuNzAiLz4KICAgICAgPHN0b3Agb2Zmc2V0PSIxIiBzdG9wLWNvbG9yPSIjMDAwIiBzdG9wLW9wYWNpdHk9Ii40MCIvPgogICAgPC9saW5lYXJHcmFkaWVudD4KICAgIDxsaW5lYXJHcmFkaWVudCBpZD0idkRpZWwiIHgxPSIwIiB5MT0iMCIgeDI9IjAiIHkyPSIxIj4KICAgICAgPHN0b3Agb2Zmc2V0PSIwIiBzdG9wLWNvbG9yPSIjZmZmIiBzdG9wLW9wYWNpdHk9Ii45MCIvPgogICAgICA8c3RvcCBvZmZzZXQ9IjEiIHN0b3AtY29sb3I9IiMwMDAiIHN0b3Atb3BhY2l0eT0iLjM4Ii8+CiAgICA8L2xpbmVhckdyYWRpZW50PgogICAgPHJhZGlhbEdyYWRpZW50IGlkPSJ2VGllbiIgY3g9IjUwJSIgY3k9IjUwJSIgcj0iNTAlIj4KICAgICAgPHN0b3Agb2Zmc2V0PSIwIiBzdG9wLWNvbG9yPSIjMDAwIiBzdG9wLW9wYWNpdHk9Ii40OCIvPgogICAgICA8c3RvcCBvZmZzZXQ9IjEiIHN0b3AtY29sb3I9IiMwMDAiIHN0b3Atb3BhY2l0eT0iMCIvPgogICAgPC9yYWRpYWxHcmFkaWVudD4KICA8L2RlZnM+CgogIDxlbGxpcHNlIGN4PSIxMzAiIGN5PSIyMTQiIHJ4PSI5NiIgcnk9IjE2IiBmaWxsPSJ1cmwoI3ZUaWVuKSIvPgoKICA8IS0tIHZyZXRlbm8gLS0+CiAgPHJlY3QgeD0iOTYiIHk9IjEwIiB3aWR0aD0iNjgiIGhlaWdodD0iMzQiIHJ4PSI4IiBmaWxsPSJ1cmwoI3ZLb3YpIi8+CiAgPHJlY3QgeD0iMTA2IiB5PSI0NCIgd2lkdGg9IjQ4IiBoZWlnaHQ9IjIyIiByeD0iNSIgZmlsbD0idXJsKCN2S292KSIvPgogIDxyZWN0IHg9IjEwMCIgeT0iMTQiIHdpZHRoPSIxMCIgaGVpZ2h0PSIyNiIgcng9IjUiIGZpbGw9IiNmZmYiIGZpbGwtb3BhY2l0eT0iLjUiLz4KICA8IS0tIGtsaWXFoXRpbmEgLS0+CiAgPHBhdGggZD0iTTExMiA2NmgzNmwtNiAxOGgtMjR6IiBmaWxsPSJ1cmwoI3ZLb3YpIi8+CiAgPCEtLSBmcsOpemEgc28genViYW1pIC0tPgogIDxyZWN0IHg9IjExOCIgeT0iODQiIHdpZHRoPSIyNCIgaGVpZ2h0PSI1MiIgcng9IjQiIGZpbGw9InVybCgjdktvdikiLz4KICA8ZyBmaWxsPSIjMDAwIiBmaWxsLW9wYWNpdHk9Ii4zMCI+CiAgICA8cGF0aCBkPSJNMTE4IDkyaDI0bC0yNCAxMnoiLz48cGF0aCBkPSJNMTE4IDEwOGgyNGwtMjQgMTJ6Ii8+PHBhdGggZD0iTTExOCAxMjRoMjRsLTI0IDEweiIvPgogIDwvZz4KICA8cGF0aCBkPSJNMTIxIDg2aDV2NDhoLTV6IiBmaWxsPSIjZmZmIiBmaWxsLW9wYWNpdHk9Ii41NSIvPgoKICA8IS0tIG9icm9ib2sgLS0+CiAgPHBhdGggZD0iTTQwIDE1MGgxODB2MjJINDB6IiBmaWxsPSJ1cmwoI3ZEaWVsKSIvPgogIDxwYXRoIGQ9Ik00MCAxNTBsMTgtMTZoMTgwbC0xOCAxNnoiIGZpbGw9IiNmZmYiIGZpbGwtb3BhY2l0eT0iLjc1Ii8+CiAgPHBhdGggZD0iTTIyMCAxNTBsMTgtMTZ2MjJsLTE4IDE2eiIgZmlsbD0iIzAwMCIgZmlsbC1vcGFjaXR5PSIuMjgiLz4KICA8IS0tIGRyYXprYSBwbyBmcmV6ZSAtLT4KICA8cGF0aCBkPSJNMTEyIDEzNGgzNnYxNmgtMzZ6IiBmaWxsPSIjMDAwIiBmaWxsLW9wYWNpdHk9Ii4zNSIvPgoKICA8IS0tIHRyaWVza3kgLS0+CiAgPGcgZmlsbD0ibm9uZSIgc3Ryb2tlPSIjZmZmIiBzdHJva2Utd2lkdGg9IjUiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIgc3Ryb2tlLW9wYWNpdHk9Ii43NSI+CiAgICA8cGF0aCBkPSJNMTYwIDEyNmMxMC00IDE2LTEyIDE0LTIyIi8+CiAgICA8cGF0aCBkPSJNMTc2IDE0MGMxNC0yIDIyLTEyIDIyLTI0Ii8+CiAgICA8cGF0aCBkPSJNMTAwIDEyNGMtMTAtNC0xNi0xMi0xNC0yMiIvPgogIDwvZz4KICA8ZyBmaWxsPSIjZmZmIiBmaWxsLW9wYWNpdHk9Ii42Ij4KICAgIDxjaXJjbGUgY3g9IjE5NiIgY3k9IjEwOCIgcj0iNCIvPjxjaXJjbGUgY3g9Ijc0IiBjeT0iMTA0IiByPSIzLjQiLz48Y2lyY2xlIGN4PSIyMTAiIGN5PSIxMjYiIHI9IjMiLz4KICA8L2c+Cjwvc3ZnPgo=';
        const OBR_UPINANIE = 'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyNjAgMjQwIiB3aWR0aD0iMjYwIiBoZWlnaHQ9IjI0MCI+CiAgPGRlZnM+CiAgICA8bGluZWFyR3JhZGllbnQgaWQ9InVLb3YiIHgxPSIwIiB5MT0iMCIgeDI9IjAiIHkyPSIxIj4KICAgICAgPHN0b3Agb2Zmc2V0PSIwIiBzdG9wLWNvbG9yPSIjZmZmIiBzdG9wLW9wYWNpdHk9Ii45NSIvPgogICAgICA8c3RvcCBvZmZzZXQ9IjEiIHN0b3AtY29sb3I9IiMwMDAiIHN0b3Atb3BhY2l0eT0iLjQwIi8+CiAgICA8L2xpbmVhckdyYWRpZW50PgogICAgPGxpbmVhckdyYWRpZW50IGlkPSJ1Q2VsbyIgeDE9IjAiIHkxPSIwIiB4Mj0iMSIgeTI9IjAiPgogICAgICA8c3RvcCBvZmZzZXQ9IjAiIHN0b3AtY29sb3I9IiNmZmYiIHN0b3Atb3BhY2l0eT0iLjk1Ii8+CiAgICAgIDxzdG9wIG9mZnNldD0iMSIgc3RvcC1jb2xvcj0iIzAwMCIgc3RvcC1vcGFjaXR5PSIuMzIiLz4KICAgIDwvbGluZWFyR3JhZGllbnQ+CiAgICA8bGluZWFyR3JhZGllbnQgaWQ9InVDZWxvUCIgeDE9IjAiIHkxPSIwIiB4Mj0iMSIgeTI9IjAiPgogICAgICA8c3RvcCBvZmZzZXQ9IjAiIHN0b3AtY29sb3I9IiMwMDAiIHN0b3Atb3BhY2l0eT0iLjI4Ii8+CiAgICAgIDxzdG9wIG9mZnNldD0iMSIgc3RvcC1jb2xvcj0iI2ZmZiIgc3RvcC1vcGFjaXR5PSIuOTAiLz4KICAgIDwvbGluZWFyR3JhZGllbnQ+CiAgICA8bGluZWFyR3JhZGllbnQgaWQ9InVEaWVsIiB4MT0iMCIgeTE9IjAiIHgyPSIwIiB5Mj0iMSI+CiAgICAgIDxzdG9wIG9mZnNldD0iMCIgc3RvcC1jb2xvcj0iI2ZmZiIgc3RvcC1vcGFjaXR5PSIuOTIiLz4KICAgICAgPHN0b3Agb2Zmc2V0PSIxIiBzdG9wLWNvbG9yPSIjMDAwIiBzdG9wLW9wYWNpdHk9Ii40MCIvPgogICAgPC9saW5lYXJHcmFkaWVudD4KICAgIDxyYWRpYWxHcmFkaWVudCBpZD0idVRpZW4iIGN4PSI1MCUiIGN5PSI1MCUiIHI9IjUwJSI+CiAgICAgIDxzdG9wIG9mZnNldD0iMCIgc3RvcC1jb2xvcj0iIzAwMCIgc3RvcC1vcGFjaXR5PSIuNDgiLz4KICAgICAgPHN0b3Agb2Zmc2V0PSIxIiBzdG9wLWNvbG9yPSIjMDAwIiBzdG9wLW9wYWNpdHk9IjAiLz4KICAgIDwvcmFkaWFsR3JhZGllbnQ+CiAgPC9kZWZzPgoKICA8ZWxsaXBzZSBjeD0iMTMwIiBjeT0iMjEyIiByeD0iMTAyIiByeT0iMTYiIGZpbGw9InVybCgjdVRpZW4pIi8+CgogIDwhLS0gc2lwa3kgcHJpdGxha3UgLS0+CiAgPGcgZmlsbD0iI2ZmZiIgZmlsbC1vcGFjaXR5PSIuODUiPgogICAgPHBhdGggZD0iTTQwIDYyaDI2di05bDIwIDE1LTIwIDE1di05SDQweiIvPgogICAgPHBhdGggZD0iTTIyMCA2MmgtMjZ2LTlsLTIwIDE1IDIwIDE1di05aDI2eiIvPgogIDwvZz4KCiAgPCEtLSB1cGluYW55IGRpZWwgbWVkemkgY2VsdXN0YW1pIC0tPgogIDxyZWN0IHg9Ijk2IiB5PSI3OCIgd2lkdGg9IjY4IiBoZWlnaHQ9Ijc0IiByeD0iMyIgZmlsbD0idXJsKCN1RGllbCkiLz4KICA8cGF0aCBkPSJNOTYgNzhsMTItMTFoNjhsLTEyIDExeiIgZmlsbD0iI2ZmZiIgZmlsbC1vcGFjaXR5PSIuODIiLz4KICA8cGF0aCBkPSJNMTY0IDc4bDEyLTExdjc0bC0xMiAxMXoiIGZpbGw9IiMwMDAiIGZpbGwtb3BhY2l0eT0iLjI2Ii8+CgogIDwhLS0gemFrbGFkbmEgenZlcmFrYSAtLT4KICA8cmVjdCB4PSIyNCIgeT0iMTUyIiB3aWR0aD0iMjEyIiBoZWlnaHQ9IjMyIiByeD0iNiIgZmlsbD0idXJsKCN1S292KSIvPgogIDxyZWN0IHg9IjI0IiB5PSIxNTIiIHdpZHRoPSIyMTIiIGhlaWdodD0iOSIgcng9IjQiIGZpbGw9IiNmZmYiIGZpbGwtb3BhY2l0eT0iLjQ1Ii8+CiAgPHJlY3QgeD0iMzgiIHk9IjE4NCIgd2lkdGg9IjE4NCIgaGVpZ2h0PSIxMiIgcng9IjUiIGZpbGw9IiMwMDAiIGZpbGwtb3BhY2l0eT0iLjI4Ii8+CgogIDwhLS0gcGV2bmEgY2VsdXN0IChkb3R5a2Egc2EgZGllbHUpIC0tPgogIDxyZWN0IHg9IjUwIiB5PSI5MiIgd2lkdGg9IjQ2IiBoZWlnaHQ9IjYwIiByeD0iNCIgZmlsbD0idXJsKCN1Q2VsbykiLz4KICA8cmVjdCB4PSI1NCIgeT0iOTYiIHdpZHRoPSI5IiBoZWlnaHQ9IjUyIiByeD0iNCIgZmlsbD0iI2ZmZiIgZmlsbC1vcGFjaXR5PSIuNTUiLz4KICA8cmVjdCB4PSI4OCIgeT0iOTIiIHdpZHRoPSI4IiBoZWlnaHQ9IjYwIiBmaWxsPSIjMDAwIiBmaWxsLW9wYWNpdHk9Ii4yMiIvPgogIDwhLS0gcG9oeWJsaXZhIGNlbHVzdCAoZG90eWthIHNhIGRpZWx1KSAtLT4KICA8cmVjdCB4PSIxNjQiIHk9IjkyIiB3aWR0aD0iNDYiIGhlaWdodD0iNjAiIHJ4PSI0IiBmaWxsPSJ1cmwoI3VDZWxvUCkiLz4KICA8cmVjdCB4PSIxNjQiIHk9IjkyIiB3aWR0aD0iOCIgaGVpZ2h0PSI2MCIgZmlsbD0iIzAwMCIgZmlsbC1vcGFjaXR5PSIuMjAiLz4KICA8cmVjdCB4PSIxOTgiIHk9Ijk2IiB3aWR0aD0iOCIgaGVpZ2h0PSI1MiIgcng9IjQiIGZpbGw9IiNmZmYiIGZpbGwtb3BhY2l0eT0iLjQ1Ii8+CgogIDwhLS0gdnJldGVubyBzbyB6YXZpdG9tIGEga2x1a291IC0tPgogIDxyZWN0IHg9IjIwOCIgeT0iMTE0IiB3aWR0aD0iMzQiIGhlaWdodD0iMTYiIHJ4PSI3IiBmaWxsPSJ1cmwoI3VLb3YpIi8+CiAgPGcgc3Ryb2tlPSIjMDAwIiBzdHJva2Utb3BhY2l0eT0iLjMwIiBzdHJva2Utd2lkdGg9IjMiPgogICAgPHBhdGggZD0iTTIxNCAxMTR2MTZNMjIyIDExNHYxNk0yMzAgMTE0djE2Ii8+CiAgPC9nPgogIDxjaXJjbGUgY3g9IjI0NiIgY3k9IjEyMiIgcj0iMTIiIGZpbGw9InVybCgjdUtvdikiLz4KICA8Y2lyY2xlIGN4PSIyNDMiIGN5PSIxMTgiIHI9IjQiIGZpbGw9IiNmZmYiIGZpbGwtb3BhY2l0eT0iLjUiLz4KPC9zdmc+Cg==';
        const OBR_UPRATOVANIE = 'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyNjAgMjQwIiB3aWR0aD0iMjYwIiBoZWlnaHQ9IjI0MCI+CiAgPGRlZnM+CiAgICA8bGluZWFyR3JhZGllbnQgaWQ9InBOYXNhZGEiIHgxPSIwIiB5MT0iMCIgeDI9IjEiIHkyPSIwIj4KICAgICAgPHN0b3Agb2Zmc2V0PSIwIiBzdG9wLWNvbG9yPSIjZmZmIiBzdG9wLW9wYWNpdHk9Ii45NSIvPgogICAgICA8c3RvcCBvZmZzZXQ9IjEiIHN0b3AtY29sb3I9IiMwMDAiIHN0b3Atb3BhY2l0eT0iLjM4Ii8+CiAgICA8L2xpbmVhckdyYWRpZW50PgogICAgPGxpbmVhckdyYWRpZW50IGlkPSJwVGVsbyIgeDE9IjAiIHkxPSIwIiB4Mj0iMCIgeTI9IjEiPgogICAgICA8c3RvcCBvZmZzZXQ9IjAiIHN0b3AtY29sb3I9IiNmZmYiIHN0b3Atb3BhY2l0eT0iLjk1Ii8+CiAgICAgIDxzdG9wIG9mZnNldD0iMSIgc3RvcC1jb2xvcj0iIzAwMCIgc3RvcC1vcGFjaXR5PSIuMzIiLz4KICAgIDwvbGluZWFyR3JhZGllbnQ+CiAgICA8bGluZWFyR3JhZGllbnQgaWQ9InBTdGV0aW55IiB4MT0iMCIgeTE9IjAiIHgyPSIwIiB5Mj0iMSI+CiAgICAgIDxzdG9wIG9mZnNldD0iMCIgc3RvcC1jb2xvcj0iI2ZmZiIgc3RvcC1vcGFjaXR5PSIuODUiLz4KICAgICAgPHN0b3Agb2Zmc2V0PSIxIiBzdG9wLWNvbG9yPSIjMDAwIiBzdG9wLW9wYWNpdHk9Ii40NSIvPgogICAgPC9saW5lYXJHcmFkaWVudD4KICAgIDxsaW5lYXJHcmFkaWVudCBpZD0icFN0b2wiIHgxPSIwIiB5MT0iMCIgeDI9IjAiIHkyPSIxIj4KICAgICAgPHN0b3Agb2Zmc2V0PSIwIiBzdG9wLWNvbG9yPSIjZmZmIiBzdG9wLW9wYWNpdHk9Ii44NSIvPgogICAgICA8c3RvcCBvZmZzZXQ9IjEiIHN0b3AtY29sb3I9IiNmZmYiIHN0b3Atb3BhY2l0eT0iLjM1Ii8+CiAgICA8L2xpbmVhckdyYWRpZW50PgogICAgPHJhZGlhbEdyYWRpZW50IGlkPSJwVGllbiIgY3g9IjUwJSIgY3k9IjUwJSIgcj0iNTAlIj4KICAgICAgPHN0b3Agb2Zmc2V0PSIwIiBzdG9wLWNvbG9yPSIjMDAwIiBzdG9wLW9wYWNpdHk9Ii40NSIvPgogICAgICA8c3RvcCBvZmZzZXQ9IjEiIHN0b3AtY29sb3I9IiMwMDAiIHN0b3Atb3BhY2l0eT0iMCIvPgogICAgPC9yYWRpYWxHcmFkaWVudD4KICA8L2RlZnM+CgogIDxlbGxpcHNlIGN4PSIxMzAiIGN5PSIyMTYiIHJ4PSIxMDIiIHJ5PSIxNSIgZmlsbD0idXJsKCNwVGllbikiLz4KCiAgPCEtLSBkb3NrYSBzdG9sYSAtLT4KICA8cGF0aCBkPSJNMTYgMTc2aDIyOHYxNkgxNnoiIGZpbGw9InVybCgjcFN0b2wpIi8+CiAgPHBhdGggZD0iTTE2IDE5MmgyMjh2MTBIMTZ6IiBmaWxsPSIjMDAwIiBmaWxsLW9wYWNpdHk9Ii4yNSIvPgoKICA8IS0tIG5hc2FkYSAtLT4KICA8ZyB0cmFuc2Zvcm09InJvdGF0ZSgyMCAxNTAgMTAwKSI+CiAgICA8cmVjdCB4PSIxNDIiIHk9IjgiIHdpZHRoPSIxNyIgaGVpZ2h0PSIxMTIiIHJ4PSI4IiBmaWxsPSJ1cmwoI3BOYXNhZGEpIi8+CiAgICA8cmVjdCB4PSIxNDUiIHk9IjEyIiB3aWR0aD0iNSIgaGVpZ2h0PSIxMDQiIHJ4PSIyLjUiIGZpbGw9IiNmZmYiIGZpbGwtb3BhY2l0eT0iLjU1Ii8+CiAgICA8IS0tIHRlbG8gbWV0bHkgLS0+CiAgICA8cGF0aCBkPSJNMTEyIDExOGg3Nmw4IDI2aC05MnoiIGZpbGw9InVybCgjcFRlbG8pIi8+CiAgICA8cGF0aCBkPSJNMTEyIDExOGg3NmwzIDloLTgyeiIgZmlsbD0iI2ZmZiIgZmlsbC1vcGFjaXR5PSIuNDUiLz4KICAgIDwhLS0gc3RldGlueSAtLT4KICAgIDxwYXRoIGQ9Ik0xMDQgMTQ0aDkybDEwIDQ0aC0xMTJ6IiBmaWxsPSJ1cmwoI3BTdGV0aW55KSIvPgogICAgPGcgc3Ryb2tlPSIjMDAwIiBzdHJva2Utb3BhY2l0eT0iLjI4IiBzdHJva2Utd2lkdGg9IjMiPgogICAgICA8cGF0aCBkPSJNMTIwIDE0NmwtNSA0ME0xMzggMTQ2bC0zIDQwTTE1NiAxNDZsMCA0ME0xNzQgMTQ2bDMgNDBNMTkwIDE0Nmw2IDQwIi8+CiAgICA8L2c+CiAgPC9nPgoKICA8IS0tIHptZXRlbmUgdHJpZXNreSBhIGNpYXJ5IHBvaHlidSAtLT4KICA8ZyBmaWxsPSIjZmZmIiBmaWxsLW9wYWNpdHk9Ii43NSI+CiAgICA8Y2lyY2xlIGN4PSI0NCIgY3k9IjE2OCIgcj0iNSIvPjxjaXJjbGUgY3g9IjYyIiBjeT0iMTYwIiByPSIzLjYiLz4KICAgIDxjaXJjbGUgY3g9IjMwIiBjeT0iMTU4IiByPSIzIi8+PGNpcmNsZSBjeD0iNzYiIGN5PSIxNzAiIHI9IjMuMiIvPgogIDwvZz4KICA8ZyBmaWxsPSJub25lIiBzdHJva2U9IiNmZmYiIHN0cm9rZS1vcGFjaXR5PSIuNSIgc3Ryb2tlLXdpZHRoPSI1IiBzdHJva2UtbGluZWNhcD0icm91bmQiPgogICAgPHBhdGggZD0iTTIyIDEzMmgzNE0xNCAxNDhoMjZNMzQgMTE2aDI2Ii8+CiAgPC9nPgo8L3N2Zz4K';
        const OBR_MERANIE = 'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyNjAgMjQwIiB3aWR0aD0iMjYwIiBoZWlnaHQ9IjI0MCI+CiAgPGRlZnM+CiAgICA8bGluZWFyR3JhZGllbnQgaWQ9Im1Lb3YiIHgxPSIwIiB5MT0iMCIgeDI9IjAiIHkyPSIxIj4KICAgICAgPHN0b3Agb2Zmc2V0PSIwIiBzdG9wLWNvbG9yPSIjZmZmIiBzdG9wLW9wYWNpdHk9Ii45OCIvPgogICAgICA8c3RvcCBvZmZzZXQ9IjEiIHN0b3AtY29sb3I9IiMwMDAiIHN0b3Atb3BhY2l0eT0iLjM4Ii8+CiAgICA8L2xpbmVhckdyYWRpZW50PgogICAgPGxpbmVhckdyYWRpZW50IGlkPSJtS292ViIgeDE9IjAiIHkxPSIwIiB4Mj0iMSIgeTI9IjAiPgogICAgICA8c3RvcCBvZmZzZXQ9IjAiIHN0b3AtY29sb3I9IiNmZmYiIHN0b3Atb3BhY2l0eT0iLjk1Ii8+CiAgICAgIDxzdG9wIG9mZnNldD0iMSIgc3RvcC1jb2xvcj0iIzAwMCIgc3RvcC1vcGFjaXR5PSIuMzIiLz4KICAgIDwvbGluZWFyR3JhZGllbnQ+CiAgICA8bGluZWFyR3JhZGllbnQgaWQ9Im1EaXNwIiB4MT0iMCIgeTE9IjAiIHgyPSIwIiB5Mj0iMSI+CiAgICAgIDxzdG9wIG9mZnNldD0iMCIgc3RvcC1jb2xvcj0iIzAwMCIgc3RvcC1vcGFjaXR5PSIuNDIiLz4KICAgICAgPHN0b3Agb2Zmc2V0PSIxIiBzdG9wLWNvbG9yPSIjMDAwIiBzdG9wLW9wYWNpdHk9Ii4yMCIvPgogICAgPC9saW5lYXJHcmFkaWVudD4KICAgIDxsaW5lYXJHcmFkaWVudCBpZD0ibURpZWwiIHgxPSIwIiB5MT0iMCIgeDI9IjAiIHkyPSIxIj4KICAgICAgPHN0b3Agb2Zmc2V0PSIwIiBzdG9wLWNvbG9yPSIjZmZmIiBzdG9wLW9wYWNpdHk9Ii44NSIvPgogICAgICA8c3RvcCBvZmZzZXQ9IjEiIHN0b3AtY29sb3I9IiMwMDAiIHN0b3Atb3BhY2l0eT0iLjQwIi8+CiAgICA8L2xpbmVhckdyYWRpZW50PgogICAgPHJhZGlhbEdyYWRpZW50IGlkPSJtVGllbiIgY3g9IjUwJSIgY3k9IjUwJSIgcj0iNTAlIj4KICAgICAgPHN0b3Agb2Zmc2V0PSIwIiBzdG9wLWNvbG9yPSIjMDAwIiBzdG9wLW9wYWNpdHk9Ii40NSIvPgogICAgICA8c3RvcCBvZmZzZXQ9IjEiIHN0b3AtY29sb3I9IiMwMDAiIHN0b3Atb3BhY2l0eT0iMCIvPgogICAgPC9yYWRpYWxHcmFkaWVudD4KICA8L2RlZnM+CgogIDxlbGxpcHNlIGN4PSIxMzAiIGN5PSIyMTQiIHJ4PSI5OCIgcnk9IjE1IiBmaWxsPSJ1cmwoI21UaWVuKSIvPgoKICA8IS0tIG1lcmFuw70gdmFsZWMgLS0+CiAgPHJlY3QgeD0iOTYiIHk9IjEyMCIgd2lkdGg9Ijc0IiBoZWlnaHQ9IjcwIiByeD0iNCIgZmlsbD0idXJsKCNtRGllbCkiLz4KICA8ZWxsaXBzZSBjeD0iMTMzIiBjeT0iMTIwIiByeD0iMzciIHJ5PSIxMSIgZmlsbD0iI2ZmZiIgZmlsbC1vcGFjaXR5PSIuODUiLz4KICA8ZWxsaXBzZSBjeD0iMTMzIiBjeT0iMTIwIiByeD0iMjQiIHJ5PSI2IiBmaWxsPSIjMDAwIiBmaWxsLW9wYWNpdHk9Ii4xOCIvPgoKICA8IS0tIHByYXZpdGtvIHBvc3V2bmVobyBtZXJhZGxhIC0tPgogIDxyZWN0IHg9IjI2IiB5PSI3NCIgd2lkdGg9IjIwOCIgaGVpZ2h0PSIyMiIgcng9IjUiIGZpbGw9InVybCgjbUtvdikiLz4KICA8ZyBzdHJva2U9IiMwMDAiIHN0cm9rZS1vcGFjaXR5PSIuMzAiIHN0cm9rZS13aWR0aD0iMi41Ij4KICAgIDxwYXRoIGQ9Ik00MCA3NHY5TTU0IDc0djEzTTY4IDc0djlNODIgNzR2MTNNOTYgNzR2OU0xMTAgNzR2MTNNMTI0IDc0djlNMTM4IDc0djEzTTE1MiA3NHY5TTE2NiA3NHYxMyIvPgogIDwvZz4KICA8cmVjdCB4PSIyNiIgeT0iNzQiIHdpZHRoPSIyMDgiIGhlaWdodD0iNiIgcng9IjMiIGZpbGw9IiNmZmYiIGZpbGwtb3BhY2l0eT0iLjQ1Ii8+CgogIDwhLS0gcGV2bmEgY2VsdXN0IC0tPgogIDxwYXRoIGQ9Ik00MCA5NmgyMnY1NmwtMjIgOHoiIGZpbGw9InVybCgjbUtvdlYpIi8+CiAgPCEtLSBwb3N1dm5hIGNlbHVzdCBzIGRpc3BsZWpvbSAtLT4KICA8cGF0aCBkPSJNMTY0IDk2aDIydjY0bC0yMi04eiIgZmlsbD0idXJsKCNtS292VikiLz4KICA8cmVjdCB4PSIxNTAiIHk9IjMwIiB3aWR0aD0iODYiIGhlaWdodD0iNDgiIHJ4PSI3IiBmaWxsPSJ1cmwoI21Lb3YpIi8+CiAgPHJlY3QgeD0iMTU4IiB5PSIzOCIgd2lkdGg9IjcwIiBoZWlnaHQ9IjMwIiByeD0iNCIgZmlsbD0idXJsKCNtRGlzcCkiLz4KICA8ZyBmaWxsPSIjZmZmIiBmaWxsLW9wYWNpdHk9Ii45Ij4KICAgIDxyZWN0IHg9IjE2NCIgeT0iNDYiIHdpZHRoPSI3IiBoZWlnaHQ9IjE2IiByeD0iMiIvPgogICAgPHJlY3QgeD0iMTc2IiB5PSI0NiIgd2lkdGg9IjciIGhlaWdodD0iMTYiIHJ4PSIyIi8+CiAgICA8cmVjdCB4PSIxOTIiIHk9IjQ2IiB3aWR0aD0iNyIgaGVpZ2h0PSIxNiIgcng9IjIiLz4KICAgIDxyZWN0IHg9IjIwNiIgeT0iNTgiIHdpZHRoPSI1IiBoZWlnaHQ9IjUiIHJ4PSIxLjUiLz4KICAgIDxyZWN0IHg9IjIxNiIgeT0iNDYiIHdpZHRoPSI3IiBoZWlnaHQ9IjE2IiByeD0iMiIvPgogIDwvZz4KPC9zdmc+Cg==';
        const OBR_PROGRAMOVANIE = 'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyNjAgMjQwIiB3aWR0aD0iMjYwIiBoZWlnaHQ9IjI0MCI+CiAgPGRlZnM+CiAgICA8bGluZWFyR3JhZGllbnQgaWQ9ImdSYW0iIHgxPSIwIiB5MT0iMCIgeDI9IjAiIHkyPSIxIj4KICAgICAgPHN0b3Agb2Zmc2V0PSIwIiBzdG9wLWNvbG9yPSIjZmZmIiBzdG9wLW9wYWNpdHk9Ii45NSIvPgogICAgICA8c3RvcCBvZmZzZXQ9IjEiIHN0b3AtY29sb3I9IiMwMDAiIHN0b3Atb3BhY2l0eT0iLjM4Ii8+CiAgICA8L2xpbmVhckdyYWRpZW50PgogICAgPGxpbmVhckdyYWRpZW50IGlkPSJnT2JyYXoiIHgxPSIwIiB5MT0iMCIgeDI9IjEiIHkyPSIxIj4KICAgICAgPHN0b3Agb2Zmc2V0PSIwIiBzdG9wLWNvbG9yPSIjMDAwIiBzdG9wLW9wYWNpdHk9Ii40NSIvPgogICAgICA8c3RvcCBvZmZzZXQ9IjEiIHN0b3AtY29sb3I9IiMwMDAiIHN0b3Atb3BhY2l0eT0iLjE4Ii8+CiAgICA8L2xpbmVhckdyYWRpZW50PgogICAgPGxpbmVhckdyYWRpZW50IGlkPSJnTm9oYSIgeDE9IjAiIHkxPSIwIiB4Mj0iMSIgeTI9IjAiPgogICAgICA8c3RvcCBvZmZzZXQ9IjAiIHN0b3AtY29sb3I9IiNmZmYiIHN0b3Atb3BhY2l0eT0iLjg1Ii8+CiAgICAgIDxzdG9wIG9mZnNldD0iMSIgc3RvcC1jb2xvcj0iIzAwMCIgc3RvcC1vcGFjaXR5PSIuMzUiLz4KICAgIDwvbGluZWFyR3JhZGllbnQ+CiAgICA8cmFkaWFsR3JhZGllbnQgaWQ9ImdUaWVuMiIgY3g9IjUwJSIgY3k9IjUwJSIgcj0iNTAlIj4KICAgICAgPHN0b3Agb2Zmc2V0PSIwIiBzdG9wLWNvbG9yPSIjMDAwIiBzdG9wLW9wYWNpdHk9Ii40NSIvPgogICAgICA8c3RvcCBvZmZzZXQ9IjEiIHN0b3AtY29sb3I9IiMwMDAiIHN0b3Atb3BhY2l0eT0iMCIvPgogICAgPC9yYWRpYWxHcmFkaWVudD4KICA8L2RlZnM+CgogIDxlbGxpcHNlIGN4PSIxMzAiIGN5PSIyMjAiIHJ4PSI5NCIgcnk9IjE0IiBmaWxsPSJ1cmwoI2dUaWVuMikiLz4KCiAgPCEtLSBtb25pdG9yIC0tPgogIDxyZWN0IHg9IjIyIiB5PSIyMiIgd2lkdGg9IjIxNiIgaGVpZ2h0PSIxNDYiIHJ4PSIxMiIgZmlsbD0idXJsKCNnUmFtKSIvPgogIDxyZWN0IHg9IjM0IiB5PSIzNCIgd2lkdGg9IjE5MiIgaGVpZ2h0PSIxMTIiIHJ4PSI2IiBmaWxsPSJ1cmwoI2dPYnJheikiLz4KICA8cmVjdCB4PSIyMiIgeT0iMjIiIHdpZHRoPSIyMTYiIGhlaWdodD0iOCIgcng9IjQiIGZpbGw9IiNmZmYiIGZpbGwtb3BhY2l0eT0iLjQ1Ii8+CiAgPCEtLSBvZGxlc2sgbmEgc2tsZSAtLT4KICA8cGF0aCBkPSJNMzQgMTQ2bDcwLTExMmgzNGwtNzAgMTEyeiIgZmlsbD0iI2ZmZiIgZmlsbC1vcGFjaXR5PSIuMDkiLz4KCiAgPCEtLSBub2hhIC0tPgogIDxwYXRoIGQ9Ik0xMTIgMTY4aDM2bDggMjhoLTUyeiIgZmlsbD0idXJsKCNnTm9oYSkiLz4KICA8cmVjdCB4PSI4MiIgeT0iMTk2IiB3aWR0aD0iOTYiIGhlaWdodD0iMTQiIHJ4PSI3IiBmaWxsPSJ1cmwoI2dOb2hhKSIvPgoKICA8IS0tIGtvZCBuYSBvYnJhem92a2UgLS0+CiAgPGcgZmlsbD0iI2ZmZiI+CiAgICA8cmVjdCB4PSI0NiIgeT0iNDYiIHdpZHRoPSI0MCIgaGVpZ2h0PSI3IiByeD0iMy41IiBmaWxsLW9wYWNpdHk9Ii44NSIvPgogICAgPHJlY3QgeD0iOTIiIHk9IjQ2IiB3aWR0aD0iNjIiIGhlaWdodD0iNyIgcng9IjMuNSIgZmlsbC1vcGFjaXR5PSIuNDUiLz4KICAgIDxyZWN0IHg9IjU4IiB5PSI2MiIgd2lkdGg9IjU0IiBoZWlnaHQ9IjciIHJ4PSIzLjUiIGZpbGwtb3BhY2l0eT0iLjYwIi8+CiAgICA8cmVjdCB4PSIxMTgiIHk9IjYyIiB3aWR0aD0iMzQiIGhlaWdodD0iNyIgcng9IjMuNSIgZmlsbC1vcGFjaXR5PSIuODUiLz4KICAgIDxyZWN0IHg9IjU4IiB5PSI3OCIgd2lkdGg9IjMwIiBoZWlnaHQ9IjciIHJ4PSIzLjUiIGZpbGwtb3BhY2l0eT0iLjg1Ii8+CiAgICA8cmVjdCB4PSI5NCIgeT0iNzgiIHdpZHRoPSI3MiIgaGVpZ2h0PSI3IiByeD0iMy41IiBmaWxsLW9wYWNpdHk9Ii40NSIvPgogICAgPHJlY3QgeD0iNDYiIHk9Ijk0IiB3aWR0aD0iNDgiIGhlaWdodD0iNyIgcng9IjMuNSIgZmlsbC1vcGFjaXR5PSIuNjAiLz4KICAgIDxyZWN0IHg9IjEwMCIgeT0iOTQiIHdpZHRoPSIyNiIgaGVpZ2h0PSI3IiByeD0iMy41IiBmaWxsLW9wYWNpdHk9Ii44NSIvPgogICAgPHJlY3QgeD0iNDYiIHk9IjExMCIgd2lkdGg9IjM0IiBoZWlnaHQ9IjciIHJ4PSIzLjUiIGZpbGwtb3BhY2l0eT0iLjQ1Ii8+CiAgICA8cmVjdCB4PSI4NiIgeT0iMTEwIiB3aWR0aD0iNTgiIGhlaWdodD0iNyIgcng9IjMuNSIgZmlsbC1vcGFjaXR5PSIuNzAiLz4KICAgIDxyZWN0IHg9IjQ2IiB5PSIxMjYiIHdpZHRoPSIxOCIgaGVpZ2h0PSI3IiByeD0iMy41IiBmaWxsLW9wYWNpdHk9Ii45Ii8+CiAgPC9nPgogIDwhLS0gemF0dm9ya3kgLS0+CiAgPGcgZmlsbD0ibm9uZSIgc3Ryb2tlPSIjZmZmIiBzdHJva2Utb3BhY2l0eT0iLjkiIHN0cm9rZS13aWR0aD0iNyIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIiBzdHJva2UtbGluZWpvaW49InJvdW5kIj4KICAgIDxwYXRoIGQ9Ik0xODYgNzRsLTE2IDE2IDE2IDE2Ii8+CiAgICA8cGF0aCBkPSJNMjAwIDc0bDE2IDE2LTE2IDE2Ii8+CiAgPC9nPgo8L3N2Zz4K';
        const OBR_CHYBA = 'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyNjAgMjQwIiB3aWR0aD0iMjYwIiBoZWlnaHQ9IjI0MCI+CiAgPGRlZnM+CiAgICA8bGluZWFyR3JhZGllbnQgaWQ9ImVUcm9qIiB4MT0iMCIgeTE9IjAiIHgyPSIxIiB5Mj0iMSI+CiAgICAgIDxzdG9wIG9mZnNldD0iMCIgc3RvcC1jb2xvcj0iI2ZmZiIgc3RvcC1vcGFjaXR5PSIuOTgiLz4KICAgICAgPHN0b3Agb2Zmc2V0PSIuNTUiIHN0b3AtY29sb3I9IiNmZmYiIHN0b3Atb3BhY2l0eT0iLjc4Ii8+CiAgICAgIDxzdG9wIG9mZnNldD0iMSIgc3RvcC1jb2xvcj0iIzAwMCIgc3RvcC1vcGFjaXR5PSIuNDAiLz4KICAgIDwvbGluZWFyR3JhZGllbnQ+CiAgICA8bGluZWFyR3JhZGllbnQgaWQ9ImVPYnJheiIgeDE9IjAiIHkxPSIwIiB4Mj0iMSIgeTI9IjEiPgogICAgICA8c3RvcCBvZmZzZXQ9IjAiIHN0b3AtY29sb3I9IiMwMDAiIHN0b3Atb3BhY2l0eT0iLjQwIi8+CiAgICAgIDxzdG9wIG9mZnNldD0iMSIgc3RvcC1jb2xvcj0iIzAwMCIgc3RvcC1vcGFjaXR5PSIuMTYiLz4KICAgIDwvbGluZWFyR3JhZGllbnQ+CiAgICA8bGluZWFyR3JhZGllbnQgaWQ9ImVSYW0iIHgxPSIwIiB5MT0iMCIgeDI9IjAiIHkyPSIxIj4KICAgICAgPHN0b3Agb2Zmc2V0PSIwIiBzdG9wLWNvbG9yPSIjZmZmIiBzdG9wLW9wYWNpdHk9Ii43MCIvPgogICAgICA8c3RvcCBvZmZzZXQ9IjEiIHN0b3AtY29sb3I9IiMwMDAiIHN0b3Atb3BhY2l0eT0iLjM1Ii8+CiAgICA8L2xpbmVhckdyYWRpZW50PgogICAgPHJhZGlhbEdyYWRpZW50IGlkPSJlVGllbiIgY3g9IjUwJSIgY3k9IjUwJSIgcj0iNTAlIj4KICAgICAgPHN0b3Agb2Zmc2V0PSIwIiBzdG9wLWNvbG9yPSIjMDAwIiBzdG9wLW9wYWNpdHk9Ii40OCIvPgogICAgICA8c3RvcCBvZmZzZXQ9IjEiIHN0b3AtY29sb3I9IiMwMDAiIHN0b3Atb3BhY2l0eT0iMCIvPgogICAgPC9yYWRpYWxHcmFkaWVudD4KICA8L2RlZnM+CgogIDxlbGxpcHNlIGN4PSIxMzAiIGN5PSIyMTYiIHJ4PSI5NiIgcnk9IjE1IiBmaWxsPSJ1cmwoI2VUaWVuKSIvPgoKICA8IS0tIG1vbml0b3IgdiBwb3phZGkgLS0+CiAgPHJlY3QgeD0iMzQiIHk9IjM0IiB3aWR0aD0iMTkyIiBoZWlnaHQ9IjEyNiIgcng9IjEwIiBmaWxsPSJ1cmwoI2VSYW0pIi8+CiAgPHJlY3QgeD0iNDQiIHk9IjQ0IiB3aWR0aD0iMTcyIiBoZWlnaHQ9Ijk2IiByeD0iNSIgZmlsbD0idXJsKCNlT2JyYXopIi8+CiAgPHBhdGggZD0iTTEwMCAxNjBoNjBsOCAyMmgtNzZ6IiBmaWxsPSJ1cmwoI2VSYW0pIi8+CiAgPHJlY3QgeD0iODAiIHk9IjE4MiIgd2lkdGg9IjEwMCIgaGVpZ2h0PSIxMyIgcng9IjYiIGZpbGw9InVybCgjZVJhbSkiLz4KICA8IS0tIHByZXJ1c2VuZSByaWFka3kga29kdSAtLT4KICA8ZyBmaWxsPSIjZmZmIiBmaWxsLW9wYWNpdHk9Ii4zNSI+CiAgICA8cmVjdCB4PSI1NiIgeT0iNTYiIHdpZHRoPSI0MiIgaGVpZ2h0PSI2IiByeD0iMyIvPgogICAgPHJlY3QgeD0iMTA2IiB5PSI1NiIgd2lkdGg9IjI2IiBoZWlnaHQ9IjYiIHJ4PSIzIi8+CiAgICA8cmVjdCB4PSI1NiIgeT0iMTIyIiB3aWR0aD0iMzQiIGhlaWdodD0iNiIgcng9IjMiLz4KICAgIDxyZWN0IHg9Ijk4IiB5PSIxMjIiIHdpZHRoPSI1MiIgaGVpZ2h0PSI2IiByeD0iMyIvPgogIDwvZz4KCiAgPCEtLSB2eXN0cmF6bnkgdHJvanVob2xuaWsgLS0+CiAgPHBhdGggZD0iTTEzMCA1Mmw4NiAxMjJINDR6IiBmaWxsPSIjMDAwIiBmaWxsLW9wYWNpdHk9Ii4zMCIgdHJhbnNmb3JtPSJ0cmFuc2xhdGUoNCA2KSIvPgogIDxwYXRoIGQ9Ik0xMzAgNTJjNiAwIDEwIDMgMTMgOGw3MyAxMDRjNSA4IDEgMTYtOSAxNkg1M2MtMTAgMC0xNC04LTktMTZsNzMtMTA0YzMtNSA3LTggMTMtOHoiIGZpbGw9InVybCgjZVRyb2opIi8+CiAgPHBhdGggZD0iTTEzMCA2NmMzIDAgNSAxIDcgNGw2OCA5N0g1NWw2OC05N2MyLTMgNC00IDctNHoiIGZpbGw9IiMwMDAiIGZpbGwtb3BhY2l0eT0iLjEwIi8+CiAgPHBhdGggZD0iTTEzMCA2NmMzIDAgNSAxIDcgNGwyMCAyOC02NiA2Mi02LTggNDUtODZ6IiBmaWxsPSIjZmZmIiBmaWxsLW9wYWNpdHk9Ii4zMCIvPgogIDwhLS0gdnlrcmljbmlrIC0tPgogIDxwYXRoIGQ9Ik0xMjEgOTZoMThsLTQgNTJoLTEweiIgZmlsbD0iIzAwMCIgZmlsbC1vcGFjaXR5PSIuNTUiLz4KICA8Y2lyY2xlIGN4PSIxMzAiIGN5PSIxNjIiIHI9IjkuNSIgZmlsbD0iIzAwMCIgZmlsbC1vcGFjaXR5PSIuNTUiLz4KPC9zdmc+Cg==';
        const OBR_UDRZBA = 'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyNjAgMjQwIiB3aWR0aD0iMjYwIiBoZWlnaHQ9IjI0MCI+CiAgPGRlZnM+CiAgICA8bGluZWFyR3JhZGllbnQgaWQ9ImRLb3YiIHgxPSIwIiB5MT0iMCIgeDI9IjAiIHkyPSIxIj4KICAgICAgPHN0b3Agb2Zmc2V0PSIwIiBzdG9wLWNvbG9yPSIjZmZmIiBzdG9wLW9wYWNpdHk9Ii45NSIvPgogICAgICA8c3RvcCBvZmZzZXQ9IjEiIHN0b3AtY29sb3I9IiMwMDAiIHN0b3Atb3BhY2l0eT0iLjQwIi8+CiAgICA8L2xpbmVhckdyYWRpZW50PgogICAgPGxpbmVhckdyYWRpZW50IGlkPSJkS2x1YyIgeDE9IjAiIHkxPSIwIiB4Mj0iMCIgeTI9IjEiPgogICAgICA8c3RvcCBvZmZzZXQ9IjAiIHN0b3AtY29sb3I9IiNmZmYiIHN0b3Atb3BhY2l0eT0iLjk3Ii8+CiAgICAgIDxzdG9wIG9mZnNldD0iLjQ4IiBzdG9wLWNvbG9yPSIjZmZmIiBzdG9wLW9wYWNpdHk9Ii42MiIvPgogICAgICA8c3RvcCBvZmZzZXQ9IjEiIHN0b3AtY29sb3I9IiMwMDAiIHN0b3Atb3BhY2l0eT0iLjQyIi8+CiAgICA8L2xpbmVhckdyYWRpZW50PgogICAgPHJhZGlhbEdyYWRpZW50IGlkPSJkS29sZXNvIiBjeD0iMzYlIiBjeT0iMzAlIiByPSI3NCUiPgogICAgICA8c3RvcCBvZmZzZXQ9IjAiIHN0b3AtY29sb3I9IiNmZmYiIHN0b3Atb3BhY2l0eT0iLjkyIi8+CiAgICAgIDxzdG9wIG9mZnNldD0iLjU4IiBzdG9wLWNvbG9yPSIjZmZmIiBzdG9wLW9wYWNpdHk9Ii40MiIvPgogICAgICA8c3RvcCBvZmZzZXQ9IjEiIHN0b3AtY29sb3I9IiMwMDAiIHN0b3Atb3BhY2l0eT0iLjQ2Ii8+CiAgICA8L3JhZGlhbEdyYWRpZW50PgogICAgPHJhZGlhbEdyYWRpZW50IGlkPSJkVGllbiIgY3g9IjUwJSIgY3k9IjUwJSIgcj0iNTAlIj4KICAgICAgPHN0b3Agb2Zmc2V0PSIwIiBzdG9wLWNvbG9yPSIjMDAwIiBzdG9wLW9wYWNpdHk9Ii40OCIvPgogICAgICA8c3RvcCBvZmZzZXQ9IjEiIHN0b3AtY29sb3I9IiMwMDAiIHN0b3Atb3BhY2l0eT0iMCIvPgogICAgPC9yYWRpYWxHcmFkaWVudD4KICA8L2RlZnM+CgogIDxlbGxpcHNlIGN4PSIxMzAiIGN5PSIyMTYiIHJ4PSIxMDQiIHJ5PSIxNSIgZmlsbD0idXJsKCNkVGllbikiLz4KCiAgPCEtLSB2ZWxrZSBvenViZW5lIGtvbGVzbyB2bGF2byBkb2xlIC0tPgogIDxnIHRyYW5zZm9ybT0idHJhbnNsYXRlKDYyLDE1MCkiPgogICAgPHBhdGggZD0iTS0xMS03MmgyMmw0IDE3IDE2IDcgMTQtMTAgMTUgMTUtMTAgMTQgNyAxNiAxNyA0djIybC0xNyA0LTcgMTYgMTAgMTQtMTUgMTUtMTQtMTAtMTYgNy00IDE3aC0yMmwtNC0xNy0xNi03LTE0IDEwLTE1LTE1IDEwLTE0LTctMTYtMTctNHYtMjJsMTctNCA3LTE2LTEwLTE0IDE1LTE1IDE0IDEwIDE2LTd6IiBmaWxsPSJ1cmwoI2RLb2xlc28pIi8+CiAgICA8cGF0aCBkPSJNLTExLTcyaDIybDQgMTcgMTYgNyAxNC0xMCAxNSAxNS0xMCAxNCA3IDE2IDE3IDR2NmwtMTctNC03LTE2IDEwLTE0LTE1LTE1LTE0IDEwLTE2LTctNC0xN2gtMjJsLTQgMTctMTYgNy0xNC0xMC0xNSAxNSAxMCAxNC03IDE2LTE3IDR2LTZsMTctNCA3LTE2LTEwLTE0IDE1LTE1IDE0IDEwIDE2LTd6IiBmaWxsPSIjZmZmIiBmaWxsLW9wYWNpdHk9Ii40MiIvPgogICAgPGNpcmNsZSByPSI0MSIgZmlsbD0iIzAwMCIgZmlsbC1vcGFjaXR5PSIuMjYiLz4KICAgIDxjaXJjbGUgcj0iMzUiIGZpbGw9InVybCgjZEtvdikiLz4KICAgIDxjaXJjbGUgcj0iMjAiIGZpbGw9IiMwMDAiIGZpbGwtb3BhY2l0eT0iLjM0Ii8+CiAgICA8Y2lyY2xlIHI9IjE0IiBmaWxsPSIjZmZmIiBmaWxsLW9wYWNpdHk9Ii4yNiIvPgogICAgPGcgZmlsbD0iI2ZmZiIgZmlsbC1vcGFjaXR5PSIuMjYiPgogICAgICA8Y2lyY2xlIGN4PSIwIiBjeT0iLTI4IiByPSI1Ii8+CiAgICAgIDxjaXJjbGUgY3g9IjI0IiBjeT0iMTUiIHI9IjUiLz4KICAgICAgPGNpcmNsZSBjeD0iLTI0IiBjeT0iMTUiIHI9IjUiLz4KICAgIDwvZz4KICAgIDxwYXRoIGQ9Ik0tMzUgMGEzNSAzNSAwIDAgMSAyMC0zMmw0IDdBMjcgMjcgMCAwIDAtMjcgMHoiIGZpbGw9IiNmZmYiIGZpbGwtb3BhY2l0eT0iLjU1Ii8+CiAgPC9nPgoKICA8IS0tIG1hbGUgb3p1YmVuZSBrb2xlc28gdnByYXZvIGhvcmUgLS0+CiAgPGcgdHJhbnNmb3JtPSJ0cmFuc2xhdGUoMjEyLDE1NikiPgogICAgPHBhdGggZD0iTS03LTQwaDE0bDMgMTAgMTAgNCA5LTYgMTAgMTAtNiA5IDQgMTAgMTAgM3YxNGwtMTAgMy00IDEwIDYgOS0xMCAxMC05LTYtMTAgNC0zIDEwaC0xNGwtMy0xMC0xMC00LTkgNi0xMC0xMCA2LTktNC0xMC0xMC0zdi0xNGwxMC0zIDQtMTAtNi05IDEwLTEwIDkgNiAxMC00eiIgZmlsbD0idXJsKCNkS29sZXNvKSIvPgogICAgPGNpcmNsZSByPSIyMSIgZmlsbD0iIzAwMCIgZmlsbC1vcGFjaXR5PSIuMjQiLz4KICAgIDxjaXJjbGUgcj0iMTYiIGZpbGw9InVybCgjZEtvdikiLz4KICAgIDxjaXJjbGUgcj0iOCIgZmlsbD0iIzAwMCIgZmlsbC1vcGFjaXR5PSIuMzQiLz4KICAgIDxwYXRoIGQ9Ik0tMTYgMGExNiAxNiAwIDAgMSA5LTE0bDMgNUExMCAxMCAwIDAgMC0xMCAweiIgZmlsbD0iI2ZmZiIgZmlsbC1vcGFjaXR5PSIuNSIvPgogIDwvZz4KCiAgPCEtLSBrb21iaW5vdmFueSBrbHVjIGNleiBzdHJlZCAtLT4KICA8ZyB0cmFuc2Zvcm09InJvdGF0ZSgtMzQgMTMyIDExMCkiPgogICAgPCEtLSB0aWVuIHBvZCBrbHVjb20gLS0+CiAgICA8ZyB0cmFuc2Zvcm09InRyYW5zbGF0ZSgzLDUpIiBmaWxsPSIjMDAwIiBmaWxsLW9wYWNpdHk9Ii4yMiI+CiAgICAgIDxyZWN0IHg9Ijc2IiB5PSIxMDYiIHdpZHRoPSIxMDgiIGhlaWdodD0iMjgiIHJ4PSIxNCIvPgogICAgICA8Y2lyY2xlIGN4PSI4NiIgY3k9IjEyMCIgcj0iMjkiLz4KICAgICAgPHBhdGggZD0iTTE2NiA5MGg1NHYyNGgtMjR2MTRoMjR2MjRoLTU0eiIvPgogICAgPC9nPgogICAgPCEtLSBkcmllayAtLT4KICAgIDxyZWN0IHg9Ijc2IiB5PSIxMDciIHdpZHRoPSIxMDgiIGhlaWdodD0iMjYiIHJ4PSIxMyIgZmlsbD0idXJsKCNkS2x1YykiLz4KICAgIDwhLS0gb2NrbyB2bGF2byAtLT4KICAgIDxwYXRoIGQ9Ik04NiA5MmEyOCAyOCAwIDEgMCAwIDU2IDI4IDI4IDAgMCAwIDAtNTZ6bTAgMTNhMTUgMTUgMCAxIDEgMCAzMCAxNSAxNSAwIDAgMSAwLTMweiIgZmlsbD0idXJsKCNkS2x1YykiLz4KICAgIDxjaXJjbGUgY3g9Ijg2IiBjeT0iMTIwIiByPSIxNSIgZmlsbD0iIzAwMCIgZmlsbC1vcGFjaXR5PSIuMzAiLz4KICAgIDwhLS0gb3R2b3JlbmEgY2VsdXN0IHZwcmF2byAtLT4KICAgIDxwYXRoIGQ9Ik0xNjYgOTBoNTR2MjRoLTI0djEyaDI0djI0aC01NHoiIGZpbGw9InVybCgjZEtsdWMpIi8+CiAgICA8IS0tIHN2ZXRsYSBob3JuYSBocmFuYSBhIG9kbGVza3kgLS0+CiAgICA8cmVjdCB4PSI4MCIgeT0iMTA4IiB3aWR0aD0iMTAwIiBoZWlnaHQ9IjciIHJ4PSIzIiBmaWxsPSIjZmZmIiBmaWxsLW9wYWNpdHk9Ii41NSIvPgogICAgPHBhdGggZD0iTTg2IDkyYTI4IDI4IDAgMCAwLTI2IDE3IDI4IDI4IDAgMCAxIDI2LTExIDI4IDI4IDAgMCAxIDI2IDExIDI4IDI4IDAgMCAwLTI2LTE3eiIgZmlsbD0iI2ZmZiIgZmlsbC1vcGFjaXR5PSIuNjAiLz4KICAgIDxyZWN0IHg9IjgwIiB5PSIxMjciIHdpZHRoPSIxMDAiIGhlaWdodD0iNiIgcng9IjMiIGZpbGw9IiMwMDAiIGZpbGwtb3BhY2l0eT0iLjI2Ii8+CiAgICA8cGF0aCBkPSJNMTY4IDkyaDUwdjZoLTUwek0xNjggMTQ0aDUwdjZoLTUweiIgZmlsbD0iI2ZmZiIgZmlsbC1vcGFjaXR5PSIuMzgiLz4KICA8L2c+Cjwvc3ZnPgo=';

        function injectStyles() {
            if (document.getElementById(STYLE_ID)) return;
            const st = document.createElement('style');
            st.id = STYLE_ID;
            st.textContent = `
/* ---------- pozadie a zakladne karty ---------- */
/* Pozadie HF Slovakia (navrh 3 z 2026-09-16) je vlozene priamo v skripte, aby
   nezaviselo od siete; svetlomodry prechod pod nim je zaloha, keby sa
   obrazok z akehokolvek dovodu nevykreslil. Obrazok je len na <body>,
   vnutorne kontajnery appky su priehladne - inak by sa kreslil viackrat.
   Vynimka je #Application - SAP naň priamo vo svojom CSS vnucuje vlastne
   pozadie (background-image:url(...) !important), ktore je nad <body> a
   nase pozadie by tak celkom prekrylo. Preto musi mat aj #Application
   rovnaky obrazok priamo (selektor s BODY_CLASS ma vyssiu specificitu
   1 trieda + 1 ID nez SAP-acke holé #Application, takze pri rovnakom
   !important vyhrava nas zapis). */
body.${BODY_CLASS}, body.${BODY_CLASS} #Application { background:url(${POZADIE}) center / cover no-repeat fixed,
  linear-gradient(135deg,#eaf2fb 0%,#dbe8f7 55%,#e9f1fb 100%) !important; }
body.${BODY_CLASS} .sapUiBody, body.${BODY_CLASS} .sapMShell, body.${BODY_CLASS} .sapMPage,
body.${BODY_CLASS} .sapMPageBgStandard, body.${BODY_CLASS} .sapMPageBgSolid,
body.${BODY_CLASS} .sapMNav, body.${BODY_CLASS} .sapMApp,
body.${BODY_CLASS} .sapMTB { background:transparent !important; }
/* sapUiBody sedi na samotnom <body>, takze ho horny (potomkovsky) selektor
   nikdy nechyti. Rusime tu preto len farbu - obrazok pozadia z pravidla
   body.${BODY_CLASS} musi ostat, inak by sme si ho zhodili. */
body.${BODY_CLASS}.sapUiBody { background-color:transparent !important; }
body.${BODY_CLASS} .sapMPanel { background:rgba(255,255,255,.8) !important; border:1px solid #e3ebf5 !important;
  border-radius:18px !important; box-shadow:0 4px 18px rgba(16,36,63,.08) !important;
  margin:10px 12px !important; box-sizing:border-box !important;
  /* SAP dava panelu fixnu sirku (px, vypocitanu este pred nasim marginom),
     takze pridany horizontalny margin sa k nej len pripocital a panel
     vytekal z prava. width:calc si od 100% sirky rodica odpocita presne
     nase dva 12px marginy, takze panel aj s marginmi sedi presne do rodica. */
  width:calc(100% - 24px) !important; max-width:calc(100% - 24px) !important; }
/* vnutorne obaly panelu od SAP maju vlastne plne pozadie - lezia nad panelom,
   takze by jeho priehladnost (alfa .8) uplne prekryli */
body.${BODY_CLASS} .sapMPanelWrappingDiv, body.${BODY_CLASS} .sapMPanelWrappingDivTb,
body.${BODY_CLASS} .sapMPanelContent, body.${BODY_CLASS} .sapMPanelBGSolid {
  background:transparent !important; }
body.${BODY_CLASS} #WorkcenterDetail--Tasks_Panel-content { padding-bottom:6px !important; }
body.${BODY_CLASS} .sapMPanelHdr, body.${BODY_CLASS} .sapMPanelHeaderTB {
  background:transparent !important; border:0 !important; }
body.${BODY_CLASS} .sapMPanelHdr .sapMTitle, body.${BODY_CLASS} .sapMPanelHdr .sapMText {
  font-size:12px !important; font-weight:800 !important; letter-spacing:.14em !important;
  text-transform:uppercase !important; color:#4a6285 !important; }

/* ---------- stavove tlacidla ako karty s ikonou ----------
   Vnutro tlacidla (UI5 <bdi id="...-BDI-content">) sa NESMIE menit - UI5 ho
   potrebuje pri stlaceni a v 2.1.0 to tlacidla odstavilo. Ikona a podnadpis
   su preto len pseudo-prvky ::before / ::after na obale .sapMBtnInner a text
   beru z data-* atributov, ktore na obal doplna JS. Skutocny text tlacidla
   zostava netknuty v .sapMBtnContent. */
/* Vysku tlacidla urcuje VYLUCNE tato jedna hodnota. Inline height, ktory na
   tlacidlo (a na .sapMBtnInner) pise SAP, je bez !important, takze ho toto
   pravidlo prebija - v DevTools ho vidno, ale neuplatnuje sa. Vnutorny
   padding do vysky nezasahuje, kym je tu pevna hodnota; .sapMBtnInner ma
   height:100%, takze len vyplni tlacidlo. */
body.${BODY_CLASS} .statusBtn { border-radius:16px !important; border:1px solid #b9cdee !important;
  background-color:#0d4a8f !important; box-shadow:0 6px 16px rgba(16,36,63,.18) !important;
  min-width:190px !important; height:64px !important; box-sizing:border-box !important; }
body.${BODY_CLASS} .statusBtn .sapMBtnInner { padding:12px 18px !important; border-radius:16px !important;
  display:flex !important; align-items:center !important; justify-content:center !important;
  height:100% !important; width:100% !important; box-sizing:border-box !important;
  background-image:none !important; background-color:#0d4a8f !important; }
body.${BODY_CLASS} .statusBtn .sapMBtnContent { justify-content:center !important; text-align:center !important;
  font-size:15px !important; font-weight:800 !important; line-height:1.2 !important;
  white-space:nowrap !important; color:#fff !important; }
body.${BODY_CLASS} .statusBtn .sapMBtnContent, body.${BODY_CLASS} .statusBtn bdi { color:#fff !important; }

/*
 * Obrazok v pozadi troch tlacidiel osobneho stavu. Kreslime ho ako
 * pseudo-prvok ::after samotneho <button>, NIE ako background vnutorneho
 * spanu: tomu totiz farbu aj "background-image:none" zapisuje inline s
 * !important funkcia paint(), a to sa zo stylesheetu prebit neda.
 * mix-blend-mode:soft-light necha povodnu farbu tlacidla dominovat -
 * obrazok ju len jemne presvetli, nesadne si na nu ako plna vrstva.
 * Velkost tlacidla sa nemeni, pseudo-prvok lezi presne v jeho ramci.
 */
body.${BODY_CLASS} .statusBtn[data-pda-obraz] {
  position:relative !important; overflow:hidden !important; }
body.${BODY_CLASS} .statusBtn[data-pda-obraz]::after {
  content:''; position:absolute; inset:0; z-index:2; pointer-events:none;
  border-radius:10px; opacity:.5; mix-blend-mode:soft-light;
  background-repeat:no-repeat; background-position:right -14px center; background-size:auto 150%; }
/* osobny stav (uvodna obrazovka) */
body.${BODY_CLASS} .statusBtn[data-pda-obraz="stretnutia"]::after { background-image:url(${OBR_STRETNUTIA}); }
body.${BODY_CLASS} .statusBtn[data-pda-obraz="prestavka"]::after { background-image:url(${OBR_PRESTAVKA}); }
body.${BODY_CLASS} .statusBtn[data-pda-obraz="cakanie"]::after { background-image:url(${OBR_CAKANIE}); }
/* stav operacie (detail pracoviska) */
body.${BODY_CLASS} .statusBtn[data-pda-obraz="vyroba"]::after { background-image:url(${OBR_VYROBA}); }
body.${BODY_CLASS} .statusBtn[data-pda-obraz="upinanie"]::after { background-image:url(${OBR_UPINANIE}); }
body.${BODY_CLASS} .statusBtn[data-pda-obraz="upratovanie"]::after { background-image:url(${OBR_UPRATOVANIE}); }
body.${BODY_CLASS} .statusBtn[data-pda-obraz="meranie"]::after { background-image:url(${OBR_MERANIE}); }
body.${BODY_CLASS} .statusBtn[data-pda-obraz="programovanie"]::after { background-image:url(${OBR_PROGRAMOVANIE}); }
body.${BODY_CLASS} .statusBtn[data-pda-obraz="chyba"]::after { background-image:url(${OBR_CHYBA}); }
body.${BODY_CLASS} .statusBtn[data-pda-obraz="udrzba"]::after { background-image:url(${OBR_UDRZBA}); }

/* "Hladat vyrobny prikaz" v hornej liste. Text nesedel zvisle na strede -
   ikona a text su dva inline prvky s roznou vyskou riadku a bez flexu sa
   zarovnavaju na uctovnu ciaru, nie na stred. Modra je tá ista, aku ma nase
   tlacidlo "Hladat" vo vyhladavacom okne (#13315c / hover #1c478a), aby
   sedelo k ostatnym neutralnym tlacidlam appky. */
/* Lista, v ktorej tlacidlo sedi, ma pevnu vysku 4rem (64px) a tlacidlo z nej
   vytekalo - trieda sapUiSmallMargin mu dava 1rem margin hore aj dole, takze
   na samotne tlacidlo zostavalo len 32px. Zvisle marginy rusime (bocny
   odstup od okraja listy ostava), tlacidlo centrujeme a mierne znizujeme,
   nech sa do listy zmesti cele. */
body.${BODY_CLASS} #Main--Button_SearchProductionOrder {
  margin-top:0 !important; margin-bottom:0 !important;
  align-self:center !important; height:auto !important; }
body.${BODY_CLASS} #Main--Button_SearchProductionOrder .sapMBtnInner {
  display:flex !important; align-items:center !important; justify-content:center !important;
  height:auto !important; padding:8px 16px !important; box-sizing:border-box !important;
  background:#13315c !important; background-image:none !important;
  border:1px solid #13315c !important; border-radius:10px !important; }
body.${BODY_CLASS} #Main--Button_SearchProductionOrder:hover .sapMBtnInner {
  background:#1c478a !important; border-color:#1c478a !important; }
body.${BODY_CLASS} #Main--Button_SearchProductionOrder .sapMBtnContent,
body.${BODY_CLASS} #Main--Button_SearchProductionOrder bdi,
body.${BODY_CLASS} #Main--Button_SearchProductionOrder .sapUiIcon { color:#fff !important; }
/* Vsetky tlacidla v boxe so stavmi maju rovnaky jemny ram a tien (ten isty
   zapis ako hover tlacidiel v hornom pruhu). Ide to na samotny <button> -
   na .sapMBtnInner pise farbu ramu inline s !important funkcia paint(),
   co by sa zo stylesheetu uz prebit nedalo. */
body.${BODY_CLASS} #WorkcenterDetail--Order_Status_Flexbox .sapMBtn {
  border-color:#7ba4ee !important;
  box-shadow:0 3px 10px rgba(16,36,63,.12) !important; }
/* aby nadvihnutie pri prechode mysou neprislo o svoj vyraznejsi tien */
body.${BODY_CLASS} #WorkcenterDetail--Order_Status_Flexbox .sapMBtn:hover {
  box-shadow:0 10px 20px rgba(16,36,63,.30) !important; }
/* VYKRES: modry ram AZ pri prechode mysou, bez tienu. Musi ist cez obal
   (#...wrapper__ > button), nie cez vlastne ID tlacidla: nizsie v tomto
   subore mu ram nastavuje presne taky selektor a ten ma vyssiu specificitu
   (ID + 2 typy) nez samotne ID tlacidla.
   Hover pre BOM je pri jeho ostatnych pravidlach nizsie - tam musi prebit
   ram, ktory si sam nastavuje .sapMBtnInner. */
body.${BODY_CLASS} #__pda_order_drawing_wrapper__ > button:hover {
  border-color:#7ba4ee !important; background:#f6f9ff !important; }

/* ---------- nadpisy sekcii ---------- */
.nd-nadpis { font:800 12px/1.3 -apple-system,"Segoe UI",Roboto,sans-serif; letter-spacing:.14em;
  text-transform:uppercase; color:#4a6285; margin:6px 0 8px 2px; }
.nd-v-riadku { flex:0 0 100% !important; width:100% !important; box-sizing:border-box; margin:2px 0 6px 6px !important; }
body.${BODY_CLASS} #WorkcenterDetail--Order_Status_Flexbox,
body.${BODY_CLASS} #WorkcenterDetail--OrderHeader_FlexBox,
body.${BODY_CLASS} #WorkcenterDetail--TimerCharts_FlexBox { flex-wrap:wrap !important; }

/* ---------- pravy stlpec appky ---------- */
body.${BODY_CLASS} #WorkcenterDetail--Order_Status_Flexbox {
  background:#fff !important; border:1px solid #e3ebf5 !important;
  box-shadow:0 4px 14px rgba(16,36,63,.06) !important;
  padding:10px 14px !important; box-sizing:border-box !important; }

/* ---------- patka ---------- */
#${FOOTER_ID} { position:fixed; left:0; right:0; bottom:0; height:34px; z-index:4;
  display:flex; align-items:center; gap:10px; padding:0 18px; box-sizing:border-box;
  background:rgba(255,255,255,.82); backdrop-filter:blur(6px);
  border-top:1px solid #e3ebf5; font:12px/1 -apple-system,"Segoe UI",Roboto,sans-serif;
  color:#7d8ea8; }
#${FOOTER_ID} img { height:18px; display:block; }
#${FOOTER_ID} .nd-f-n { font-weight:800; color:#13315c; letter-spacing:.04em; }
#${FOOTER_ID} .nd-f-r { margin-left:auto; }
body.${BODY_CLASS} { padding-bottom:34px !important; box-sizing:border-box; }

/* ---------- horny pruh: navigacia ako biele pilulky, odhlasenie cervene ---------- */
.nd-topbar { background:rgba(255,255,255,.88) !important; border-bottom:1px solid #e3ebf5 !important; }
.nd-topbar .sapMBtn .sapMBtnInner { background:#fff !important; border:1px solid #dfe7f2 !important;
  border-radius:12px !important; box-shadow:0 2px 8px rgba(16,36,63,.10) !important; color:#13315c !important;
  font-weight:700 !important; padding:6px 14px !important; }
.nd-topbar .sapMBtn .sapMBtnContent, .nd-topbar .sapMBtn bdi, .nd-topbar .sapMBtn .sapUiIcon { color:#13315c !important; }
.nd-topbar [id$="Button_Logout"] .sapMBtnInner { background:#e23b3b !important; border-color:#c72f2f !important; }
.nd-topbar [id$="Button_Logout"] .sapMBtnContent, .nd-topbar [id$="Button_Logout"] bdi,
.nd-topbar [id$="Button_Logout"] .sapUiIcon { color:#fff !important; }
/* expand/collapse sipka panela Workcenter_Panel sa niekedy ocitne pod tou istou
   "nd-topbar" znackou (hornyPruh() vyssie) - vynimka nesmie menit specificitu
   pravidiel vyssie (tym sa raz uz omylom rozbila farba tlacidla Odhlasenie),
   preto samostatne pravidlo s umelo zvysenou specificitou (opakovany atribut) */
[id*="expandButton"][id*="expandButton"][id*="expandButton"] .sapMBtnInner {
  border: none !important;
  box-shadow: none !important;
  border-radius: 0 !important;
  padding: 0 !important;
}
/* Selektory nizsie su zamerne bez view-prefixu (Main--, WorkcenterDetail--, ...),
   aby platili na vsetkych podstrankach - SAPUI5 ma v kazdom view rovnaky
   nazov controlu, len s inym prefixom pred "--". */
header.sapMPageHeader {
    min-height: 62px !important;
    position: relative !important;
}
[id$="Bar_Header"] {
    background-color: #ffffff !important;
    padding-top: 8px !important;
    padding-bottom: 8px !important;
    box-sizing: border-box !important;
    min-height: 62px !important;
    position: relative !important;
    border-bottom: none !important;
}
[id$="Bar_Header-BarRight"] {
    padding-right: 16px !important;
    box-sizing: border-box !important;
}
/* header narastol z povodnych 44px na 62px (54px + 8px padding-bottom);
   sekcia s obsahom je vzdy priamy "section" susediaci za header.sapMPageHeader */
header.sapMPageHeader + section {
    top: 62px !important;
}
[id$="Hour_Date_Title"] {
    margin-bottom: 0 !important;
    margin-top: 0 !important;
}
#__pda_search_sidebar__ {
    margin: 12px !important;
    border-radius: 18px !important;
    /* rovnaky vypocet ako pri .sapMPanel - sirku si tento kontajner nastavuje
       inline na 100%, takze bez !important by sa nas prepocet neuplatnil a
       panel by aj s marginmi vytekal vpravo mimo rodica */
    width: calc(100% - 24px) !important;
    max-width: calc(100% - 24px) !important;
    box-sizing: border-box !important;
    /* mierne priehladne, aby bolo cez kartu jemne vidiet pozadie
       (inline background-color prebijeme cez !important) */
    background-color: rgba(255,255,255,.8) !important;
}
#__pda_custom_search_ui__ {
    padding: 16px !important;
    box-sizing: border-box !important;
}
#__pda_custom_search_ui__ > div:first-child {
    padding: 0 !important;
    margin-left: 20px !important;
}
[id$="Bar_Header-BarRight"] .sapMBtnInner {
    display: flex !important;
    align-items: center !important;
    justify-content: center !important;
    border-radius: 10px !important;
}
[id$="Bar_Header-BarRight"] .sapMBtnInner img,
[id$="Bar_Header-BarRight"] .sapMBtnInner .sapMBtnContent,
[id$="Bar_Header-BarRight"] .sapMBtnInner .sapMBtnContent bdi {
    vertical-align: middle !important;
}
/* rovnaky hover ako pri riadkoch zakaziek (#WorkcenterDetail--Work_List .pda-pill-on:hover) */
[id$="Bar_Header-BarRight"] .sapMBtn:hover .sapMBtnInner {
    border-color: #7ba4ee !important;
    box-shadow: 0 3px 10px rgba(16,36,63,.12) !important;
}

/* ---------- riadok akcii pod stavmi: Vykres | Components/BOM | Stroj ON/OFF | Operation Complete ---------- */
/* flex-wrap je tu kvoli nadpisu "Dokumentácia": ma .nd-v-riadku s
   flex-basis:100%, co ho posunie na vlastny riadok len vtedy, ked sa riadok
   moze zalomit - inak by sa vsetko len stlacilo do jedneho riadku */
body.${BODY_CLASS} #__pda_detail_rightcol__ { flex:0 0 100% !important; width:100% !important; max-width:none !important;
  flex-direction:row !important; flex-wrap:wrap !important; align-items:center !important; gap:12px !important; margin:0 0 10px !important;
  padding:10px 14px !important; background:#fff; border:1px solid #e3ebf5; border-radius:16px;
  box-shadow:0 4px 14px rgba(16,36,63,.06); box-sizing:border-box; }
/* Zvisle marginy nadpisu by sa scitali s gap:12px riadku - tie nulujeme a
   medzeru necháme na gap. Lavy margin 6px vsak ostava: rovnaky ma z
   .nd-v-riadku aj "Stav operácie" a "SAP časy", takze vsetky nadpisy
   zacinaju na tej istej zvislici. */
body.${BODY_CLASS} #__pda_detail_rightcol__ > [data-nd-nadpis] { margin:0 0 0 6px !important; }
body.${BODY_CLASS} #__pda_detail_rightcol__ .pda-machine { order:2; margin-left:auto !important; gap:8px !important; }
body.${BODY_CLASS} #__pda_detail_rightcol__ .pda-machine .sapMLabel { display:none !important; }
body.${BODY_CLASS} #__pda_detail_rightcol__ .pda-machine::before { content:'Stroj'; font:700 13px/1 -apple-system,"Segoe UI",Roboto,sans-serif;
  color:#4a6285; margin-right:4px; }
body.${BODY_CLASS} #__pda_detail_rightcol__ #WorkcenterDetail--Confirm_Button { order:3; }
body.${BODY_CLASS} #__pda_order_drawing_wrapper__ > button { border:1px solid #dfe7f2 !important; border-radius:12px !important;
  padding:8px 14px !important; box-shadow:0 2px 8px rgba(16,36,63,.08) !important; }
/* Components/BOM: rovnako vysoke ako VYKRES. Nedavame pevnu vysku - v riadku
   sa len natiahne (align-self:stretch) na vysku najvyssieho prvku, a to je
   prave tlacidlo VYKRES (trojriadkove). Ak sa VYKRES zmeni, BOM ide s nim. */
/* Popri triede sa cielime aj cez ID: pri prekresleni po nacitani zakazky
   UI5 na chvilu prepise class a tlacidlo by inak blyslo v povodnej sirke,
   kym mu DomWatch triedu vrati. ID si UI5 nemeni. */
body.${BODY_CLASS} #__pda_detail_rightcol__ .nd-bom,
body.${BODY_CLASS} #__pda_detail_rightcol__ #WorkcenterDetail--BoM_Button {
  align-self:stretch !important; flex:0 0 220px !important;
  width:220px !important; min-width:0 !important;
  /* height:auto je tu nutnost, nie ozdoba: UI5 pise tlacidlu vysku priamo
     do inline stylu a align-self:stretch sa uplatni LEN vtedy, ked je
     vyska auto. S pevnou vyskou sa natiahnutie ticho ignoruje. */
  height:auto !important;
  padding:0 !important; box-sizing:border-box !important; }
/* Ram, tien a pozadie kresli vnutorny .sapMBtnInner, nie samotne tlacidlo -
   preto musi vyplnit cele tlacidlo, inak by ramcek obopinal len text. */
body.${BODY_CLASS} .nd-bom .sapMBtnInner,
body.${BODY_CLASS} #WorkcenterDetail--BoM_Button .sapMBtnInner { background:#fff !important; border:1px solid #dfe7f2 !important; border-radius:12px !important;
  padding:0 !important; box-shadow:0 2px 8px rgba(16,36,63,.08) !important; color:#13315c !important; font-weight:700 !important;
  height:100% !important; width:100% !important; box-sizing:border-box !important;
  display:flex !important; align-items:center !important; justify-content:center !important; }
/* Hover BOM: ram si nastavuje .sapMBtnInner hore skratkou "border", a ten
   zapis ma rovnaku specificitu ako povodne hover pravidlo vyssie v subore -
   pri zhode vyhrava neskorsi, takze hover prestal fungovat. Tu je preto
   silnejsi selektor (dve ID) a stoji az za nim. */
body.${BODY_CLASS} #__pda_detail_rightcol__ #WorkcenterDetail--BoM_Button:hover .sapMBtnInner,
body.${BODY_CLASS} #__pda_detail_rightcol__ .nd-bom:hover .sapMBtnInner {
  border-color:#7ba4ee !important; background:#f6f9ff !important; }
/* Operation Complete: rovnaka vyska ako BOM - opat cez natiahnutie v riadku,
   nie pevnou hodnotou. height:auto je nutne, inak UI5-acku inline vysku
   align-self:stretch ticho ignoruje (to iste ako pri BOM). */
body.${BODY_CLASS} #__pda_detail_rightcol__ #WorkcenterDetail--Confirm_Button {
  align-self:stretch !important; height:auto !important; box-sizing:border-box !important; }
/* Bledozelene pozadie s tmavozelenym textom; pri prechode mysou sa farby
   vymenia. Zelene odtiene su tie, ktore uz v skripte su (tag "vyrába"). */
body.${BODY_CLASS} #WorkcenterDetail--Confirm_Button .sapMBtnInner { background:#e7f6ec !important;
  background-image:none !important; border:1px solid #b6e2c5 !important; border-radius:12px !important;
  padding:10px 22px !important; box-shadow:0 2px 8px rgba(29,122,60,.12) !important;
  height:100% !important; box-sizing:border-box !important;
  display:flex !important; align-items:center !important; justify-content:center !important;
  transition:background-color .12s, color .12s !important; }
body.${BODY_CLASS} #WorkcenterDetail--Confirm_Button .sapMBtnContent,
body.${BODY_CLASS} #WorkcenterDetail--Confirm_Button bdi { color:#1d7a3c !important; font-weight:800 !important; font-size:14px !important; }
body.${BODY_CLASS} #WorkcenterDetail--Confirm_Button .sapMBtnInner::before { content:'✓'; color:#1d7a3c; font-weight:900; margin-right:8px; }
/* hover: prehodene - tmavozelene pozadie, bledozeleny text */
body.${BODY_CLASS} #WorkcenterDetail--Confirm_Button:hover .sapMBtnInner {
  background:#1d7a3c !important; border-color:#1d7a3c !important;
  box-shadow:0 4px 12px rgba(29,122,60,.30) !important; }
body.${BODY_CLASS} #WorkcenterDetail--Confirm_Button:hover .sapMBtnContent,
body.${BODY_CLASS} #WorkcenterDetail--Confirm_Button:hover bdi { color:#e7f6ec !important; }
body.${BODY_CLASS} #WorkcenterDetail--Confirm_Button:hover .sapMBtnInner::before { color:#e7f6ec; }

/* ---------- karta ZAKAZKA A MATERIAL: vsetkych 5 riadkov (3 zo SAP
   formulara + 2 vlastne) jednoducho "Popis: hodnota" v jednom riadku,
   bez stlpcov/mriezky - popis a hodnota su inline za sebou (popis tucny
   + dvojbodka cez ::after, hodnota normalnym rezom hned za nim), cely
   riadok sa prirodzene zalomi ako bezny text ak sa nezmesti na jeden
   riadok. Zive SAP prvky su presunute do .nd-riadok JS-om (kartaZakazky()),
   povodny prazdny Layout kontajner je skryty nizsie. ---------- */
body.${BODY_CLASS} #WorkcenterDetail--Main_SimpleForm .sapUiForm { background:#fff !important; border:1px solid #e3ebf5 !important;
  border-radius:16px !important; padding:14px 18px !important; box-shadow:0 4px 14px rgba(16,36,63,.06) !important;
  display:flex !important; flex-direction:column !important; gap:6px !important; }
body.${BODY_CLASS} #WorkcenterDetail--Main_SimpleForm--Layout { display:none !important; }
body.${BODY_CLASS} #WorkcenterDetail--Main_SimpleForm .sapUiForm .nd-riadok {
  display:block !important; width:100% !important; margin:0 !important; line-height:1.5 !important; }
/* font-size tu musi byt uvedena: prve tri riadky su zive SAP prvky s
   vlastnou velkostou 14px, kym nase vlastne .k/.v ju nemaju a zdedili by
   16px z karty - Mnozstvo a Pracovisko potom vychadzali vacsie */
body.${BODY_CLASS} #WorkcenterDetail--Main_SimpleForm .sapUiForm .nd-riadok > .sapMLabel,
body.${BODY_CLASS} #WorkcenterDetail--Main_SimpleForm .sapUiForm .nd-riadok > .k {
  display:inline !important; font-weight:700 !important; color:#4a6285 !important;
  font-size:14px !important;
  max-width:none !important; width:auto !important; margin:0 !important; }
body.${BODY_CLASS} #WorkcenterDetail--Main_SimpleForm .sapUiForm .nd-riadok .sapMLabelColonAndRequired { display:none !important; }
body.${BODY_CLASS} #WorkcenterDetail--Main_SimpleForm .sapUiForm .nd-riadok > .sapMLabel::after,
body.${BODY_CLASS} #WorkcenterDetail--Main_SimpleForm .sapUiForm .nd-riadok > .k::after {
  content:':' !important; }
body.${BODY_CLASS} #WorkcenterDetail--Main_SimpleForm .sapUiForm .nd-riadok > .sapMText,
body.${BODY_CLASS} #WorkcenterDetail--Main_SimpleForm .sapUiForm .nd-riadok > .v {
  display:inline !important; font-weight:400 !important; color:#13315c !important;
  font-size:14px !important;
  max-width:none !important; margin-left:6px !important; }

/* ---------- mriezka pod riadkom akcii: zakazka+material | SAP casy (hore),
   popis operacie | paralelne procesy (dole) - namiesto povodneho radenia
   pod sebou. Prvky sa presuvaju funkciou usporiadajDetailMriezku(), tu sa
   len rozmiestnuju cez grid-area podla ID (poradie v DOM je jedno). ---------- */
body.${BODY_CLASS} #WorkcenterDetail--OrderDetails_FlexBox.nd-detail-grid {
  display:grid !important; grid-template-columns:minmax(0,1fr) minmax(0,1fr) !important;
  grid-template-rows:auto auto !important;
  align-items:stretch !important; gap:12px !important;
  height:auto !important; width:100% !important; margin:0 0 10px !important; }
/* explicitne riadok/stlpec namiesto grid-template-areas - odolnejsie voci
   tomu, ci sa retazec s pomenovanymi oblastami spravne naparsuje */
body.${BODY_CLASS} #WorkcenterDetail--Main_SimpleForm { grid-column:1 !important; grid-row:1 !important;
  width:100% !important; max-width:none !important; height:auto !important; margin:0 !important; min-width:0 !important; }
body.${BODY_CLASS} #WorkcenterDetail--TimerCharts_FlexBox { grid-column:2 !important; grid-row:1 !important;
  width:100% !important; height:auto !important; margin:0 !important; min-width:0 !important;
  background:#fff !important; border:1px solid #e3ebf5 !important; border-radius:16px !important;
  padding:10px 14px !important; box-shadow:0 4px 14px rgba(16,36,63,.06) !important; box-sizing:border-box !important;
  align-items:center !important; }
body.${BODY_CLASS} #WorkcenterDetail--SimpleForm_FlexBox { grid-column:1 !important; grid-row:2 !important;
  width:100% !important; max-width:none !important; height:100% !important; margin:0 !important; min-width:0 !important; }
body.${BODY_CLASS} #WorkcenterDetail--SimpleForm_FlexBox #__pda_opis_button__ { height:100% !important;
  margin:0 !important; max-width:none !important; box-sizing:border-box !important; align-content:start !important; }
body.${BODY_CLASS} #WorkcenterDetail--Order_Info_Buttons_FlexBox { grid-column:2 !important; grid-row:2 !important;
  width:100% !important; height:100% !important; margin:0 !important; min-width:0 !important; }
body.${BODY_CLASS} #WorkcenterDetail--Order_Info_Buttons_FlexBox .sapMList { background:#fff !important;
  border:1px solid #e3ebf5 !important; border-radius:16px !important; box-shadow:0 4px 14px rgba(16,36,63,.06) !important;
  height:100% !important; box-sizing:border-box !important; overflow:hidden !important; }
/* rovnaky nadpis ako .nd-nadpis ("SAP casy"). Dolezity je line-height:
   SAP dava hlavicke zoznamu vysoky pevny riadok (3rem), takze aj po
   height:auto ostal nadpis opticky zapichnuty v prazdnom pase. */
/*
 * Nadpis "Paralelné procesy". ZAMERNE tu NEURCUJEME font-family (ani
 * skratku "font", ktora by rodinu prepisala tiez): ked sme mu vnutili
 * Segoe UI, vysiel vyrazne tucnejsi nez ostatne nadpisy. Segoe UI ma
 * totiz skutocny rez 800, kym font appky ho nema a rovnakych 800 v nom
 * prehliadac vykresli len ako bezny bold. Nadpisy tak pri rovnakom
 * zapise vyzerali rozdielne. Rodinu preto dedi ako vsetko ostatne.
 * Padding 18px zlava je zhodny s kartou "Popis operácie".
 */
body.${BODY_CLASS} #WorkcenterDetail--Order_Info_Buttons_FlexBox #WorkcenterDetail--OrderStatus_List-header,
body.${BODY_CLASS} #WorkcenterDetail--Order_Info_Buttons_FlexBox .sapMListHdrText {
  font-size:12px !important; font-weight:800 !important; line-height:1.3 !important;
  letter-spacing:.14em !important; text-transform:uppercase !important; color:#4a6285 !important;
  background:transparent !important; border:0 !important;
  padding:12px 18px 8px !important; height:auto !important; min-height:0 !important; }
/* zoznam paralelnych procesov: odsadenie okolo poloziek */
body.${BODY_CLASS} #WorkcenterDetail--OrderStatus_List-listUl { padding:12px !important; box-sizing:border-box !important; }
/* jedna polozka zoznamu: meno vlavo, cas + Zastavit vpravo - jeden riadok
   namiesto dvoch (predtym VBox: hore nazov+meno, dole cas+tlacidlo) */
body.${BODY_CLASS} #WorkcenterDetail--OrderStatus_List-listUl .sapMFlexBoxFit.sapMVBox {
  flex-direction:row !important; justify-content:space-between !important; align-items:center !important; }
/* stav a meno operatora vedla seba v jednom riadku (UI5 ich dava ako VBox
   pod seba) a zvisle na stred zelenej pilulky */
body.${BODY_CLASS} #WorkcenterDetail--OrderStatus_List-listUl [id^="WorkcenterDetail--Title_FlexBox-"] {
  flex:1 1 auto !important; min-width:0 !important;
  flex-direction:row !important; align-items:center !important;
  justify-content:flex-start !important; gap:14px !important; margin:0 !important; }
/* UI5 rozdava vlastne triedy sapUiTinyMargin* - v jednoriadkovom rozlozeni
   uz len rozhadzuju zvisle centrovanie, takze ich tu nulujeme */
body.${BODY_CLASS} #WorkcenterDetail--OrderStatus_List-listUl [id^="WorkcenterDetail--Title_FlexBox-"] > .sapMLabel,
body.${BODY_CLASS} #WorkcenterDetail--OrderStatus_List-listUl [id^="WorkcenterDetail--Timer_Label-"] {
  margin:0 !important; line-height:1.2 !important; }
body.${BODY_CLASS} #WorkcenterDetail--OrderStatus_List-listUl [id^="WorkcenterDetail--Timer_FlexBox-"] {
  flex:0 0 auto !important; margin-left:12px !important; gap:12px !important; }

/* ---------- SAP casy: ploche kolace s percentom v strede a legendou ---------- */
body.${BODY_CLASS} .pda-3d { transform:none !important; margin-top:6px !important;
  filter:drop-shadow(0 6px 10px rgba(16,36,63,.16)) !important; }
body.${BODY_CLASS} .pda-3d:hover { transform:none !important; }
.nd-pct { position:absolute; transform:translate(-50%,-50%); text-align:center; pointer-events:none;
  font-family:-apple-system,"Segoe UI",Roboto,sans-serif; }
.nd-pct .c { display:block; font-size:22px; font-weight:800; color:#13315c; line-height:1; }
.nd-pct .h { display:block; font-size:11px; color:#4a6285; margin-top:3px; }
.nd-legenda { display:flex; gap:14px; justify-content:center; margin-top:8px;
  font:12px/1.2 -apple-system,"Segoe UI",Roboto,sans-serif; color:#4a6285; }
.nd-legenda i { display:inline-block; width:9px; height:9px; border-radius:50%; margin-right:5px; vertical-align:middle; }
.nd-legenda .a i { background:#2b7fe0; }
.nd-legenda .b i { background:#c9d4e2; }

/* ---------- SAP casy v karte: kompaktne kolace vedla seba (nadpis, potom
   maly kolac s percentom, cas pod nim), bez legendy - setri miesto v mriezke ---------- */
body.${BODY_CLASS} #WorkcenterDetail--TimerCharts_FlexBox .sapMVBox { flex:1 1 0 !important; min-width:0 !important;
  align-items:center !important; }
body.${BODY_CLASS} #WorkcenterDetail--TimerCharts_FlexBox .sapMLabel { order:1 !important; font-size:10.5px !important;
  letter-spacing:.08em !important; text-transform:uppercase !important; color:#4a6285 !important; font-weight:800 !important;
  margin:0 0 2px !important; }
body.${BODY_CLASS} #WorkcenterDetail--TimerCharts_FlexBox canvas.pda-3d { order:2 !important;
  width:120px !important; height:120px !important; margin:2px auto !important; }
body.${BODY_CLASS} #WorkcenterDetail--TimerCharts_FlexBox .sapMText { order:3 !important;
  font-size:11px !important; color:#6b7c95 !important; margin:2px 0 0 !important; }
body.${BODY_CLASS} #WorkcenterDetail--TimerCharts_FlexBox .nd-legenda { display:none !important; }
body.${BODY_CLASS} #WorkcenterDetail--TimerCharts_FlexBox .nd-pct { transform:translate(-50%, -20%) !important; }
body.${BODY_CLASS} #WorkcenterDetail--TimerCharts_FlexBox .nd-pct .c { font-size:16px !important; }
body.${BODY_CLASS} #WorkcenterDetail--TimerCharts_FlexBox .nd-pct .h { display:none !important; }

/* ---------- karta "Prehľad zdrojov" v lavom stlpci ----------
   V boxe ostavaju len dve tlacidla. Nadpis "Resource over time", maly graf aj
   riadok s prazdnymi placeholdermi skryvame - graf sa aj tak otvara az v
   dialogu cez tlacidlo s lupou. Tlacidla su v appke bez textu (len ikona),
   popisky preto doplname cez ::after. */
body.${BODY_CLASS} #__pda_left_box_graf__ #WorkcenterDetail--WorkplaceDetailsButtons_FlexBox,
body.${BODY_CLASS} #__pda_left_box_graf__ #WorkcenterDetail--Over_Title_Title,
body.${BODY_CLASS} #__pda_left_box_graf__ #WorkcenterDetail--ChartFlexBox { display:none !important; }
body.${BODY_CLASS} #__pda_left_box_graf__ #WorkcenterDetail--Over_Time_FlexBox,
body.${BODY_CLASS} #__pda_left_box_graf__ #WorkcenterDetail--Over_Time_Header {
  width:100% !important; height:auto !important; margin:0 !important; padding:0 !important; }
/* riadok s tlacidlami: dve rovnake polovice cez celu sirku */
body.${BODY_CLASS} #__pda_left_box_graf__ #WorkcenterDetail--Buttons_FlexBox {
  width:100% !important; height:auto !important; margin:0 !important;
  justify-content:stretch !important; gap:10px !important; }
body.${BODY_CLASS} #__pda_left_box_graf__ #WorkcenterDetail--Buttons_FlexBox > .sapMBtn {
  flex:1 1 0 !important; width:auto !important; min-width:0 !important;
  height:auto !important; margin:0 !important; }
body.${BODY_CLASS} #__pda_left_box_graf__ #WorkcenterDetail--Buttons_FlexBox .sapMBtnInner {
  display:flex !important; align-items:center !important; justify-content:center !important;
  gap:8px !important; width:100% !important; height:auto !important;
  padding:0 12px !important; box-sizing:border-box !important; line-height:26px !important;
  background:#fff !important; background-image:none !important;
  border:1px solid #dfe7f2 !important; border-radius:12px !important;
  box-shadow:0 2px 8px rgba(16,36,63,.08) !important;
  font-size:13px !important; font-weight:700 !important; color:#13315c !important; }
body.${BODY_CLASS} #__pda_left_box_graf__ #WorkcenterDetail--Buttons_FlexBox .sapMBtn:hover .sapMBtnInner {
  border-color:#7ba4ee !important; background:#f6f9ff !important; }
body.${BODY_CLASS} #__pda_left_box_graf__ #WorkcenterDetail--Buttons_FlexBox .sapUiIcon {
  color:#13315c !important; font-size:16px !important; margin:0 !important; }
body.${BODY_CLASS} #WorkcenterDetail--Chart_Button .sapMBtnInner::after { content:'Graf'; }
body.${BODY_CLASS} #WorkcenterDetail--Status_Overview_Button .sapMBtnInner::after { content:'Tabuľka'; }

/* ---------- POPIS OPERACIE ako karta s "Cely text" ---------- */
body.${BODY_CLASS} #__pda_opis_button__ { display:grid !important; grid-template-columns:1fr auto !important;
  grid-template-rows:auto auto !important; column-gap:12px !important; row-gap:6px !important;
  background:#fff !important; border:1px solid #e3ebf5 !important; border-radius:16px !important;
  padding:12px 18px !important; box-shadow:0 4px 14px rgba(16,36,63,.06) !important; }
body.${BODY_CLASS} #__pda_opis_button__ .ikona { display:none !important; }
body.${BODY_CLASS} #__pda_opis_button__ .stred { display:contents !important; }
body.${BODY_CLASS} #__pda_opis_button__ .nadpis { grid-column:1; grid-row:1; font-size:12px !important; letter-spacing:.14em !important;
  color:#4a6285 !important; }
body.${BODY_CLASS} #__pda_opis_button__ .sipka { grid-column:2; grid-row:1; font-size:13px !important; }
body.${BODY_CLASS} #__pda_opis_button__ .ukazka { grid-column:1 / span 2; grid-row:2; white-space:normal !important;
  display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden;
  font-size:14px !important; line-height:1.5 !important; color:#17202e !important; }

/* ---------- pravy panel PDA ---------- */
body.${BODY_CLASS} #__pda_hf_menu__ { border-radius:18px !important; padding:12px !important; }
body.${BODY_CLASS} #__pda_hf_menu__ .hf-logo { display:none !important; }
body.${BODY_CLASS} #__pda_hf_menu__ .hf-nadpis { text-align:left; font-size:12px; letter-spacing:.14em; color:#4a6285; margin:2px 0 6px 2px; }
body.${BODY_CLASS} #__pda_hf_menu__ .hf-btn { border:1px solid #e3ebf5 !important; background:#f6f9fd !important;
  box-shadow:0 2px 6px rgba(16,36,63,.06) !important; padding:10px 12px !important; }
body.${BODY_CLASS} #__pda_hf_menu__ .hf-btn .ik { width:38px !important; height:38px; border-radius:10px; background:#e6f0fb;
  color:#1e6fd9; display:flex; align-items:center; justify-content:center; font-size:18px; }
body.${BODY_CLASS} #__pda_hf_menu__ .hf-btn .ik.foto { background:transparent; }
body.${BODY_CLASS} #__pda_hf_menu__ .hf-btn .n { font-size:14px !important; }
.nd-hf-blok { display:flex; align-items:center; gap:12px; margin-top:auto; padding:14px 6px 4px; }
.nd-hf-blok img { width:54px; height:54px; border-radius:12px; display:block; }
.nd-hf-blok .t { font:800 12px/1.35 -apple-system,"Segoe UI",Roboto,sans-serif; letter-spacing:.06em; color:#13315c; text-transform:uppercase; }
.nd-slogan { font:800 12px/1.4 -apple-system,"Segoe UI",Roboto,sans-serif; letter-spacing:.08em; color:#9fb0c8;
  text-transform:uppercase; padding:8px 6px 2px; }

/* ---------- lavy zoznam: pocet operacii v hlavicke ---------- */
.nd-pocet { font:600 12px/1 -apple-system,"Segoe UI",Roboto,sans-serif; color:#4a6285; margin-left:10px; white-space:nowrap; }
`;
            document.head.appendChild(st);
        }

        /* --- horny pruh: logo a nazov vedla povodnych ovladacich prvkov --- */

        /*
         * Cely horny pruh sa najde tak, ze sa od tlacidla Odhlasenie ide nahor,
         * kym prvok nie je siroky aspon na 80 % okna - to je pruh cez celu sirku
         * (nazvy tried UI5 sa nedaju spolahlivo uhadnut, sirka je istejsia).
         */
        function hornyPruh() {
            const odhlasenie = document.querySelector('[id$="Button_Logout"]');
            if (!odhlasenie) return;
            let pruh = odhlasenie.parentElement;
            for (let i = 0; i < 8 && pruh; i++) {
                if (pruh.getBoundingClientRect().width >= W.innerWidth * 0.8) break;
                pruh = pruh.parentElement;
            }
            if (!pruh || pruh === document.body) return;
            if (!pruh.classList.contains('nd-topbar')) pruh.classList.add('nd-topbar');
        }

        /* --- nadpisy sekcii --- */

        /*
         * Nadpis ide DOVNUTRA kontajnera ako jeho prvy riadok (flex-basis 100 %),
         * nie ako sused pred neho. Sused sa v 2.1.2 vkladal dokola: modul
         * hlavicky presuva riadok tlacidiel stale na prve miesto, cim sa nadpis
         * dostal za neho, kontrola "je pred nim nadpis?" zlyhala a vlozil sa
         * dalsi - a tak stale, pri kazdom tiku. Vnutri kontajnera sa presun
         * kontajnera nadpisu nedotkne.
         */
        /*
         * Trom tlacidlam osobneho stavu priradi obrazok do pozadia. Priradzuje
         * sa podla TEXTU, nie podla poradia: ID tychto tlacidiel koncia
         * indexom (…-0, -1, -2) a poradie tlacidiel meni ina cast suity,
         * takze podla ID by sa obrazky po preusporiadani prehodili.
         * Diakritiku zo stringu odstranujeme, nech "čakanie" sedi aj ked ho
         * appka napise inak.
         */
        const OBRAZKY_STAVOV = [
            // osobny stav (uvodna obrazovka)
            [/^stretnut/,      'stretnutia'],
            [/^prestav/,       'prestavka'],
            [/^cakan/,         'cakanie'],
            // stav operacie (detail pracoviska)
            [/^vyroba/,        'vyroba'],
            [/^upinanie/,      'upinanie'],
            [/^upratovanie/,   'upratovanie'],
            [/^meranie/,       'meranie'],
            [/^programovanie/, 'programovanie'],
            [/^chyba/,         'chyba'],
            [/^udrzb/,         'udrzba'],
        ];

        function obrazkyStavov() {
            document.querySelectorAll('.statusBtn').forEach((b) => {
                const t = (b.textContent || '').trim().toLowerCase()
                    .normalize('NFD').replace(/[̀-ͯ]/g, '');
                const najdene = OBRAZKY_STAVOV.find((r) => r[0].test(t));
                if (!najdene) return;
                if (b.dataset.pdaObraz !== najdene[1]) b.dataset.pdaObraz = najdene[1];
            });
        }

        function nadpisDo(kontajner, text, znacka) {
            if (!kontajner) return;
            let d = null;
            for (const ch of kontajner.children) {
                if (ch.dataset && ch.dataset.ndNadpis === znacka) { d = ch; break; }
            }
            if (!d) {
                d = document.createElement('div');
                d.className = 'nd-nadpis nd-v-riadku';
                d.dataset.ndNadpis = znacka;
                d.textContent = text;
            }
            if (kontajner.firstElementChild !== d) kontajner.insertBefore(d, kontajner.firstElementChild);
        }

        /* --- patka --- */

        function patka() {
            if (document.getElementById(FOOTER_ID)) return;
            const f = document.createElement('div');
            f.id = FOOTER_ID;
            const img = document.createElement('img');
            img.src = LOGO_HF_MALE;
            img.alt = '';
            const n = document.createElement('span'); n.className = 'nd-f-n'; n.textContent = 'HF SLOVAKIA';
            const s = document.createElement('span'); s.textContent = 'Better Parts. A Cleaner Tomorrow.';
            const r = document.createElement('span'); r.className = 'nd-f-r';
            let verzia = '';
            try { verzia = GM_info && GM_info.script ? 'v' + GM_info.script.version : ''; } catch (e) { /* ignore */ }
            r.textContent = ('PDA App Extension ' + verzia).trim() + ' · Powered by HF Slovakia';
            f.appendChild(img); f.appendChild(n); f.appendChild(s); f.appendChild(r);
            document.body.appendChild(f);
        }

        /* --- vybrana operacia z premennych appky (len citanie) --- */

        function vybranaOperacia() {
            try {
                const main = W.sap.ui.getCore().byId('Main');
                const op = main && main.getController().getGlobals().getVar('oSelectedWorkcenterOperation');
                return op && op.productionOrderNo ? op : null;
            } catch (e) { return null; }
        }

        /*
         * Riadok akcii pod stavovymi tlacidlami: VYKRES | Components/BOM |
         * Stroj ON/OFF | Operation Complete. Stlpec s tymito prvkami uz drzi
         * modul hlavicky (#__pda_detail_rightcol__); tu sa len presunie na
         * zaciatok hlavicky (za nadpis) a CSS z neho spravi vodorovny riadok.
         * Tlacidlo Components / BOM appky sa don presunie z praveho stlpca -
         * je to ten isty prvok, len na inom mieste, jeho funkcia sa nemeni.
         */
        function riadokAkcii() {
            const col = document.getElementById('__pda_detail_rightcol__');
            const header = document.getElementById('WorkcenterDetail--OrderHeader_FlexBox');
            if (!col || !header || col.parentElement !== header) return;

            const prvy = header.firstElementChild;
            const nadpisJe = prvy && prvy.dataset && prvy.dataset.ndNadpis;
            const ciel = nadpisJe ? prvy.nextElementSibling : prvy;
            if (ciel !== col) header.insertBefore(col, ciel);

            document.querySelectorAll('.sapMBtn').forEach((b) => {
                if (b.closest('.sapMDialog')) return;
                if (!/^components/i.test((b.textContent || '').trim())) return;
                /*
                 * Triedu dopisujeme pri KAZDOM behu, nie len pri presune.
                 * UI5 si po nacitani zakazky tlacidlo prekresli a pritom
                 * prepise cely atribut class podla vlastneho zoznamu tried -
                 * nasa nd-bom z neho vypadne, hoci tlacidlo v riadku zostane.
                 * Predtym tu bola podmienka "ak uz je v riadku, preskoc", cize
                 * sa trieda nikdy nevratila a tlacidlu spadla sirka aj vyska.
                 */
                if (!b.classList.contains('nd-bom')) b.classList.add('nd-bom');
                if (col.contains(b)) return;   // presuva sa len raz, uz je na mieste
                const kotva = col.querySelector('.pda-machine');
                col.insertBefore(b, kotva || null);
            });
        }

        /*
         * Karta ZAKAZKA A MATERIAL: vsetkych 5 riadkov (3 zo SAP UI5
         * ResponsiveGrid formulara - Zakaznicka zakazka/Material/Production
         * Order + 2 vlastne - Mnozstvo/Pracovisko z vybranej operacie) sa
         * zobrazuje jednotne, jednoducho "Popis: hodnota" v jednom riadku
         * (bez stlpcov/mriezky). Zive SAP prvky (label + span s hodnotou)
         * sa len presuvaju do spolocneho .nd-riadok wrappera (appka ich
         * stale viaze), povodny Main_SimpleForm--Layout kontajner ostane v
         * DOM prazdny a je skryty cez CSS (display:none).
         */
        function kartaZakazky() {
            const form = document.querySelector('#WorkcenterDetail--Main_SimpleForm .sapUiForm');
            if (!form) return;

            // riadok zo zivych SAP prvkov (label + hodnota sa presunu do
            // spolocneho .nd-riadok wrappera, id drzi wrapper stabilny naprac apply()).
            function riadokZoSap(riadokId, label, value) {
                if (!label || !value) { const stary = document.getElementById(riadokId); if (stary) stary.remove(); return; }
                let r = document.getElementById(riadokId);
                if (!r) { r = document.createElement('div'); r.id = riadokId; r.className = 'nd-riadok'; }
                if (label.parentElement !== r) r.appendChild(label);
                if (value.parentElement !== r) r.appendChild(value);
                form.appendChild(r);
            }
            riadokZoSap('__pda_r_sales__', document.getElementById('WorkcenterDetail--SalesOrder_Label'),
                        document.getElementById('WorkcenterDetail--SalesOrder_Text'));
            riadokZoSap('__pda_r_material__', document.getElementById('WorkcenterDetail--Material_Label'),
                        document.getElementById('WorkcenterDetail--Material_Text'));
            riadokZoSap('__pda_r_prodorder__', document.getElementById('WorkcenterDetail--ProdOrder_Label'),
                        document.getElementById('WorkcenterDetail--ProdOrder_Text'));

            const op = vybranaOperacia();
            function riadokVlastny(riadokId, k, v) {
                if (!v) { const stary = document.getElementById(riadokId); if (stary) stary.remove(); return; }
                let r = document.getElementById(riadokId);
                if (!r) {
                    r = document.createElement('div'); r.id = riadokId; r.className = 'nd-riadok';
                    const kk = document.createElement('span'); kk.className = 'k';
                    const vv = document.createElement('span'); vv.className = 'v';
                    r.appendChild(kk); r.appendChild(vv);
                }
                const kk = r.querySelector('.k'), vv = r.querySelector('.v');
                if (kk.textContent !== k) kk.textContent = k;
                if (vv.textContent !== v) vv.textContent = v;
                form.appendChild(r);
            }
            if (!op) {
                riadokVlastny('__pda_r_mnozstvo__', 'Množstvo', '');
                riadokVlastny('__pda_r_pracovisko__', 'Pracovisko', '');
                return;
            }
            const mnozstvo = (op.targetQuantity !== undefined && op.targetQuantity !== null && String(op.targetQuantity) !== '')
                ? String(op.targetQuantity) + ' ks' : '';
            riadokVlastny('__pda_r_mnozstvo__', 'Množstvo', mnozstvo);
            const prac = [op.workcenter, op.workcenterDescription].filter(Boolean).join(' - ');
            riadokVlastny('__pda_r_pracovisko__', 'Pracovisko', prac);
        }

        /*
         * SAP casy: percento "hotovo" v strede kolaca + legenda pod nim.
         * Pocita sa z textu casov appky ("00:00:00 / 01:40:00" alebo "3144 / 3144 min"),
         * kolac sa nekresli odznova - nas popis lezi nad nim.
         */
        const KOTVY_CASOV = ['WorkcenterDetail--SetupTime_Text', 'WorkcenterDetail--MachineTime_Text',
                             'WorkcenterDetail--LaborTime_Text'];

        function sekundy(t) {
            t = String(t || '').trim();
            if (/:/.test(t)) return t.split(':').map(Number).reduce((a, b) => a * 60 + (isNaN(b) ? 0 : b), 0);
            const n = parseFloat(t.replace(',', '.'));
            return isNaN(n) ? 0 : n;
        }

        function percentaKolacov() {
            KOTVY_CASOV.forEach((id) => {
                const text = document.getElementById(id);
                if (!text) return;
                const casti = (text.textContent || '').split('/');
                if (casti.length < 2) return;
                const hotovo = sekundy(casti[0]);
                const plan = sekundy(casti[1]);
                const pct = plan > 0 ? Math.round(hotovo / plan * 100) : 0;
                const zostava = Math.max(0, 100 - pct);

                const box = text.closest('.sapMVBox') || text.closest('.sapMFlexBox') || text.parentElement;
                if (!box) return;
                let canvas = null;
                box.querySelectorAll('canvas, svg').forEach((g) => {
                    if (!canvas && g.id !== 'ResourceDetails' && g.id !== 'DialogChart' &&
                        g.getBoundingClientRect().width >= 60) canvas = g;
                });
                if (!canvas || !canvas.parentElement) return;

                const par = canvas.parentElement;
                if (par.style.position !== 'relative') par.style.position = 'relative';
                let lbl = null;
                for (const ch of par.children) { if (ch.classList.contains('nd-pct')) { lbl = ch; break; } }
                if (!lbl) {
                    lbl = document.createElement('div'); lbl.className = 'nd-pct';
                    const c = document.createElement('span'); c.className = 'c';
                    const h = document.createElement('span'); h.className = 'h'; h.textContent = 'hotovo';
                    lbl.appendChild(c); lbl.appendChild(h);
                    par.appendChild(lbl);
                }
                const c = lbl.querySelector('.c');
                const pctTxt = pct + '%';
                if (c.textContent !== pctTxt) c.textContent = pctTxt;
                const lx = Math.round(canvas.offsetLeft + canvas.offsetWidth / 2) + 'px';
                const ly = Math.round(canvas.offsetTop + canvas.offsetHeight / 2) + 'px';
                if (lbl.style.left !== lx) lbl.style.left = lx;
                if (lbl.style.top !== ly) lbl.style.top = ly;

                let leg = null;
                for (const ch of box.children) { if (ch.classList.contains('nd-legenda')) { leg = ch; break; } }
                if (!leg) {
                    leg = document.createElement('div'); leg.className = 'nd-legenda';
                    const a = document.createElement('span'); a.className = 'a';
                    const b = document.createElement('span'); b.className = 'b';
                    a.appendChild(document.createElement('i')); a.appendChild(document.createTextNode(''));
                    b.appendChild(document.createElement('i')); b.appendChild(document.createTextNode(''));
                    leg.appendChild(a); leg.appendChild(b);
                    box.appendChild(leg);
                }
                const ta = 'Hotovo ' + pct + '%', tb = 'Zostáva ' + zostava + '%';
                const na = leg.querySelector('.a').lastChild, nb = leg.querySelector('.b').lastChild;
                if (na.textContent !== ta) na.textContent = ta;
                if (nb.textContent !== tb) nb.textContent = tb;
            });
        }

        /* --- POPIS OPERACIE: odkaz "Cely text" namiesto "otvorit" --- */
        function popisKarta() {
            const sipka = document.querySelector('#__pda_opis_button__ .sipka');
            if (sipka && sipka.textContent !== 'Celý text ›') sipka.textContent = 'Celý text ›';
        }

        /*
         * Karty pod riadkom akcii usporiadane do mriezky 2x2 namiesto pod
         * sebou: zakazka a material | SAP casy (hore), popis operacie |
         * paralelne procesy (dole). CSS grid-column/grid-row funguje len na
         * PRIAMYCH potomkoch #WorkcenterDetail--OrderDetails_FlexBox, preto
         * sem musia byt presunute vsetky 4 karty - vratane TimerCharts_FlexBox,
         * ktore povodne nie je jeho priamym dietaom, ale je zanorene o uroven
         * hlbsie vnutri SimpleForm_FlexBox (za tlacidlom Popis operacie).
         * Main_SimpleForm bol zas vnutri OrderHeader_FlexBox (za riadkom
         * akcii, ktoreho sa toto nedotyka) a Order_Info_Buttons_FlexBox bol
         * az za OrderDetails_FlexBox ako jeho vlastny súrodenec. Prvky sa len
         * presuvaju (su zive, appka ich stale ovlada), viditelne miesto v
         * mriezke urcuje CSS cez grid-column/grid-row podla ID, takze na
         * poradi v DOM uz nezalezi.
         */
        function usporiadajDetailMriezku() {
            const grid = document.getElementById('WorkcenterDetail--OrderDetails_FlexBox');
            const zakazka = document.getElementById('WorkcenterDetail--Main_SimpleForm');
            const casy = document.getElementById('WorkcenterDetail--TimerCharts_FlexBox');
            const paralelne = document.getElementById('WorkcenterDetail--Order_Info_Buttons_FlexBox');
            if (!grid || !zakazka || !casy || !paralelne) return;
            if (!grid.classList.contains('nd-detail-grid')) grid.classList.add('nd-detail-grid');
            if (zakazka.parentElement !== grid) grid.appendChild(zakazka);
            if (casy.parentElement !== grid) grid.appendChild(casy);
            if (paralelne.parentElement !== grid) grid.appendChild(paralelne);
        }

        /* --- hlavicka zoznamu "Parallel Process Handling" -> "Paralelne procesy" --- */
        function prekladHlavicky() {
            const h = document.getElementById('WorkcenterDetail--OrderStatus_List-header');
            if (h && /parallel process/i.test(h.textContent || '')) h.textContent = 'Paralelné procesy';
        }

        /* --- lavy zoznam: pocet operacii vedla nadpisu "Pracovny zoznam" --- */
        function pocetZoznamu() {
            const box = document.getElementById('__pda_left_box_zoznam__') ||
                        document.getElementById('WorkcenterDetail--LeftColumn_FlexBox');
            if (!box) return;
            let title = null;
            box.querySelectorAll('.sapMTitle, .sapMLabel, .sapMText').forEach((el) => {
                if (title) return;
                const t = (el.textContent || '').toLowerCase();
                if (t.indexOf('pracovn') !== -1 && t.indexOf('zoznam') !== -1 && !el.querySelector('.nd-pocet')) title = el;
            });
            if (!title) return;
            let pocet = document.getElementById('__pda_nd_pocet__');
            if (!pocet) {
                pocet = document.createElement('span');
                pocet.id = '__pda_nd_pocet__';
                pocet.className = 'nd-pocet';
            }
            if (pocet.previousElementSibling !== title) title.insertAdjacentElement('afterend', pocet);
            const n = document.querySelectorAll('#WorkcenterDetail--Work_List .sapMLIB').length;
            const txt = n + (n === 1 ? ' operácia' : (n >= 2 && n <= 4 ? ' operácie' : ' operácií'));
            if (pocet.textContent !== txt) pocet.textContent = txt;
        }

        /* --- pravy panel PDA: blok HF "System integrovany pre vyrobu" a slogan dole --- */
        function pravyPanelPDA() {
            const panel = document.getElementById('__pda_hf_menu__');
            if (!panel || panel.querySelector('.nd-hf-blok')) return;
            const blok = document.createElement('div');
            blok.className = 'nd-hf-blok';
            const img = document.createElement('img'); img.src = LOGO_HF_MALE; img.alt = 'HF';
            const t = document.createElement('div'); t.className = 't'; t.textContent = 'Systém integrovaný pre výrobu';
            blok.appendChild(img); blok.appendChild(t);
            const slogan = document.createElement('div');
            slogan.className = 'nd-slogan';
            slogan.textContent = 'Smart manufacturing. Real results.';
            panel.appendChild(blok);
            panel.appendChild(slogan);
        }

        function apply() {
            injectStyles();
            if (!document.body.classList.contains(BODY_CLASS)) document.body.classList.add(BODY_CLASS);

            hornyPruh();
            patka();

            nadpisDo(document.getElementById('WorkcenterDetail--Order_Status_Flexbox'), 'Stav operácie', 'stav');
            // Nadpis "Zákazka a materiál" ide priamo do karty (Main_SimpleForm--Form),
            // rovnako ako "SAP časy" ide do TimerCharts_FlexBox - predtym isiel do
            // OrderHeader_FlexBox, co uz nie je ta ista vizualna karta (Main_SimpleForm
            // sa z neho medzicasom presunul do mriezky), takze nadpis visel mimo karty.
            // Stary osirely nadpis (ak tam este ostal zo starsej verzie) sa odstrani.
            const orderHeaderZvysok = document.getElementById('WorkcenterDetail--OrderHeader_FlexBox');
            if (orderHeaderZvysok) {
                const staryNadpis = orderHeaderZvysok.querySelector('[data-nd-nadpis="zakazka"]');
                if (staryNadpis) staryNadpis.remove();
            }
            nadpisDo(document.getElementById('WorkcenterDetail--Main_SimpleForm--Form'), 'Zákazka a materiál', 'zakazka');
            nadpisDo(document.getElementById('WorkcenterDetail--TimerCharts_FlexBox'), 'SAP časy', 'casy');
            nadpisDo(document.getElementById('__pda_detail_rightcol__'), 'Dokumentácia', 'dokumentacia');
            nadpisDo(document.getElementById('__pda_left_box_graf__'), 'Prehľad zdrojov', 'zdroje');

            // kazda cast zvlast v try/catch - chyba v jednej nesmie zhodit ostatne
            [usporiadajDetailMriezku, riadokAkcii, kartaZakazky, percentaKolacov, popisKarta, prekladHlavicky,
             pocetZoznamu, pravyPanelPDA, obrazkyStavov].forEach((f) => {
                try { f(); } catch (e) { console.warn(LOG, 'novy dizajn:', f.name, e); }
            });
        }

        DomWatch.add(apply);
        onReady(apply);
    }

    /* ---------- 3.20 VERZIA 3J (Jaro) - uprava dizajnu ---------- */

    /*
     * Verzia 3J vznikla z produkcneho buildu 2.3.0 (Daniel Gabris, 2026-09-23)
     * a zije mimo gitu v C:\Claude Code\PDA verzia 3J\. Vsetky funkcie ostavaju
     * presne tie iste - tento modul meni VYLUCNE vzhlad (CSS) a par textov
     * v nasich vlastnych prvkoch. Ziadny prvok appky sa nemaze, nepresuva
     * ani nestlaca; data a cisla su tie, ktore ukazuje appka.
     *
     * 3.0.0:
     *   - SAP CASY: kazdy graf vo vlastnej oramovanej karte, nadpis a cas nad
     *     kolacom, vacsi kolac s percentom a slovom "hotovo" v strede, pod nim
     *     legenda Hotovo / Zostava; karty sa roztahuju vedla seba podla sirky
     *   - PARALELNE PROCESY: kazda polozka ako biela karta, pred textom ikona
     *     (zelena dlazdica so sipkami), stav tucne a meno pod nim, velky cas,
     *     Zastavit ako cervene tlacidlo
     *   - mimo kariet je stranka priehladna - vidno pozadie HF Slovakia
     *
     * Styl sa vklada az po module noveho dizajnu a selektory maju navyse
     * `html`, takze pri rovnakom !important vzdy vyhraju.
     */
    function modV3J() {
        const STYLE_ID = '__pda_3j_styles__';
        const B = 'html body.pda-nd';
        const CASY = '#WorkcenterDetail--TimerCharts_FlexBox';
        const ZOZNAM = '#WorkcenterDetail--OrderStatus_List-listUl';
        const SEKCIA = '#WorkcenterDetail--Order_Info_Buttons_FlexBox';

        // ikona pred paralelnym procesom: dve protismerne sipky (biele), vlozene ako SVG
        const IKONA_PROCESU = "url(\"data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' " +
            "fill='none' stroke='white' stroke-width='2.3' stroke-linecap='round' stroke-linejoin='round'>" +
            "<path d='M4 8.5h14.5M15 5l3.5 3.5L15 12'/><path d='M20 15.5H5.5M9 12l-3.5 3.5L9 19'/></svg>\")";

        function injectStyles() {
            if (document.getElementById(STYLE_ID)) return;
            const st = document.createElement('style');
            st.id = STYLE_ID;
            st.textContent = `
/* ================= paleta 3J: svetlomodre sklo namiesto bielej =================
   Plochy nie su biele, ale svetlomodre a polopriehladne (ako v navrhu) - pozadie
   cez ne jemne presvita. --nd-karta su velke karty, --nd-vnutro polozky v nich. */
html body.pda-nd { --nd-karta:rgba(231,237,246,.76); --nd-vnutro:rgba(244,247,251,.9);
  --nd-zebra:rgba(233,238,246,.92); --nd-okraj:rgba(255,255,255,.82);
  --nd-tien:0 8px 26px rgba(20,60,120,.10); }

${B} .pda-left-box, ${B} #__pda_hf_menu__, ${B} #__pda_detail_rightcol__,
${B} #WorkcenterDetail--Order_Status_Flexbox, ${B} #__pda_opis_button__,
${B} #WorkcenterDetail--Main_SimpleForm .sapUiForm, ${B} #WorkcenterDetail--Main_SimpleForm--Form {
  background:var(--nd-karta) !important; border:1px solid var(--nd-okraj) !important;
  box-shadow:var(--nd-tien) !important; -webkit-backdrop-filter:blur(10px); backdrop-filter:blur(10px); }
${B} #WorkcenterDetail--Order_Status_Flexbox { border-radius:18px !important; }
/* vseobecne panely (napr. na inych obrazovkach) - nie biele, ale svetlomodre */
${B} .sapMPanel { background:rgba(229,238,250,.6) !important; border-color:var(--nd-okraj) !important; }
/* horny pruh a patka priehladne */
${B} .nd-topbar { background:rgba(229,238,250,.55) !important; border-bottom:1px solid var(--nd-okraj) !important;
  -webkit-backdrop-filter:blur(10px); backdrop-filter:blur(10px); }
${B} #__pda_nd_footer__ { background:rgba(229,238,250,.6) !important; border-top:1px solid var(--nd-okraj) !important; }
/* pilulky v pracovnom zozname */
${B} #WorkcenterDetail--Work_List .sapMLIB.pda-pill-on { background:var(--nd-vnutro) !important;
  border-color:rgba(190,208,232,.9) !important; }
${B} #WorkcenterDetail--Work_List .sapMLIB.pda-pill-on:nth-child(even) { background:var(--nd-zebra) !important; }
/* vybrany zaznam: tmavomodry. Pravidla palety vyssie maju vyssiu specificitu nez
   povodne pravidlo pre vyber, takze ho prebili a vybrany riadok ostal svetly -
   preto tu vyber opakujeme s este vyssou specificitou. */
${B} #WorkcenterDetail--Work_List .sapMLIB.pda-pill-on.sapMLIBSelected,
${B} #WorkcenterDetail--Work_List .sapMLIB.pda-pill-on.sapMLIBSelected:nth-child(even) {
  background:#13315c !important; border-color:#0b2447 !important;
  box-shadow:0 4px 14px rgba(19,49,92,.35) !important; }
/* polozky praveho panela PDA */
${B} #__pda_hf_menu__ .hf-btn { background:var(--nd-vnutro) !important; border-color:var(--nd-okraj) !important; }

/* ================= hlavicky vsetkych kariet ako v navrhu =================
   Vacsie (16 px), polotucne, velkymi pismenami, len jemne rozpalcovane.
   Vsetky maju to iste pismo (Segoe UI) - ked malo pismo len jedna z nich,
   vyzerala tucnejsie nez ostatne (Segoe ma skutocny tucny rez, pismo appky nie). */
${B} .nd-nadpis, ${B} #WorkcenterDetail--OrderStatus_List-header,
${B} #WorkcenterDetail--Order_Info_Buttons_FlexBox .sapMListHdrText,
${B} #__pda_hf_menu__ .hf-nadpis, ${B} #__pda_opis_button__ .nadpis,
${B} .pda-left-box .sapMPanelHdr .sapMTitle, ${B} .pda-left-box .sapMPanelHeaderTB .sapMTitle {
  font-family:"Segoe UI",-apple-system,Roboto,sans-serif !important; font-size:16px !important;
  font-weight:600 !important; letter-spacing:.03em !important; text-transform:uppercase !important;
  color:#2d4b73 !important; line-height:1.3 !important; }

/* ================= hlavicka ako podfarbeny pas (ako v navrhu) =================
   Hlavicka nie je len text nad obsahom, ale pas cez celu sirku karty - jemne
   modrejsi nez karta a od obsahu oddeleny tenkou ciarou. Karty s hlavickou maju
   preto horny okraj bez odsadenia a boky presne 16 px; hlavicka sa zapornym
   okrajom -16 px roztiahne az k hranam a zaoblene rohy karty ju orezu
   (overflow:hidden). */
${B} { --nd-pas:rgba(184,203,230,.55); --nd-pas-ciara:rgba(110,140,185,.38); }
${B} ${CASY}, ${B} #WorkcenterDetail--Main_SimpleForm .sapUiForm, ${B} #WorkcenterDetail--Order_Status_Flexbox,
${B} #__pda_detail_rightcol__, ${B} #__pda_left_box_graf__ {
  padding-top:0 !important; padding-left:16px !important; padding-right:16px !important; overflow:hidden !important; }
${B} .nd-nadpis { background:var(--nd-pas) !important; border-bottom:1px solid var(--nd-pas-ciara) !important;
  padding:11px 16px 9px !important; margin:0 -16px 12px !important; border-radius:0 !important;
  width:calc(100% + 32px) !important; max-width:none !important; flex:0 0 calc(100% + 32px) !important;
  box-sizing:border-box !important; }
/* paralelne procesy: hlavicka je hlavicka zoznamu appky, uz sedi cez celu sirku */
${B} #WorkcenterDetail--OrderStatus_List-header, ${B} #WorkcenterDetail--Order_Info_Buttons_FlexBox .sapMListHdrText {
  background:var(--nd-pas) !important; border-bottom:1px solid var(--nd-pas-ciara) !important;
  padding:11px 16px 9px !important; }
/* popis operacie: hlavicka a "Cely text" su dve bunky mriezky, pas je preto
   samostatny pseudo-prvok pod nimi */
${B} #__pda_opis_button__ { position:relative !important; overflow:hidden !important;
  padding-top:11px !important; row-gap:18px !important; }
${B} #__pda_opis_button__::before { content:''; position:absolute; left:0; right:0; top:0; height:40px;
  background:var(--nd-pas); border-bottom:1px solid var(--nd-pas-ciara); pointer-events:none; }
${B} #__pda_opis_button__ > * { position:relative; z-index:1; }
/* pravy panel PDA */
${B} #__pda_hf_menu__ .hf-nadpis { background:var(--nd-pas) !important; border-bottom:1px solid var(--nd-pas-ciara) !important;
  margin:-12px -12px 8px !important; padding:12px 14px 10px !important; text-align:left !important; }
/* pracovny zoznam: hlavicka panela appky ako zaobleny pas */
${B} .pda-left-box .sapMPanelHdr, ${B} .pda-left-box .sapMPanelHeaderTB {
  background:var(--nd-pas) !important; border-bottom:1px solid var(--nd-pas-ciara) !important;
  border-radius:12px 12px 0 0 !important; }

/* ================= ZAKAZKA A MATERIAL: riadky oddelene ciarami =================
   Popisok vlavo v stlpci, hodnota vpravo, medzi riadkami tenka ciara, pekne pismo.
   Dvojbodku za popiskom pridaval dizajn (nie appka), v novom rozlozeni ju netreba. */
${B} #WorkcenterDetail--Main_SimpleForm .sapUiForm { gap:0 !important; padding:14px 20px 10px !important; }
${B} #WorkcenterDetail--Main_SimpleForm .sapUiForm .nd-riadok {
  display:grid !important; grid-template-columns:minmax(120px, 36%) minmax(0,1fr) !important;
  column-gap:16px !important; align-items:baseline !important; padding:8px 2px !important;
  border-bottom:1px solid rgba(140,165,200,.32) !important;
  font-family:"Segoe UI",-apple-system,Roboto,sans-serif !important; }
${B} #WorkcenterDetail--Main_SimpleForm .sapUiForm .nd-riadok:last-child { border-bottom:0 !important; }
${B} #WorkcenterDetail--Main_SimpleForm .sapUiForm .nd-riadok > .sapMLabel,
${B} #WorkcenterDetail--Main_SimpleForm .sapUiForm .nd-riadok > .k {
  display:block !important; font-family:inherit !important; font-size:15px !important; font-weight:400 !important;
  color:#3d5578 !important; text-transform:none !important; letter-spacing:0 !important; }
${B} #WorkcenterDetail--Main_SimpleForm .sapUiForm .nd-riadok > .sapMLabel::after,
${B} #WorkcenterDetail--Main_SimpleForm .sapUiForm .nd-riadok > .k::after { content:'' !important; }
${B} #WorkcenterDetail--Main_SimpleForm .sapUiForm .nd-riadok > .sapMText,
${B} #WorkcenterDetail--Main_SimpleForm .sapUiForm .nd-riadok > .v {
  display:block !important; font-family:inherit !important; font-size:15px !important; font-weight:600 !important;
  color:#13315c !important; margin-left:0 !important; }
${B} #WorkcenterDetail--Main_SimpleForm .sapUiForm .nd-riadok .sapMText * { font-family:inherit !important; }
${B} .nd-mat-title { font-family:"Segoe UI",-apple-system,Roboto,sans-serif !important; font-size:21px !important;
  font-weight:700 !important; color:#13315c !important; margin:2px 0 6px !important; padding-bottom:10px !important;
  border-bottom:1px solid rgba(140,165,200,.32) !important; }

/* ================= SAP CASY o nieco sirsie =================
   Mriezka detailu: lavy stlpec (zakazka, popis) uzsi, pravy (SAP casy,
   paralelne procesy) sirsi, aby sa tri kolace pohodlne zmestili vedla seba. */
${B} #WorkcenterDetail--OrderDetails_FlexBox.nd-detail-grid {
  grid-template-columns:minmax(0,.85fr) minmax(0,1.15fr) !important; }

/* ================= mimo kariet priehladne - vidno pozadie ================= */
${B} .sapMPanel.pda-3j-priehladny { background:transparent !important; border-color:transparent !important;
  box-shadow:none !important; }
${B} .sapMPanel.pda-3j-priehladny > .sapMPanelWrappingDiv,
${B} .sapMPanel.pda-3j-priehladny > .sapMPanelWrappingDivTb,
${B} .sapMPanel.pda-3j-priehladny > .sapMPanelContent { background:transparent !important; }

/* ================= SAP CASY ================= */
/* sekcia ako polopriehladna "sklenena" karta, vnutri tri samostatne karty */
${B} ${CASY} { background:var(--nd-karta) !important; -webkit-backdrop-filter:blur(10px); backdrop-filter:blur(10px);
  border:1px solid var(--nd-okraj) !important; border-radius:18px !important;
  box-shadow:var(--nd-tien) !important; padding:14px 16px 16px !important;
  gap:12px !important; align-items:stretch !important; }
${B} ${CASY} > .nd-nadpis { margin:0 0 2px 2px !important; }
${B} ${CASY} > .sapMVBox { flex:1 1 0 !important; min-width:140px !important; position:relative !important;
  background:var(--nd-vnutro) !important; border:1px solid var(--nd-okraj) !important; border-radius:14px !important;
  box-shadow:0 2px 10px rgba(20,60,120,.06) !important; padding:12px 10px 12px !important;
  align-items:center !important; box-sizing:border-box !important; }
/* nadpis grafu - normalne pismo, nie velke rozpalcovane */
${B} ${CASY} > .sapMVBox .sapMLabel { order:1 !important; font-size:15px !important; font-weight:800 !important;
  letter-spacing:0 !important; text-transform:none !important; color:#13315c !important; margin:0 !important;
  text-align:center !important; }
/* cas hned pod nadpisom, tucne */
${B} ${CASY} > .sapMVBox .sapMText { order:2 !important; font-size:14px !important; font-weight:800 !important;
  color:#13315c !important; margin:2px 0 8px !important; text-align:center !important; }
/* vacsi kolac; rastie so sirkou okna, ale ma strop, aby sa zmestili tri vedla seba */
${B} ${CASY} canvas.pda-3d { order:3 !important; width:clamp(110px, 10.5vw, 180px) !important;
  height:clamp(110px, 10.5vw, 180px) !important; margin:4px auto 8px !important; }
/* percento a "hotovo" presne v strede kolaca */
${B} ${CASY} .nd-pct { transform:translate(-50%, -50%) !important; }
${B} ${CASY} .nd-pct .c { font-size:26px !important; font-weight:800 !important; }
${B} ${CASY} .nd-pct .h { display:block !important; font-size:12px !important; margin-top:4px !important; }
/* legenda pod kolacom */
${B} ${CASY} .nd-legenda { display:flex !important; order:4 !important; flex-wrap:wrap !important;
  gap:6px 16px !important; justify-content:center !important; margin-top:2px !important;
  font-size:12px !important; font-weight:600 !important; color:#2a3d5c !important; }
${B} ${CASY} .nd-legenda i { width:11px !important; height:11px !important; }

/* ================= PARALELNE PROCESY ================= */
${B} ${SEKCIA} .sapMList { background:var(--nd-karta) !important; -webkit-backdrop-filter:blur(10px);
  backdrop-filter:blur(10px); border:1px solid var(--nd-okraj) !important; box-shadow:var(--nd-tien) !important; }
/* jedna polozka = biela karta (namiesto zelenej plochy) */
${B} ${ZOZNAM} .sapMLIB, ${B} ${ZOZNAM} .pda-aktivita { background:var(--nd-vnutro) !important;
  border:1px solid var(--nd-okraj) !important; border-radius:14px !important;
  box-shadow:0 2px 10px rgba(16,36,63,.06) !important; padding:10px 14px !important; margin:0 0 8px !important; }
${B} ${ZOZNAM} .sapMLIB:last-child { margin-bottom:0 !important; }
/* ikona pred textom + dva riadky: stav tucne, meno pod nim */
${B} ${ZOZNAM} [id^="WorkcenterDetail--Title_FlexBox-"] { display:grid !important;
  grid-template-columns:46px minmax(0,1fr) !important; grid-template-rows:auto auto !important;
  column-gap:14px !important; row-gap:2px !important; align-items:center !important; }
${B} ${ZOZNAM} [id^="WorkcenterDetail--Title_FlexBox-"]::before { content:''; grid-column:1; grid-row:1 / span 2;
  width:46px; height:46px; border-radius:12px;
  background:linear-gradient(160deg,#28b463 0%,#1b8a4c 100%) center / 26px 26px no-repeat;
  background-image:${IKONA_PROCESU}, linear-gradient(160deg,#28b463 0%,#1b8a4c 100%);
  background-size:26px 26px, 100% 100%; background-position:center, center; background-repeat:no-repeat;
  box-shadow:0 4px 10px rgba(27,138,76,.28); }
${B} ${ZOZNAM} [id^="WorkcenterDetail--Title_FlexBox-"] > :nth-child(1) { grid-column:2 !important; grid-row:1 !important;
  font-size:15px !important; font-weight:800 !important; color:#13315c !important; }
${B} ${ZOZNAM} [id^="WorkcenterDetail--Title_FlexBox-"] > :nth-child(2) { grid-column:2 !important; grid-row:2 !important;
  font-size:13px !important; font-weight:500 !important; color:#4a6285 !important; }
${B} ${ZOZNAM} [id^="WorkcenterDetail--Title_FlexBox-"] > * { margin:0 !important; line-height:1.3 !important; }
/* velky cas */
${B} ${ZOZNAM} [id^="WorkcenterDetail--Timer_Label-"] { font-size:22px !important; font-weight:800 !important;
  color:#13315c !important; letter-spacing:.02em !important; }
${B} ${ZOZNAM} [id^="WorkcenterDetail--Timer_FlexBox-"] { align-items:center !important; gap:16px !important; }
/* Zastavit: cervene tlacidlo (len vzhlad - je to stale tlacidlo appky) */
${B} ${ZOZNAM} .sapMBtn .sapMBtnInner { background:linear-gradient(180deg,#e8453c 0%,#cc2b25 100%) !important;
  border:0 !important; border-radius:12px !important; padding:10px 20px !important;
  box-shadow:0 5px 14px rgba(204,43,37,.32) !important; height:auto !important; }
${B} ${ZOZNAM} .sapMBtn .sapMBtnContent, ${B} ${ZOZNAM} .sapMBtn bdi,
${B} ${ZOZNAM} .sapMBtn .sapUiIcon { color:#fff !important; font-weight:800 !important; font-size:15px !important; }
`;
            document.head.appendChild(st);
        }

        // Velke panely stranky (Osobny stav, pracovisko) su len obaly - bez vlastnej
        // plochy, aby medzi kartami presvitalo pozadie. Karty v nich ostavaju biele.
        function priehladnePanely() {
            const left = document.getElementById('WorkcenterDetail--LeftColumn_FlexBox');
            const panely = [left && left.closest('.sapMPanel')];
            document.querySelectorAll('.pda-panel-tesny').forEach((p) => panely.push(p));
            panely.forEach((p) => {
                if (p && !p.classList.contains('pda-3j-priehladny')) p.classList.add('pda-3j-priehladny');
            });
        }

        // v patke nech je jasne, ze bezi verzia 3J
        function patka() {
            const r = document.querySelector('#__pda_nd_footer__ .nd-f-r');
            if (!r) return;
            let v = '';
            try { v = GM_info && GM_info.script ? GM_info.script.version : ''; } catch (e) { /* ignore */ }
            const t = 'PDA App Extension · verzia 3J' + (v ? ' (' + v + ')' : '') + ' · Jaro · HF Slovakia';
            if (r.textContent !== t) r.textContent = t;
        }

        function apply() {
            injectStyles();
            [priehladnePanely, patka].forEach((f) => {
                try { f(); } catch (e) { console.warn(LOG, 'verzia 3J:', f.name, e); }
            });
        }

        DomWatch.add(apply);
        onReady(apply);
    }

    /* -------------------- 3.21 Ladiaci vypis ---------------------------- */

    function modDebugLog() {
        XhrBus.subscribe((ev) => {
            console.log('--- [PDA executeBO] ---');
            console.log('request:', ev.requestRaw);
            console.log('response:', ev.responseRaw);
            console.log('-----------------------');
        });
    }

    /* ========================================================================
     *  4. ZOZNAM MODULOV
     *     Nove vylepsenie = napisat funkciu vyssie a pridat sem jeden riadok.
     * ====================================================================== */

    const MODULES = [
        {
            id: 'header',
            name: 'Vylepšená hlavička',
            desc: 'Slovenské popisky pri ikonách, väčšie meno používateľa, zvýraznené odhlásenie.',
            def: true,
            run: modEnhancedHeader,
        },
        {
            id: 'statusButtons',
            name: 'Farebné tlačidlá',
            desc: 'Farby tlačidiel podľa pravidiel v nastaveniach (predvolene: výroba zelená, prestoj oranžový, chyba červená). Po zapnutí „Nastavovanie tlačidiel“ meníš farby pravým klikom.',
            def: true,
            run: modButtonColors,
        },
        {
            id: 'crossSearch',
            name: 'Vyhľadávanie zákaziek',
            desc: 'Vyhľadávací riadok, ktorý hľadá naprieč všetkými pracoviskami naraz.',
            def: true,
            run: modCrossSearch,
        },
        {
            id: 'preventBack',
            name: 'Blokovanie tlačidla Späť',
            desc: 'Zabráni nechcenému vypadnutiu z aplikácie cez tlačidlo Späť v prehliadači.',
            def: true,
            run: modPreventBack,
        },
        {
            id: 'usersPanel',
            name: 'Panel rýchleho prepínania používateľov',
            desc: 'Bočný panel s tlačidlami na prihlásenie. Používateľov zadáš nižšie v nastaveniach.',
            def: false,
            needs: 'Zdieľaný terminál',
            run: modUsersPanel,
        },
        {
            id: 'scanner',
            name: 'Čiarový skener a RFID karty',
            desc: 'Číta skener aj čítačku kariet. Karta prepne používateľa, číslo zákazky sa vloží do vyhľadávania.',
            def: false,
            needs: 'Hardvér',
            run: modInputListener,
        },
        {
            id: 'drawing',
            name: 'Tlačidlo výkresu',
            desc: 'Načíta Excel s výkresmi a k otvorenej zákazke ukáže číslo výkresu a revíziu.',
            def: false,
            needs: 'Firemná sieť',
            run: modDrawingButton,
        },
        {
            id: 'overview',
            name: 'Prehľad pracovísk na úvode',
            desc: 'Zbalí úvodnú obrazovku do kategórií (Assembly, Welding, Machining…). Vidíš len čísla strojov a či pracujú; klik otvorí stroj.',
            def: true,
            run: modWorkcenterOverview,
        },
        {
            id: 'orderList',
            name: 'Zoznam zákaziek ako pilulky',
            desc: 'V detaile pracoviska stlačí každú zákazku do jedného kompaktného riadku a zoznam natiahne až po spodok obrazovky — zmestí sa ich viac.',
            def: true,
            run: modOrderListPills,
        },
        {
            id: 'opDescription',
            name: 'Popis operácie na celú obrazovku',
            desc: 'Namiesto drobného textu veľké tlačidlo „Popis operácie“; po kliknutí sa popis ukáže cez celú obrazovku vo veľkom písme (zavrie sa klikom vedľa alebo Esc).',
            def: true,
            run: modOperationDescription,
        },
        {
            id: 'chartStyle',
            name: 'Krajší graf vyťaženia',
            desc: 'Graf pod pracovným zoznamom: čas dole len ako HH:MM, zaoblené a tenšie pásy, nižšie plátno (uvoľní miesto zoznamu). Údaje sa nemenia, len vzhľad.',
            def: true,
            run: modChartStyle,
        },
        {
            id: 'statusTable',
            name: 'Krajšia tabuľka stavov',
            desc: 'Tabuľka, ktorá sa otvorí tlačidlom s mriežkou pri grafe: dátum a čas v krátkom tvare, stav ako farebná pilulka, vyššie a striedavo podfarbené riadky.',
            def: true,
            run: modStatusTable,
        },
        {
            id: 'detailHeader',
            name: 'Kompaktná hlavička detailu',
            desc: 'Zákazka, materiál a Production Order stlačí do jedného kompaktného boxu a prepínač Machine, okienko VÝKRES aj tlačidlo Operation Complete dá do jedného riadku. Uvoľní sa tým miesto dole.',
            def: true,
            run: modDetailHeader,
        },
        {
            id: 'donut3d',
            name: 'Priestorové koláčové grafy',
            desc: 'Tri koláče s časmi SAP (Setup / Machine / Labor) nakloní ako pohľad zboku a pridá tieň. Dáta ani hodnoty sa nemenia, iba vzhľad.',
            def: true,
            run: modDonut3D,
        },
        {
            id: 'fullLeft',
            name: 'Ľavý panel na celú výšku',
            desc: '⚠️ Rozpracované — zatiaľ rozhadzuje rozloženie (stĺpec sa pripne na zlé miesto a panel pracoviska sa zosype). Nezapínať, kým to nedoladíme.',
            def: false,
            run: modFullHeightLayout,
        },
        {
            id: 'hfMenu',
            name: 'Menu HF Slovakia (vpravo)',
            desc: 'Zvislý panel pri pravom okraji s pripravovanými funkciami HF Slovakia (CHIPS, privolanie majstra, odvoz materiálu, TOOLSHOP, Flexus). Tlačidlá zatiaľ ukážu okno „vo vývoji“.',
            def: true,
            run: modHfMenu,
        },
        {
            id: 'newDesign',
            name: 'Nový dizajn HF Slovakia',
            desc: 'Prezlečenie celej aplikácie: svetlomodré pozadie, biele karty, horný pruh s logom, stavové tlačidlá s ikonou a podnadpisom, nadpisy sekcií, pätka. Všetky pôvodné ovládacie prvky ostávajú — menia sa len farby a tvary.',
            def: true,
            run: modNewDesign,
        },
        {
            id: 'v3j',
            name: 'Dizajn 3J (Jaro)',
            desc: 'Úpravy vzhľadu verzie 3J: SAP časy ako tri samostatné karty s veľkým koláčom, percentom a legendou; paralelné procesy ako karty s ikonou; mimo kariet priehľadná stránka s pozadím. Mení len vzhľad.',
            def: true,
            run: modV3J,
        },
        {
            id: 'debug',
            name: 'Ladiaci výpis do konzoly',
            desc: 'Vypisuje sieťovú komunikáciu aplikácie do konzoly prehliadača. Bežne netreba.',
            def: false,
            run: modDebugLog,
        },
    ];


    /* ------------------------------------------------------------------
     *  Subor s nastaveniami (zaloha / prenos medzi terminalmi).
     *  Prehliadac nevie pisat na disk podla cesty; vie to len do suboru, ktory
     *  clovek raz vyberie v systemovom okne "Ulozit ako" (File System Access API).
     *  Odkaz na subor sa uklada do IndexedDB, takze prezije obnovenie stranky;
     *  po restarte Chromu treba raz potvrdit pristup. Po kazdom ulozeni
     *  nastaveni sa subor prepise; "Nacitat zo suboru" nahra vsetko naspat
     *  (aj na inom pocitaci - napr. zo sietoveho disku).
     * ---------------------------------------------------------------- */
    const SETTINGS_FILE_KEYS = {
        modules: KEY_MODULES, users: KEY_USERS, pdm: KEY_PDM, excel: KEY_EXCEL,
        groups: KEY_GROUPS, buttons: KEY_BUTTONS, admin: KEY_ADMIN,
    };

    const SettingsFile = (function () {
        const DB = 'pda_settings_db', STORE = 'handles', KEY = 'settings_file_handle';
        let cached; // undefined = este necitane

        function openDb() {
            return new Promise((res, rej) => {
                const q = W.indexedDB.open(DB, 1);
                q.onupgradeneeded = () => q.result.createObjectStore(STORE);
                q.onsuccess = () => res(q.result);
                q.onerror = () => rej(q.error);
            });
        }
        async function get() {
            if (cached !== undefined) return cached;
            try {
                const db = await openDb();
                cached = await new Promise((res, rej) => {
                    const r = db.transaction(STORE, 'readonly').objectStore(STORE).get(KEY);
                    r.onsuccess = () => res(r.result || null);
                    r.onerror = () => rej(r.error);
                });
            } catch (e) { cached = null; }
            return cached;
        }
        async function set(handle) {
            const db = await openDb();
            await new Promise((res, rej) => {
                const tx = db.transaction(STORE, 'readwrite');
                tx.objectStore(STORE).put(handle, KEY);
                tx.oncomplete = res;
                tx.onerror = () => rej(tx.error);
            });
            cached = handle;
        }
        async function clear() {
            const db = await openDb();
            await new Promise((res, rej) => {
                const tx = db.transaction(STORE, 'readwrite');
                tx.objectStore(STORE).delete(KEY);
                tx.oncomplete = res;
                tx.onerror = () => rej(tx.error);
            });
            cached = null;
        }
        return { get, set, clear, supported: typeof W.showSaveFilePicker === 'function' };
    })();

    function suiteVersion() {
        try { return GM_info && GM_info.script ? String(GM_info.script.version) : ''; } catch (e) { return ''; }
    }

    // vsetko, co je v ulozisku Tampermonkey, v jednom objekte
    function collectSettings() {
        const out = {};
        for (const k of Object.keys(SETTINGS_FILE_KEYS)) out[k] = loadJson(SETTINGS_FILE_KEYS[k], null);
        return { aplikacia: 'PDA Suite', verzia: suiteVersion(), ulozene: new Date().toISOString(), nastavenia: out };
    }

    // nahra nastavenia zo suboru do uloziska; vrati pocet prevzatych casti
    function applySettingsObject(obj) {
        const n = obj && obj.nastavenia ? obj.nastavenia : obj;
        if (!n || typeof n !== 'object' || Array.isArray(n)) throw new Error('súbor neobsahuje nastavenia PDA Suite');
        let count = 0;
        for (const k of Object.keys(SETTINGS_FILE_KEYS)) {
            if (n[k] !== undefined && n[k] !== null) { saveJson(SETTINGS_FILE_KEYS[k], n[k]); count++; }
        }
        if (!count) throw new Error('v súbore nie je žiadna známa časť nastavení');
        return count;
    }

    async function ensureFilePermission(handle, mode, interactive) {
        let p = await handle.queryPermission({ mode });
        if (p === 'granted') return true;
        if (interactive) p = await handle.requestPermission({ mode });
        return p === 'granted';
    }

    async function writeSettingsToHandle(handle) {
        const w = await handle.createWritable();
        await w.write(JSON.stringify(collectSettings(), null, 2));
        await w.close();
    }

    // po ulozeni nastaveni prepise vybrany subor (ak je a je povoleny); nikdy nehadze chybu
    async function mirrorSettingsToFile(interactive) {
        try {
            const h = await SettingsFile.get();
            if (!h) return false;
            if (!(await ensureFilePermission(h, 'readwrite', !!interactive))) {
                console.log(LOG, 'súbor s nastaveniami: chýba povolenie na zápis (potvrď v nastaveniach)');
                return false;
            }
            await writeSettingsToHandle(h);
            console.log(LOG, 'nastavenia zapísané do súboru', h.name);
            return true;
        } catch (e) {
            console.warn(LOG, 'zápis nastavení do súboru zlyhal', e);
            return false;
        }
    }

    async function pickSettingsFile() {
        const h = await W.showSaveFilePicker({
            suggestedName: 'pda-suite-nastavenia.json',
            types: [{ description: 'Nastavenia PDA Suite', accept: { 'application/json': ['.json'] } }],
        });
        await SettingsFile.set(h);
        await writeSettingsToHandle(h);
        return h;
    }

    async function loadSettingsFromHandle(handle) {
        if (!(await ensureFilePermission(handle, 'read', true))) throw new Error('chýba povolenie na čítanie súboru');
        const file = await handle.getFile();
        return applySettingsObject(JSON.parse(await file.text()));
    }

    function downloadSettings() {
        const blob = new Blob([JSON.stringify(collectSettings(), null, 2)], { type: 'application/json' });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = 'pda-suite-nastavenia.json';
        document.body.appendChild(a);
        a.click();
        setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
    }

    /* ========================================================================
     *  5. PANEL NASTAVENI
     * ====================================================================== */

    const OVERLAY_ID = '__pda_settings_overlay__';
    const GEAR_ID = '__pda_settings_gear__';
    const PASS_ID = '__pda_settings_pass__';

    function injectSettingsStyles() {
        if (document.getElementById('__pda_settings_styles__')) return;
        const style = document.createElement('style');
        style.id = '__pda_settings_styles__';
        style.textContent = `
#${GEAR_ID} {
  position: fixed; right: 14px; bottom: 14px; z-index: 2147483000;
  width: 38px; height: 38px; border-radius: 50%; border: none;
  background: #313175; color: #fff; font-size: 19px; line-height: 1;
  cursor: pointer; opacity: .55; box-shadow: 0 2px 8px rgba(0,0,0,.3);
  transition: opacity .15s;
}
#${GEAR_ID}:hover { opacity: 1; }
#${OVERLAY_ID} {
  position: fixed; inset: 0; background: rgba(0,0,0,.5); z-index: 2147483001;
  display: flex; align-items: center; justify-content: center;
  font: 14px/1.5 -apple-system, "Segoe UI", Roboto, sans-serif; color: #1a1a1f;
}
.pda-set-box {
  background: #fff; border-radius: 12px; width: 94%; max-width: 720px;
  max-height: 88vh; display: flex; flex-direction: column; overflow: hidden;
  box-shadow: 0 10px 40px rgba(0,0,0,.35);
}
.pda-set-head {
  display: flex; justify-content: space-between; align-items: center;
  padding: 14px 18px; background: #313175; color: #fff; flex-shrink: 0;
}
.pda-set-head b { font-size: 1.02rem; }
.pda-set-x { background: none; border: 0; color: #fff; font-size: 24px; line-height: 1; cursor: pointer; }
.pda-set-body { padding: 6px 18px 18px; overflow: auto; }
.pda-set-body h3 {
  font-size: .76rem; text-transform: uppercase; letter-spacing: .07em;
  color: #6b7180; margin: 20px 0 8px; font-weight: 600;
}
.pda-mod {
  display: flex; gap: 11px; align-items: flex-start;
  padding: 10px 12px; border: 1px solid #e3e6eb; border-radius: 9px; margin-bottom: 7px;
}
.pda-mod input { margin-top: 3px; width: 17px; height: 17px; flex-shrink: 0; cursor: pointer; }
.pda-mod .nm { font-weight: 600; }
.pda-mod .ds { color: #5c6370; font-size: .87rem; }
.pda-badge {
  font-size: .66rem; text-transform: uppercase; letter-spacing: .05em;
  background: #eef0f4; color: #5c6370; border-radius: 20px; padding: 2px 8px; margin-left: 7px;
  white-space: nowrap;
}
.pda-set-body table { width: 100%; border-collapse: collapse; }
.pda-set-body th {
  text-align: left; font-size: .74rem; text-transform: uppercase;
  letter-spacing: .05em; color: #6b7180; padding: 0 6px 5px 0; font-weight: 600;
}
.pda-set-body td { padding: 0 6px 6px 0; }
.pda-set-body input[type=text], .pda-set-body input[type=password] {
  width: 100%; box-sizing: border-box; padding: 6px 9px;
  border: 1px solid #ccd1d9; border-radius: 6px; font: inherit; font-size: .9rem;
}
.pda-set-body input[type=color] { width: 38px; height: 30px; padding: 0; border: 1px solid #ccd1d9;
  border-radius: 6px; background: #fff; cursor: pointer; vertical-align: middle; }
.pda-set-body input[type=number] { width: 64px; padding: 6px 8px; border: 1px solid #ccd1d9; border-radius: 6px; font: inherit; font-size: .9rem; }
.pda-sw { display: inline-block; width: 16px; height: 16px; border-radius: 4px; border: 1px solid rgba(0,0,0,.15); vertical-align: middle; margin-right: 6px; }
.pda-del { background: #fdecea; border: 1px solid #f0b4ae; color: #b0201a;
  border-radius: 6px; cursor: pointer; padding: 5px 10px; font-size: .82rem; }
.pda-add { background: #eef0f4; border: 1px solid #ccd1d9; border-radius: 7px;
  cursor: pointer; padding: 6px 13px; font-size: .85rem; margin-top: 3px; }
.pda-note { color: #6b7180; font-size: .82rem; margin: 7px 0 0; }
.pda-set-foot {
  display: flex; justify-content: flex-end; gap: 9px; align-items: center;
  padding: 13px 18px; border-top: 1px solid #e3e6eb; flex-shrink: 0; background: #fafbfc;
}
.pda-btn-primary { background: #313175; color: #fff; border: 0; border-radius: 8px;
  padding: 9px 18px; font-weight: 600; cursor: pointer; font-size: .9rem; }
.pda-btn-plain { background: #fff; border: 1px solid #ccd1d9; border-radius: 8px;
  padding: 9px 16px; cursor: pointer; font-size: .9rem; }
`;
        document.head.appendChild(style);
    }

    function openSettings() {
        if (document.getElementById(OVERLAY_ID) || document.getElementById(PASS_ID)) return;
        const pwd = String((settings.admin && settings.admin.password) || '');
        if (!pwd) { openSettingsUnlocked(); return; }
        injectSettingsStyles();

        const ov = document.createElement('div');
        ov.id = PASS_ID;
        ov.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:2147483001;display:flex;' +
            'align-items:center;justify-content:center;font:14px/1.5 -apple-system,"Segoe UI",Roboto,sans-serif;color:#1a1a1f;';
        const box = document.createElement('div');
        box.style.cssText = 'background:#fff;border-radius:12px;width:92%;max-width:360px;padding:18px 20px;box-shadow:0 10px 40px rgba(0,0,0,.35);';
        const t = document.createElement('div');
        t.style.cssText = 'font-weight:600;font-size:1.02rem;margin-bottom:10px;';
        t.textContent = 'Nastavenia PDA Suite — heslo';
        const inp = document.createElement('input');
        inp.type = 'password';
        inp.autocomplete = 'off';
        inp.style.cssText = 'width:100%;box-sizing:border-box;padding:9px 11px;border:1px solid #ccd1d9;border-radius:8px;font:inherit;font-size:1rem;';
        const msg = document.createElement('div');
        msg.style.cssText = 'color:#b0201a;font-size:.85rem;min-height:1.2em;margin-top:6px;';
        const row = document.createElement('div');
        row.style.cssText = 'display:flex;justify-content:flex-end;gap:8px;margin-top:12px;';
        const cancel = document.createElement('button');
        cancel.type = 'button'; cancel.className = 'pda-btn-plain'; cancel.textContent = 'Zrušiť';
        const ok = document.createElement('button');
        ok.type = 'button'; ok.className = 'pda-btn-primary'; ok.textContent = 'Otvoriť';
        row.appendChild(cancel); row.appendChild(ok);
        box.appendChild(t); box.appendChild(inp); box.appendChild(msg); box.appendChild(row);
        ov.appendChild(box);
        document.body.appendChild(ov);
        setTimeout(() => inp.focus(), 50);

        const close = () => ov.remove();
        const submit = () => {
            if (inp.value === pwd) { close(); openSettingsUnlocked(); }
            else { msg.textContent = 'Nesprávne heslo.'; inp.value = ''; inp.focus(); }
        };
        cancel.addEventListener('click', close);
        ok.addEventListener('click', submit);
        inp.addEventListener('keydown', (e) => { if (e.key === 'Enter') submit(); if (e.key === 'Escape') close(); });
        ov.addEventListener('click', (e) => { if (e.target === ov) close(); });
    }

    function openSettingsUnlocked() {
        if (document.getElementById(OVERLAY_ID)) return;
        injectSettingsStyles();

        const draftModules = {};
        MODULES.forEach((m) => { draftModules[m.id] = isModuleOn(m); });
        const draftUsers = settings.users.map((u) => ({ ...u }));
        const draftPdm = { ...settings.pdm };
        const draftExcel = { ...settings.excel };
        const draftGroups = { value: String(settings.groups || '') };
        const draftButtons = buttonRules().map((r) => Object.assign({}, r));
        const draftAdmin = Object.assign({ password: '123456', pickMode: false }, settings.admin || {});

        const overlay = document.createElement('div');
        overlay.id = OVERLAY_ID;

        const box = document.createElement('div');
        box.className = 'pda-set-box';

        // --- hlavicka ---
        const head = document.createElement('div');
        head.className = 'pda-set-head';
        head.innerHTML = '<b>Nastavenia PDA Suite</b>';
        const closeBtn = document.createElement('button');
        closeBtn.className = 'pda-set-x';
        closeBtn.type = 'button';
        closeBtn.textContent = '×';
        head.appendChild(closeBtn);

        // --- telo ---
        const body = document.createElement('div');
        body.className = 'pda-set-body';

        const hMods = document.createElement('h3');
        hMods.textContent = 'Moduly';
        body.appendChild(hMods);

        MODULES.forEach((m) => {
            const row = document.createElement('div');
            row.className = 'pda-mod';

            const cb = document.createElement('input');
            cb.type = 'checkbox';
            cb.checked = draftModules[m.id];
            cb.addEventListener('change', () => { draftModules[m.id] = cb.checked; });

            const txt = document.createElement('div');
            const nm = document.createElement('div');
            nm.className = 'nm';
            nm.textContent = m.name;
            if (m.needs) {
                const badge = document.createElement('span');
                badge.className = 'pda-badge';
                badge.textContent = m.needs;
                nm.appendChild(badge);
            }
            const ds = document.createElement('div');
            ds.className = 'ds';
            ds.textContent = m.desc;
            txt.appendChild(nm);
            txt.appendChild(ds);

            const lbl = document.createElement('label');
            lbl.style.cssText = 'display:flex;gap:11px;align-items:flex-start;cursor:pointer;';
            lbl.appendChild(cb);
            lbl.appendChild(txt);
            row.appendChild(lbl);
            body.appendChild(row);
        });

        // --- pouzivatelia ---
        const hUsers = document.createElement('h3');
        hUsers.textContent = 'Používatelia pre rýchle prepínanie';
        body.appendChild(hUsers);

        const usersWrap = document.createElement('div');
        body.appendChild(usersWrap);

        function renderUsers() {
            usersWrap.innerHTML = '';
            const table = document.createElement('table');
            table.innerHTML =
                '<thead><tr><th>Meno (presne ako v zozname PDA)</th><th style="width:120px">Osobné číslo</th><th style="width:150px">ID karty</th><th style="width:40px"></th></tr></thead>';
            const tbody = document.createElement('tbody');

            draftUsers.forEach((u, i) => {
                const tr = document.createElement('tr');

                const mk = (field, type) => {
                    const td = document.createElement('td');
                    const inp = document.createElement('input');
                    inp.type = type;
                    inp.value = u[field] || '';
                    inp.addEventListener('input', () => { draftUsers[i][field] = inp.value.trim(); });
                    td.appendChild(inp);
                    return td;
                };

                tr.appendChild(mk('username', 'text'));
                tr.appendChild(mk('password', 'password'));
                tr.appendChild(mk('cardId', 'text'));

                const tdDel = document.createElement('td');
                const del = document.createElement('button');
                del.type = 'button';
                del.className = 'pda-del';
                del.textContent = '×';
                del.title = 'Odstrániť';
                del.addEventListener('click', () => { draftUsers.splice(i, 1); renderUsers(); });
                tdDel.appendChild(del);
                tr.appendChild(tdDel);

                tbody.appendChild(tr);
            });

            table.appendChild(tbody);
            usersWrap.appendChild(table);

            const add = document.createElement('button');
            add.type = 'button';
            add.className = 'pda-add';
            add.textContent = '+ Pridať používateľa';
            add.addEventListener('click', () => {
                draftUsers.push({ username: '', password: '', cardId: '' });
                renderUsers();
            });
            usersWrap.appendChild(add);

            const note = document.createElement('p');
            note.className = 'pda-note';
            note.textContent = 'Tieto údaje sa ukladajú len lokálne v Tampermonkey na tomto počítači — nikdy sa neposielajú na GitHub ani nikam inam.';
            usersWrap.appendChild(note);
        }
        renderUsers();

        // --- hromadne zadanie / prenos na iny pocitac ---
        const ioWrap = document.createElement('div');
        ioWrap.style.cssText = 'margin-top:12px;border:1px dashed #ccd1d9;border-radius:9px;padding:11px 12px;';

        const ioTitle = document.createElement('div');
        ioTitle.textContent = 'Hromadné zadanie / prenos na iný počítač';
        ioTitle.style.cssText = 'font-weight:600;font-size:.88rem;margin-bottom:4px;';

        const ioHelp = document.createElement('div');
        ioHelp.className = 'pda-note';
        ioHelp.style.marginTop = '0';
        const IO_HELP_DEFAULT = 'Jeden používateľ na riadok vo formáte:  Meno;osobné číslo;ID karty   (ID karty môže ostať prázdne)';
        ioHelp.textContent = IO_HELP_DEFAULT;

        const ta = document.createElement('textarea');
        ta.rows = 5;
        ta.spellcheck = false;
        ta.placeholder = 'Ján Novák;12345;0116984be8';
        ta.style.cssText = 'width:100%;box-sizing:border-box;margin-top:7px;padding:8px 10px;' +
            'border:1px solid #ccd1d9;border-radius:6px;font:12px/1.5 ui-monospace,Consolas,monospace;resize:vertical;';

        const ioBtns = document.createElement('div');
        ioBtns.style.cssText = 'display:flex;gap:8px;margin-top:7px;';

        const btnImport = document.createElement('button');
        btnImport.type = 'button';
        btnImport.className = 'pda-add';
        btnImport.style.marginTop = '0';
        btnImport.textContent = 'Načítať z textu';
        btnImport.addEventListener('click', () => {
            const parsed = ta.value
                .split('\n')
                .map((line) => line.trim())
                .filter((line) => line && line.indexOf(';') !== -1)
                .map((line) => {
                    const parts = line.split(';').map((s) => (s || '').trim());
                    return { username: parts[0], password: parts[1] || '', cardId: parts[2] || '' };
                })
                .filter((u) => u.username);

            if (parsed.length === 0) {
                ioHelp.style.color = '#b0201a';
                ioHelp.textContent = 'Nenašiel sa žiadny platný riadok. Formát je: Meno;osobné číslo;ID karty';
                return;
            }

            draftUsers.length = 0;
            parsed.forEach((u) => draftUsers.push(u));
            renderUsers();
            ioHelp.style.color = '';
            ioHelp.textContent = 'Načítaných používateľov: ' + parsed.length + '. Ešte to ulož tlačidlom dole.';
        });

        const btnExport = document.createElement('button');
        btnExport.type = 'button';
        btnExport.className = 'pda-add';
        btnExport.style.marginTop = '0';
        btnExport.textContent = 'Vypísať súčasných';
        btnExport.addEventListener('click', () => {
            ta.value = draftUsers.map((u) => [u.username, u.password, u.cardId].join(';')).join('\n');
            ioHelp.style.color = '';
            ioHelp.textContent = IO_HELP_DEFAULT;
        });

        ioBtns.appendChild(btnImport);
        ioBtns.appendChild(btnExport);
        ioWrap.appendChild(ioTitle);
        ioWrap.appendChild(ioHelp);
        ioWrap.appendChild(ta);
        ioWrap.appendChild(ioBtns);
        body.appendChild(ioWrap);

        // --- tlacidla: farby a nastavovanie pravym klikom ---
        const hBtn = document.createElement('h3');
        hBtn.textContent = 'Tlačidlá — farby';
        body.appendChild(hBtn);

        const pickRow = document.createElement('div');
        pickRow.className = 'pda-mod';
        const pickLbl = document.createElement('label');
        pickLbl.style.cssText = 'display:flex;gap:11px;align-items:flex-start;cursor:pointer;';
        const pickCb = document.createElement('input');
        pickCb.type = 'checkbox';
        pickCb.checked = !!draftAdmin.pickMode;
        pickCb.addEventListener('change', () => { draftAdmin.pickMode = pickCb.checked; });
        const pickTxt = document.createElement('div');
        pickTxt.innerHTML = '<div class="nm">Nastavovanie tlačidiel pravým klikom</div>' +
            '<div class="ds">Keď je zapnuté, pravý klik na ktorékoľvek tlačidlo v aplikácii ukáže paletu 12 farieb a „Reset tlačidla“. Výber sa hneď uloží. Po skončení to vypni, aby operátori omylom nemenili farby.</div>';
        pickLbl.appendChild(pickCb); pickLbl.appendChild(pickTxt);
        pickRow.appendChild(pickLbl);
        body.appendChild(pickRow);

        const btnWrap = document.createElement('div');
        body.appendChild(btnWrap);

        function renderButtons() {
            btnWrap.innerHTML = '';
            const table = document.createElement('table');
            table.innerHTML = '<thead><tr><th>Text tlačidla</th><th style="width:150px">ID (ak nemá text)</th><th style="width:70px">Poradie</th><th style="width:110px">Pozadie</th><th style="width:90px">Text</th><th style="width:40px"></th></tr></thead>';
            const tbody = document.createElement('tbody');
            if (draftButtons.length === 0) {
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 6; td.className = 'pda-note'; td.style.padding = '6px 0';
                td.textContent = 'Zatiaľ žiadne pravidlá. Zapni nastavovanie a klikni pravým na tlačidlo v aplikácii.';
                tr.appendChild(td); tbody.appendChild(tr);
            }
            draftButtons.forEach((r, i) => {
                const tr = document.createElement('tr');
                const mk = (field, type, extra) => {
                    const td = document.createElement('td');
                    const inp = document.createElement('input');
                    inp.type = type;
                    if (type === 'number') { inp.step = '0.5'; inp.value = Number.isFinite(Number(r[field])) ? r[field] : 1.5; }
                    else inp.value = r[field] || '';
                    if (extra) extra(inp, td);
                    inp.addEventListener('input', () => {
                        draftButtons[i][field] = type === 'number' ? Number(inp.value) : inp.value.trim();
                        if (field === 'bg') { draftButtons[i].fg = contrastColor(inp.value); renderButtons(); }
                    });
                    td.appendChild(inp);
                    return td;
                };
                tr.appendChild(mk('text', 'text'));
                tr.appendChild(mk('id', 'text'));
                tr.appendChild(mk('poradie', 'number'));
                tr.appendChild(mk('bg', 'color', (inp, td) => { const hex = document.createElement('span'); hex.style.cssText = 'font-size:.78rem;color:#6b7180;margin-left:6px;'; hex.textContent = r.bg || ''; td.appendChild(hex); }));
                tr.appendChild(mk('fg', 'color'));
                const tdDel = document.createElement('td');
                const del = document.createElement('button');
                del.type = 'button'; del.className = 'pda-del'; del.textContent = '×'; del.title = 'Odstrániť';
                del.addEventListener('click', () => { draftButtons.splice(i, 1); renderButtons(); });
                tdDel.appendChild(del);
                tr.appendChild(tdDel);
                tbody.appendChild(tr);
            });
            table.appendChild(tbody);
            btnWrap.appendChild(table);

            const tools = document.createElement('div');
            tools.style.cssText = 'display:flex;gap:8px;flex-wrap:wrap;margin-top:4px;';
            const add = document.createElement('button');
            add.type = 'button'; add.className = 'pda-add'; add.style.marginTop = '0'; add.textContent = '+ Pridať tlačidlo';
            add.addEventListener('click', () => { draftButtons.push({ text: '', id: '', bg: '#2563eb', fg: contrastColor('#2563eb'), poradie: 1.5 }); renderButtons(); });
            const defaults = document.createElement('button');
            defaults.type = 'button'; defaults.className = 'pda-add'; defaults.style.marginTop = '0'; defaults.textContent = 'Obnoviť predvolené farby';
            defaults.addEventListener('click', () => { draftButtons.length = 0; DEFAULT_BUTTON_RULES.forEach((r) => draftButtons.push(Object.assign({}, r))); renderButtons(); });
            tools.appendChild(add); tools.appendChild(defaults);
            btnWrap.appendChild(tools);

            const note = document.createElement('p');
            note.className = 'pda-note';
            note.textContent = 'Pravidlo sa hľadá podľa textu tlačidla (stačí časť textu), ID len keď tlačidlo text nemá. Poradie riadi zoradenie stavových tlačidiel na pracovisku (0 = prvé). Farba textu sa pri zmene pozadia dopočíta sama.';
            btnWrap.appendChild(note);
        }
        renderButtons();

        // --- vykresy ---
        const hPdm = document.createElement('h3');
        hPdm.textContent = 'Služba výkresov (PDM)';
        body.appendChild(hPdm);

        const pdmTable = document.createElement('table');
        pdmTable.innerHTML = '<thead><tr><th>Adresa služby</th><th style="width:210px">API kľúč (nepovinné)</th></tr></thead>';
        const pdmBody = document.createElement('tbody');
        const pdmTr = document.createElement('tr');

        const tdBase = document.createElement('td');
        const inpBase = document.createElement('input');
        inpBase.type = 'text';
        inpBase.value = draftPdm.base || '';
        inpBase.placeholder = 'http://172.16.77.134:9000';
        inpBase.addEventListener('input', () => { draftPdm.base = inpBase.value.trim(); });
        tdBase.appendChild(inpBase);

        const tdKey = document.createElement('td');
        const inpKey = document.createElement('input');
        inpKey.type = 'password';
        inpKey.value = draftPdm.key || '';
        inpKey.addEventListener('input', () => { draftPdm.key = inpKey.value.trim(); });
        tdKey.appendChild(inpKey);

        pdmTr.appendChild(tdBase);
        pdmTr.appendChild(tdKey);
        pdmBody.appendChild(pdmTr);
        pdmTable.appendChild(pdmBody);
        body.appendChild(pdmTable);

        // --- Excel s vykresmi ---
        const hExcel = document.createElement('h3');
        hExcel.textContent = 'Excel s výkresmi';
        body.appendChild(hExcel);

        const excelUrlTable = document.createElement('table');
        excelUrlTable.innerHTML = '<thead><tr><th>Adresa Excelu (nepovinné)</th></tr></thead>';
        const excelUrlBody = document.createElement('tbody');
        const excelUrlTr = document.createElement('tr');
        const tdUrl = document.createElement('td');
        const inpUrl = document.createElement('input');
        inpUrl.type = 'text';
        inpUrl.value = draftExcel.url || '';
        inpUrl.placeholder = 'C:\\Users\\meno\\OneDrive - HF MIXING GROUP\\HFSK O.4 Production - Data source\\AutomatedOQ180.xlsx';
        inpUrl.addEventListener('input', () => { draftExcel.url = inpUrl.value.trim(); });
        tdUrl.appendChild(inpUrl);
        excelUrlTr.appendChild(tdUrl);
        excelUrlBody.appendChild(excelUrlTr);
        excelUrlTable.appendChild(excelUrlBody);
        body.appendChild(excelUrlTable);

        const excelNote = document.createElement('p');
        excelNote.className = 'pda-note';
        excelNote.innerHTML =
            'Sem môžeš dať <b>cestu na disku</b> tak ako v Python appke (<code>C:\\…</code> alebo <code>\\\\server\\…</code>) — Excel sa načíta sám pri každom otvorení aplikácie. ' +
            '<b>Podmienka:</b> na stránke <code>chrome://extensions</code> → Tampermonkey → Podrobnosti zapnúť <b>„Povoliť prístup k URL adresám súborov"</b>.<br>' +
            'Funguje aj adresa <b>http(s)://</b>, ak by Excel visel na serveri.<br>' +
            'Ak pole necháš prázdne, súbor sa vyberá ručne tlačidlom <b>„Excel (nepovinné)"</b> a prehliadač si ho zapamätá.';
        body.appendChild(excelNote);

        const colTable = document.createElement('table');
        colTable.style.marginTop = '9px';
        colTable.innerHTML =
            '<thead><tr><th>Stĺpec — číslo zákazky</th><th>Stĺpec — číslo výkresu</th><th>Stĺpec — verzia</th></tr></thead>';
        const colBody = document.createElement('tbody');
        const colTr = document.createElement('tr');

        [['colOrder', 'H'], ['colDrawing', 'AH'], ['colVersion', 'AI']].forEach(([field, ph]) => {
            const td = document.createElement('td');
            const inp = document.createElement('input');
            inp.type = 'text';
            inp.value = draftExcel[field] || '';
            inp.placeholder = ph;
            inp.addEventListener('input', () => { draftExcel[field] = inp.value.trim().toUpperCase(); });
            td.appendChild(inp);
            colTr.appendChild(td);
        });

        colBody.appendChild(colTr);
        colTable.appendChild(colBody);
        body.appendChild(colTable);

        const colNote = document.createElement('p');
        colNote.className = 'pda-note';
        colNote.textContent = 'Písmená stĺpcov tak, ako ich vidíš v Exceli. Predvolene H, AH a AI. ' +
            'Excel je nepovinný — ak chýba, výkres sa hľadá priamo podľa čísla materiálu.';
        body.appendChild(colNote);


        // --- subor s nastaveniami ---
        const hFile = document.createElement('h3');
        hFile.textContent = 'Súbor s nastaveniami (záloha a prenos na iný počítač)';
        body.appendChild(hFile);

        const fileInfo = document.createElement('div');
        fileInfo.className = 'pda-note';
        fileInfo.style.marginTop = '0';
        fileInfo.textContent = 'Zisťujem…';
        body.appendChild(fileInfo);

        const fileMsg = document.createElement('div');
        fileMsg.className = 'pda-note';
        fileMsg.style.cssText = 'margin:4px 0 0;min-height:1.2em;color:#2f7d43;';

        const fileRow = document.createElement('div');
        fileRow.style.cssText = 'display:flex;gap:8px;flex-wrap:wrap;margin-top:8px;';
        const mkBtn = (label, title) => {
            const b = document.createElement('button');
            b.type = 'button'; b.className = 'pda-add'; b.style.marginTop = '0'; b.textContent = label;
            if (title) b.title = title;
            fileRow.appendChild(b);
            return b;
        };
        const bPick = mkBtn('Vybrať, kam sa má ukladať…', 'Systémové okno Uložiť ako — vyberieš miesto a názov súboru');
        const bSave = mkBtn('Uložiť do súboru teraz');
        const bLoad = mkBtn('Načítať zo súboru', 'Nahrá všetky nastavenia z vybraného súboru a obnoví stránku');
        const bDown = mkBtn('Stiahnuť kópiu', 'Uloží kópiu do priečinka Stiahnuté súbory');
        const bUp = mkBtn('Nahrať zo súboru…', 'Vyber .json súbor s nastaveniami');
        const bForget = mkBtn('Zabudnúť súbor');
        body.appendChild(fileRow);
        body.appendChild(fileMsg);

        const fileNote = document.createElement('p');
        fileNote.className = 'pda-note';
        fileNote.textContent = 'Po každom „Uložiť“ (aj po zmene farby pravým klikom) sa vybraný súbor prepíše. ' +
            'Na inom počítači: „Nahrať zo súboru…“, alebo vyber ten istý súbor na sieťovom disku a daj „Načítať“. ' +
            'Po reštarte Chromu môže prehliadač raz vyžiadať potvrdenie prístupu k súboru. ' +
            'Súbor obsahuje aj osobné čísla používateľov — ulož ho tam, kam majú prístup len vedúci.';
        body.appendChild(fileNote);

        const upInput = document.createElement('input');
        upInput.type = 'file'; upInput.accept = '.json,application/json'; upInput.style.display = 'none';
        body.appendChild(upInput);

        function fileSay(text, isError) { fileMsg.style.color = isError ? '#b0201a' : '#2f7d43'; fileMsg.textContent = text; }
        function fileRefresh() {
            SettingsFile.get().then((h) => {
                if (h) fileInfo.textContent = 'Súbor: ' + h.name + ' — prepisuje sa po každom uložení.';
                else fileInfo.textContent = SettingsFile.supported
                    ? 'Zatiaľ nie je vybraný žiadny súbor. Klikni „Vybrať, kam sa má ukladať…“.'
                    : 'Tento prehliadač nepodporuje výber miesta — použi „Stiahnuť kópiu“ a „Nahrať zo súboru…“.';
                bPick.disabled = bSave.disabled = !SettingsFile.supported;
                bForget.style.display = h ? '' : 'none';
            });
        }
        fileRefresh();

        bPick.addEventListener('click', () => {
            pickSettingsFile().then((h) => { fileSay('Uložené do ' + h.name + '.'); fileRefresh(); })
                .catch((e) => { if (!e || e.name !== 'AbortError') fileSay('Nepodarilo sa: ' + (e && e.message || e), true); });
        });
        bSave.addEventListener('click', () => {
            SettingsFile.get().then((h) => {
                if (!h) { fileSay('Najprv vyber, kam sa má ukladať.', true); return; }
                return mirrorSettingsToFile(true).then((ok) => fileSay(ok ? 'Zapísané do ' + h.name + '.' : 'Zápis zlyhal — povoľ prístup k súboru.', !ok));
            });
        });
        bLoad.addEventListener('click', () => {
            (async () => {
                let h = await SettingsFile.get();
                if (!h) {
                    if (typeof W.showOpenFilePicker !== 'function') { fileSay('Použi „Nahrať zo súboru…“.', true); return; }
                    const picked = await W.showOpenFilePicker({ types: [{ description: 'Nastavenia PDA Suite', accept: { 'application/json': ['.json'] } }], multiple: false });
                    h = picked[0];
                    await SettingsFile.set(h);
                }
                const n = await loadSettingsFromHandle(h);
                fileSay('Načítaných častí: ' + n + '. Obnovujem stránku…');
                shared.intentionalReload = true;
                setTimeout(() => location.reload(), 600);
            })().catch((e) => { if (!e || e.name !== 'AbortError') fileSay('Načítanie zlyhalo: ' + (e && e.message || e), true); });
        });
        bDown.addEventListener('click', () => { downloadSettings(); fileSay('Kópia sa sťahuje.'); });
        bUp.addEventListener('click', () => upInput.click());
        upInput.addEventListener('change', () => {
            const f = upInput.files && upInput.files[0];
            if (!f) return;
            f.text().then((txt) => {
                const n = applySettingsObject(JSON.parse(txt));
                fileSay('Nahraných častí: ' + n + '. Obnovujem stránku…');
                shared.intentionalReload = true;
                setTimeout(() => location.reload(), 600);
            }).catch((e) => fileSay('Nahratie zlyhalo: ' + (e && e.message || e), true));
        });
        bForget.addEventListener('click', () => { SettingsFile.clear().then(() => { fileSay('Súbor zabudnutý (na disku ostáva).'); fileRefresh(); }); });

        // --- heslo do nastaveni ---
        const hPwd = document.createElement('h3');
        hPwd.textContent = 'Heslo do nastavení';
        body.appendChild(hPwd);
        const pwdTable = document.createElement('table');
        pwdTable.innerHTML = '<thead><tr><th>Heslo (pýta sa pri otvorení ozubeného kolieska)</th></tr></thead>';
        const pwdBody = document.createElement('tbody');
        const pwdTr = document.createElement('tr');
        const pwdTd = document.createElement('td');
        const pwdInp = document.createElement('input');
        pwdInp.type = 'password';
        pwdInp.autocomplete = 'new-password';
        pwdInp.value = draftAdmin.password || '';
        pwdInp.addEventListener('input', () => { draftAdmin.password = pwdInp.value; });
        pwdTd.appendChild(pwdInp); pwdTr.appendChild(pwdTd); pwdBody.appendChild(pwdTr); pwdTable.appendChild(pwdBody);
        body.appendChild(pwdTable);
        const pwdNote = document.createElement('p');
        pwdNote.className = 'pda-note';
        pwdNote.textContent = 'Prázdne = bez hesla. Ukladá sa len lokálne v Tampermonkey na tomto počítači.';
        body.appendChild(pwdNote);

        // --- paticka ---
        const foot = document.createElement('div');
        foot.className = 'pda-set-foot';

        const cancel = document.createElement('button');
        cancel.type = 'button';
        cancel.className = 'pda-btn-plain';
        cancel.textContent = 'Zrušiť';

        const save = document.createElement('button');
        save.type = 'button';
        save.className = 'pda-btn-primary';
        save.textContent = 'Uložiť a obnoviť stránku';

        foot.appendChild(cancel);
        foot.appendChild(save);

        box.appendChild(head);
        box.appendChild(body);
        box.appendChild(foot);
        overlay.appendChild(box);
        document.body.appendChild(overlay);

        function close() {
            overlay.remove();
            document.removeEventListener('keydown', onEsc);
        }
        function onEsc(e) { if (e.key === 'Escape') close(); }

        closeBtn.addEventListener('click', close);
        cancel.addEventListener('click', close);
        overlay.addEventListener('click', (e) => { if (e.target === overlay) close(); });
        document.addEventListener('keydown', onEsc);

        save.addEventListener('click', () => {
            saveJson(KEY_MODULES, draftModules);
            saveJson(KEY_USERS, draftUsers.filter((u) => u.username));
            saveJson(KEY_PDM, draftPdm);
            saveJson(KEY_EXCEL, draftExcel);
            saveJson(KEY_GROUPS, draftGroups.value);
            saveJson(KEY_BUTTONS, draftButtons.map(normalizeRule).filter((r) => r.text || r.id));
            saveJson(KEY_ADMIN, { password: String(draftAdmin.password || ''), pickMode: !!draftAdmin.pickMode });
            close();
            mirrorSettingsToFile(true).then(() => {
                shared.intentionalReload = true;
                location.reload();
            });
        });
    }

    function ensureGear() {
        if (document.getElementById(GEAR_ID)) return;
        injectSettingsStyles();
        const gear = document.createElement('button');
        gear.id = GEAR_ID;
        gear.type = 'button';
        gear.title = 'Nastavenia PDA Suite';
        gear.textContent = '⚙';
        gear.addEventListener('click', openSettings);
        document.body.appendChild(gear);
    }

    /* ========================================================================
     *  6. SPUSTENIE
     * ====================================================================== */

    const active = [];
    MODULES.forEach((mod) => {
        if (!isModuleOn(mod)) return;
        try {
            mod.run();
            active.push(mod.id);
        } catch (e) {
            console.warn(LOG, 'modul "' + mod.id + '" sa nepodarilo spustiť', e);
        }
    });

    console.log(LOG, 'aktívne moduly:', active.length ? active.join(', ') : 'žiadne');

    onReady(ensureGear);
    DomWatch.add(ensureGear);

    try {
        GM_registerMenuCommand('Nastavenia PDA Suite', openSettings);
    } catch (e) { /* nedostupne mimo Tampermonkey */ }
})();
