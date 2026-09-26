// ==UserScript==
// @name         PDA Suite 3J (Jaro · HF Slovakia)
// @namespace    http://tampermonkey.net/pda-suite-3j
// @version      3.2.13
// @description  PDA Suite 3J - Jarova verzia (vlastny repozitar JaroTvarozek/PDA-3J). Vychadza z produkcneho buildu 2.3.0, vsetky funkcie zachovane, upravuje sa len dizajn.
// @author       Gabris, Tvarozek
// @updateURL    https://github.com/JaroTvarozek/PDA-3J/raw/refs/heads/main/pda-suite-3j.user.js
// @downloadURL  https://github.com/JaroTvarozek/PDA-3J/raw/refs/heads/main/pda-suite-3j.user.js
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
/* overscroll-behavior: ked zoznam dojde kolieskom na koniec, neposunie sa cela stranka (3.2.13) */
#${SCROLL_ID} { border:0 !important; background:transparent !important; border-radius:14px !important;
  overscroll-behavior:contain !important; }
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

        /*
         * Miesto pod lavym stlpcom az po spodok posuvnej sekcie stranky: spodne
         * odsadenia, ramy a okraje vsetkych obalov (panel pracoviska ma vlastny
         * padding a okraj). Predtym sa s nim nepocitalo, takze obsah bol o 20 - 30 px
         * vyssi nez obrazovka - stranka sa dala o tolko posunut (napr. kolieskom
         * na konci zoznamu) a horna cast so Stretnutia / Prestavka zaliezla pod
         * hlavicku (3.2.13). Hodnoty nezavisia od vysky zoznamu.
         */
        function miestoPodStlpcom(el) {
            const sekcia = el.closest('section');
            let sucet = parseFloat(getComputedStyle(el).marginBottom) || 0;
            for (let e = el.parentElement; e && e !== sekcia && e !== document.body; e = e.parentElement) {
                const cs = getComputedStyle(e);
                sucet += (parseFloat(cs.paddingBottom) || 0) + (parseFloat(cs.borderBottomWidth) || 0) +
                    (parseFloat(cs.marginBottom) || 0);
            }
            if (sekcia) sucet += parseFloat(getComputedStyle(sekcia).paddingBottom) || 0;
            return { sucet, sekcia };
        }

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
                    const posledny = grafR && grafR.height > 0 ? graf : box;
                    const { sucet, sekcia } = miestoPodStlpcom(posledny);
                    // spodok sekcie (nad patkou), odtial sa odrata miesto pod stlpcom
                    const dno = Math.min(sekcia ? sekcia.getBoundingClientRect().bottom : W.innerHeight,
                        W.innerHeight - PATKA_3J);
                    const spodok = dno - Math.max(OKRAJ_3J, sucet + 2);
                    // horny okraj zoznamu bez posunu stranky - vysledok nezavisi od toho, ci je posunuta
                    const horeZoznamu = scR.top + (sekcia ? sekcia.scrollTop : 0);
                    ciel = Math.round(Math.max(MIN_HEIGHT, spodok - podBoxom - podZoznamomVBoxe - horeZoznamu));
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
/* okno v strede obrazovky, velke podla textu: sirka podla najdlhsieho riadku (max 1150 px),
   vyska podla poctu riadkov (max 90 % obrazovky, potom posuvnik vpravo) */
#${OVERLAY_ID} .karta { background:#fff; border-radius:16px; width:auto; max-width:min(1150px,94vw);
  min-width:min(460px,94vw); max-height:90vh;
  display:flex; flex-direction:column; overflow:hidden; box-shadow:0 24px 70px rgba(16,36,63,.4);
  font-family:-apple-system,"Segoe UI",Roboto,sans-serif; }
#${OVERLAY_ID} .hl { display:flex; align-items:center; gap:14px; padding:14px 20px; background:#13315c; color:#fff; }
#${OVERLAY_ID} .hl .n { font-size:16px; font-weight:800; letter-spacing:.08em; text-transform:uppercase; }
#${OVERLAY_ID} .hl .z { font-size:13px; opacity:.8; }
#${OVERLAY_ID} .hl .x { margin-left:auto; background:none; border:0; color:#fff; font-size:30px; line-height:1;
  cursor:pointer; padding:0 4px; }
#${OVERLAY_ID} .telo { padding:14px 30px 22px; overflow-x:hidden; overflow-y:auto; font-size:21px; line-height:1.5;
  color:#17202e; overflow-wrap:anywhere; scrollbar-color:#8ea3c2 #eef3fa; }
/* kazdy kus popisu (medzi ciarkami) na vlastnom riadku, s jemnou deliacou ciarou */
#${OVERLAY_ID} .telo .r { padding:7px 0; border-bottom:1px solid #e6edf6; }
#${OVERLAY_ID} .telo .r:last-child { border-bottom:0; }
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

        /*
         * Popis po riadkoch: pri kazdej ciarke zacina novy riadok (ciarka sa uz
         * nezobrazi), bez medzier na zaciatku a konci, viac medzier za sebou je
         * jedna. Desatinna ciarka medzi cislicami (0,5 mm) sa nedeli, ",," je jeden
         * predel. Povodne zalomenia riadkov ostavaju.
         */
        function riadkyPopisu(text) {
            const riadky = [];
            String(text || '').split(/\r?\n/).forEach((riadok) => {
                riadok.replace(/(\d),(?=\d)/g, '$1\u0001').split(',').forEach((kus) => {
                    const t = kus.replace(/\u0001/g, ',').replace(/\s+/g, ' ').trim();
                    if (t) riadky.push(t);
                });
            });
            return riadky;
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
            riadkyPopisu(text).forEach((riadok) => {
                const r = document.createElement('div');
                r.className = 'r';
                r.textContent = riadok;
                telo.appendChild(r);
            });

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
     * Je to NAS vlastny prvok pripnuty k oknu - z rozlozenia appky sa nic
     * nevybera a nic sa nepresuva. Lezi NAD obsahom stranky (miesto mu stranka
     * nevyhradzuje - tak to chcel pouzivatel, 3.2.13) a da sa schovat pasikom.
     * Je na obrazovke otvoreneho pracoviska aj na uvodnej obrazovke.
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
         * Panel je na obrazovke otvoreneho pracoviska a (od 3.2.13) aj na uvodnej
         * obrazovke s prehladom pracovisk. Na ostatnych strankach (prihlasenie,
         * Reporty, Spravy, Rozvrh, Admin) nie je. Kotva, od ktorej panel zacina,
         * je na oboch obrazovkach prva karta pod tlacidlami Stretnutia / Prestavka:
         * v detaile panel s pracovnym zoznamom, na uvode karta Vyhladat zakazku
         * (bez nej karta Pracoviska). Obrazovka sa pozna podla toho, ci je karta
         * naozaj vidiet - prvky uvodu ostavaju v stranke aj ked je otvoreny detail.
         */
        function panelPracoviska() {
            const left = document.getElementById('WorkcenterDetail--LeftColumn_FlexBox');
            const p = left && left.closest('.sapMPanel');
            if (!p) return null;
            const r = p.getBoundingClientRect();
            return r.height > 0 && r.width > 0 ? p : null;
        }

        // aktualna stranka appky je uvod (Main)? ked sa neda zistit, rozhodne len viditelnost karty
        function naUvode() {
            try {
                const app = W.sap && W.sap.ui && W.sap.ui.getCore().byId('Application');
                const str = app && app.getCurrentPage && app.getCurrentPage();
                if (str && str.getId && String(str.getId()) !== 'Main') return false;
            } catch (e) { /* ignore */ }
            return true;
        }

        function panelDomov() {
            if (!naUvode()) return null;
            const karty = [document.getElementById('__pda_search_sidebar__'), document.getElementById('Main--Workcenter_Panel')];
            for (const k of karty) {
                if (!k) continue;
                const r = k.getBoundingClientRect();
                if (r.width > 0 && r.height > 0) return k;
            }
            return null;
        }

        // po prepnuti stranky (detail <-> uvod) sa panel prestavi hned, nie az pri dalsej zmene v stranke
        let navigaciaNapojena = false;
        function napojNavigaciu() {
            if (navigaciaNapojena) return;
            try {
                const app = W.sap && W.sap.ui && W.sap.ui.getCore().byId('Application');
                if (app && app.attachAfterNavigate) {
                    app.attachAfterNavigate(() => setTimeout(apply, 0));
                    navigaciaNapojena = true;
                }
            } catch (e) { /* ignore */ }
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
            napojNavigaciu();
            const detail = panelPracoviska() || panelDomov();
            const panel = document.getElementById(PANEL_ID);
            const t = document.getElementById(TAB_ID);

            // mimo detailu pracoviska a uvodu nie je vidiet ani panel, ani pasik
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
             * Zaciatok je zarovnany s kotvou (prva karta pod tlacidlami
             * Stretnutia / Prestavka), dole siaha po patku noveho dizajnu, takze
             * vyplni celu volnu plochu vpravo a text v patke ostane vidiet.
             * Obsah stranky sa posuva (hlavne na uvode s rozbalenymi oblastami),
             * preto sa berie poloha kotvy BEZ posunu (+ scrollTop sekcie) - panel
             * tak pri posuvani stoji a nikdy nevylezie nad hlavicku.
             */
            const sekcia = detail.closest('section');
            const posun = sekcia ? sekcia.scrollTop : 0;
            const horeSekcie = sekcia ? sekcia.getBoundingClientRect().top : 0;
            const hore = Math.round(Math.max(horeSekcie + OKRAJ, detail.getBoundingClientRect().top + posun));
            const patka = document.getElementById('__pda_nd_footer__');
            const dole = (OKRAJ + (patka && patka.offsetHeight ? patka.offsetHeight : 0)) + 'px';
            if (hore > 0 && (Math.abs((parseFloat(p.style.top) || 0) - hore) > 4 || p.style.bottom !== dole)) {
                p.style.top = hore + 'px';
                p.style.bottom = dole;
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
        const POZADIE = 'data:image/jpeg;base64,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';

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
     * a od 3.1.3 ma vlastny repozitar JaroTvarozek/PDA-3J. Vsetky funkcie ostavaju
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

        // pozadie karty Zakazka a material: obrobok vpravo dole na bielom (JPEG 1100x619)
        const KARTA_ZAKAZKA = 'data:image/jpeg;base64,/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAcFBQYFBAcGBQYIBwcIChELCgkJChUPEAwRGBUaGRgVGBcbHichGx0lHRcYIi4iJSgpKywrGiAvMy8qMicqKyr/2wBDAQcICAoJChQLCxQqHBgcKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKir/wAARCAJrBEwDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD6QooNH1oAKKKKACiiigAooooAKKDRQAUZopKAFzRRSUAFLRRQAUZoooAKKKKACiiigAooooAKTNLRQAmaM0tFABmiikoAWikpaACko70UAFLRSUALSUUtABmiiigBKWiigApKWigAzRRSUALRRRQAZpKWigApKWkoAOaWikoAWkopaAEpaSigAzRmiigBaM0UlAC5pKKKAFpKKKAFpKKWgBKWikoAM0UUUAFFFFAC0maWigBKXtRSUALSUtFACUUUUAFLSCloAM0maKDQAuaTNFFABRS0lAC0lFFABS0lFABRmiigBaSlpKACiiigAozRRQAUUUUALmkoooAM0tJRQAtFJRQAfhS0lFABmiiigAooooAWikooAM0ZoooAWkoooAKKKKADNFFFABmjNFFABRmiigBaSiigAzRRRQAUUUUAFOHSmU4dKACijvRQAZooooAKKKKACkoooAKKKKACiiigBaSiigApaSloAKSigUALSUUtACUtJRQAtJS0UAJ9KKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooNFABRRRQAUUUUAFFFFAB2ooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACjNFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRR2oAKKKKACiiigAooooAKKKKACiiigAooooAKKKKAClHSkpw6UAJRmiigAooo6UAFJRRQAtJS0lABRRRQAUUCigAoopaAEopaSgAopaKACk6UtFABQelJS0AJRS9qSgA6UUUUAFFFFABRRRQAUUd6KACilpKACiiigAooooAKKKKACiiigAooooAKKKKACg0tFACUUppKACiigUAFFFKaAEooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKOtABRRQKAClpKKAClpKKADNFHaigAooooAKKKKACiiigAooooAWkpaKAEooooAKKKDQAUUUUAFFLRQAlFFFABRRRQAUUtJQAUUUUAFFFFABRQaKACiiigAooooAKKKWgBKKKKACiiigAo6UUUAFFFFABRRiigAooooAKKKKACiiigAooooAKKKKACiiigA7U4dKb2pw6UAJRRRQAUc0GigBKKKKAFpKWkoAWkpaKACiiigAooooAKKKKACiiigApDS0lABRRRQAUUUUAFFFFABRQKKAFpKO9HSgApaKSgApaQ0tACUUUUAFFFFAC0UnaigAoopaAEpaSigA70YopaACjvRRQAUUUUAJQaO9FABRS0GgBKKWkoAOtGOKKKACilooATtRRS0AJ1ooooAKWkooAKKKKACijtRQAUUUUAFFFFAB3ope1JQAtJRSigBKKKU0AJiiiigAoopaAEooooAKWkoNABS0UlABRRRQAUv1pO9LQAlGaKKAClpKWgBKMUUtACUUtJQAUtJS0AJRR3ooAKKKKACiiloASl7UlLQAnalpO9LQAlFLRQAlFLmkoAKKKKAClpKKACilooASilpO1ABRRRQAUUUUAFFFFABRRRQAUUUtACU4dKbTh0oASiijvQAlFLSUAFLSUDpQAoooooAKKKKACiikoAWiiigAooooASl70UlAC0GkooAKKKKACiiigAooooAKWkooAKDS0negBaSlpKAFxmikooAO9FFFABRS0lABS0lLQAlFFLQAUlFHWgAooxQRQAUUUd6ACjvRRQAUUUUALSUUtABSUGigBaM0lFAC0UlHQ0AHel60UUAJRS0lABS0UlABR3opaAEpaSigAopaKAEooNFAC0lLRQAUlGKKAF7UUlLQAlLRSUALSUUUALSUUtACUtFBoAKSig0AGcUUUtACUUUUAHaiiloAQ0UCloAKKSigAooooAKKKWgBKKKKACiiloAKKKSgBaKSloASgUUUALSUUUAFFLSUAFFHaigApaKSgAope9JQAtJRRQAUUUUAFFFFABRRRQAUUUUAFLSUtACU4dKbTh0oASij6UUAFJS0lABRRS0AFFFJQAtFFFABRRSUALRRRQAd6KKKACkoooAKKWkoAKKKWgBKKKKACiiigAooooAKWkooAKKWk70AApaT8aKAClopKAFpKWigBKKWigApKWkoAWijtSUALRRSUABopaKACkpTSUAFLRSUAFFLRQAUlFFABRS0lABRS0UAJRRRQAtFJS0AHagUlHegAxRR0paAEHWlpKWgApO9FFABRmlooASlpKKAFoopKAFpKKWgAooooAKTtS0lABRRRQAopKKKACiiloASiiigAo7UUZoAKKKKAClpBS0AJRRRQAUUUUAFFLSUAFA60UUALSUtJQAtJS0lAC0UlFAAaOlFFAC0UUUAFJRRQAtFFHagBKKKKAFpKWkoAWkpaTvQAUUtJQAUUUUAFFFFABRRRQAUUUUAFOHSm04dKAEooooAKKKKACiiigApKAKKACiiigBaKBRQAUlFLQAUUUlAC9KTvS0UAJQKUUdKAEpe1J3ooAOtKaKSgAopaQUAFFLSd6AFpKWkoAKKKWgApKKO1ABRS0UAJ2paKTNAC+9FFFABRRRQAUUUUAFJS0UAFFJS0AFFFFABSUtFACUUtJQAtJS9qKACiiigApKWigBKWjvRQAUUUUAFGaKQ0ALRRRQAUlLQaAEpcUlL2oAKKSl70AFJS0negBaKSloAKKSloASlpO9FAC0lLRQAUlL2pKAA0vakooAKKKKACiiigAooooAKKWkoAKWkoNABRR2pRQAlLRRQAUUneigApaSigApe1JS0AJS0UUAJRRRQAUtFJQAtJRRQAUtJRQAUtJS0AFFJRQAtJRRQAUtJRQAUdKKKACiijtQAUUUUAFFFFABS0lFABTh0ptOHSgBKPpRRQAUUUUAHeiiigAooooASilpKAFpKWk70ALRRQKAEopaKACikpc0AFFHaigApMUtFAB2pKWkoAKWijNABQaKO1ABSUUvagBKKKWgBKKKWgA7UUlLQAUUUnegBaBRRQAUUUUAANFJS0AFFJRQAUtJS0AFFGaKAENL2oooASlpKU0AJS0UUAFFFJQAtJS0lABS0lHNAAKWiigApKDS0AFFJS0AFBoooASlpBS0AJS0lLQAfWkoooAKWikoAWkoooAKXvSdqWgANAopDxQAtFFJQAUUtJQAUUUUAFFFLQAlFFFABRRS0AFHajtSUALQKTtRQAtHekpaACiikoAKKU0lAC0UUlAC0lLSUALRSUUAFFFFAC0UlFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRmiigAooooAKKKKACiiigApw6U2nDpQAlFFFAB3ooooAKKKKACikpaAEooFFAB3pTRik9qACil6UUAFJ3paKAA0dqSigBaSlFBoASig0dqAClpKWgBO9FL2zSUALRSUUALSUUtABRRSUAFLRRQAd6TvRRQAtFFFABR+FIaMUAL2pKBRQAtHeiigBKKWkoAWkopaAEoooxQAtFJ0ooAKKKKAFooooAQUtFFABQKSloASlo4o7UAA5ooooAKKTvRQAUUYooAM0tJRQAtJRS0AHaiikFAC0UGigBKWikoAKWkooAKKWkoAWkpaSgApaSigAooFFABRRRQAUZpeopMUALSUUUALR2oooASijiigBc0UUUAHWiikoAKWkpaACkoooAKWkooAWikooAO9LSUUAFFBooAKKKKACiiigAooo60AFFFFABRRRQAUUUUAGaKKKACiiigAooooAKKKKACiiigApw6U2nDpQAlFFFABRRRQAUUUUAFJS0lABRS0UAFFFFACUtFFABRRSUALSUtJQAtBpKU0AJRRRQAUtJRQAUUUtACUtFJQAUtJRQAtFFJ3oAWiikoAWiiigApKKKAFopKWgA7UlLRQAlFLSdTQAZpaKSgBaO9IKWgAoopPpQAvak70tFACUtFJQAtFFFABRRRQAUUUlABRRS0AFFJS0AJRRRQAUUtBoASlzSUuKACiikFAC0UUUABpKKWgBKKD7UUAFFFFAC0d6KKACg0neg0AKKSgUtACZooooAKMUUUAFFFFABRRRQAUdqWkoAKWikoAWkopaADtRRSUAFLSd6KACilooASiiigAooooAKKKKACiiigAooooAKKSimIWikzRQAtFJRQAtFJRQAtFJRQAUtJS0AFFJRSGLRRRQAUUUUAFFJRQAtOHSmU4dKACiijvQAUUUUAFFHaigAooooAKKKKACiiigAooooAKKKKAE70UUUALQaSigAooooAKWkooAKKXtR2oASiiloASloooAO9FFFABRRRQAUlFLQAUUUUAJRRQBQAvSkNL3ooASlpKKAFopKXNACUtFFABRRRQAUUUUAJmlpKXtQAUUUUABpKWigBKWiigApKWigApKWigApKWigBKKKWgBKWkooAWiiigAoFFFACd6KUikoAKKKWgAopKKAFooooAKO1FFABRRRQAlLSUtACCloxSUALRRRQAUUUUAFFFFACUtFFABSA0tFACYope1JQAUvSkooAKKKKACijvRQAUUUUAFJRRTEFFFJQAuaKSigBc0lFFABS0lFAC0UlFABS0lLQAUUlFAC0UUUAFFFFABRRRQAUUUUAFFFFABTh0ptOXpQAUUUUhhRRRQAUUUUAFJS0UABoopKAFooooAKKKKACiiigAopKKACg0tGaAEpaSigAo7UCigBaSiigBTSUUvagBKWiigAooooAKTvS0UAFFFFABRRSUALRRSUAFBoxS0AJRS0lABRRRQAUtFFABRRRQAUlLRQAlLRRQAUUUUAJS0UUAFFFFABRRRQAUUUUAFFFJ3oAWijNFABRRRQAUUUUAFFFJQAvaig0UAFFFJQAUtFFABRRSUAFFFLQAUnWlpKAFpKBRQAUtJS0AFGaKSgBaKKKACiiigAopKWgBKWkooAOlFFFABRRRQAUUUUAFJS0lMAooooEFJRRQAUUUx320wHZxRvHrUBYnvSVVibljcKM1B+NOD4NKw7ktFNDZpTQAuaKTNFIB2aM03NFAx1FJRQAtFFJQAtFFFIAooooAKKKKAFo6UlFAC05elMp6/doAKKKKQwo7UYooAKKKKACiiigAoo6UUAFFFFABRRQKACikpaACk70tJigBaKKM0AFBopKADvRS0negAFLRRQAlFFBoAKXtSUtACZpaKKACkpaSgAoopaACkpc0lAC0lFFAC9qKSloAKSlooASjNFFAC9qO1FJQAtFJmigBaKKKAEopaKACiiigAooooAKKKKACiik70AB6UClooAKKKKAEpaTtS0AFFFFABRRRQAUUUUAFFFFABRRSGgBTRRSUALRSGloAKKKSgAope1JmgAozS0UAJRRS0AJS0hoFAC0lLQaAEoozS0AJRRRQAUUGigAooooAKKKSmIKKKKACiikoAKKKKYBRSFsCozJRYRITgZzVZmy1K8ueKjHNWkJscKeBxTApp65HWgQEYpKcWHrTcj1oAUEinhqYGFFAyTNLmoxmnZpDHZpc0zNLSAdRmkzRmgBc0tJRQAtFJRSGLRSUtAC0UmaWgAoopKAFp6/dpmacvSkAtFFApDCiiigAooooAKBRRQAUUUUAFFFFABRRRQAlLRRQAUUUUAFJS0UAFFFFACUtFFACUUUUALSdKWk70AFLSUtABRR2ooAKKTpR1oAKKO1HSgA70tJS0AFJRS0AJS0lKKACiiigAooooASiiloAKSilFABRRRigAoozRQAUUlLQAlLSGgUALRRRQAUUUmKAFopOaXNABRR2pKAFooooAKKKKACkopaACiiigAooooAKKKKACkpaKACkpaKACiiigANFFJQAUtFFAAaSjtRQAUUtJQAtIaKKACiiigAooooAKKKKAEooopgGaKKSgQtFJRmgAoo70hNMAzTS2KQtUbNTSFcV3qPrSnk0AVRIhWmgEGpfrShc07gIp9abJIFHFSMuBVWaM9aSBjWdjyKjMrDgmnBuxFKYt/QZ+lWSKjk1Or461XC7fY07zKVhlsEHpS4qoshzxU6TA8NUtFXJMUUoII4oxSGFGaMUlADs0UlGaAFpaSikAuaKSigB1FJRQAtFJS0hhT16VHUi/doYhaKO9FSUFHeiigAooooAKO9J3paAEo70HrRQAvaikpaACkpc0UAJRS0maAFopOtLQAUGjtR9aACiiigBM4paQ0tABSUtJQAtFFHagAxRRRQAlFFFAC0UfWkzxQAZooooAKKKWgAopKWgAopO9LQAUhpaQ0ALRSUtABSDrS0lAC0UUUAFAopMc0AFLmikoAWiikoAU0UUUAHeiikoAWiikoAWkopRQAUUE0lAC0UhpaACjvSUtABR+FFFABRR2ooAKKKKACiikoAM0tJS0AFFJS0AFFJS0AIKWkpaAEo70YooAKKWkzQAUUUUAGaKBRQAUUUUAFFFFACUUUGmIKKSloAKSjNJmmAtJmkJppaiwhSaaWpM0lUIQmk604gYpKYgxScDrSNIFFVZJznimkBaLimmfbVTz8Uhk3U7CuWvtY7mnCVXFUCKC5jQtT5RXJbq4jgOWxTV1q2SLG4H2HNchrGqO9x5IOB3pLN/wB3mtfZ6akc+uhq6hrxEhMSnFV7XXy0n7wEVkX0mGNV4Hy9aKCsTzO526avAVHzgmr0colQMprz65YgAgkH1Fbmg6s8o8lxlhxn1rOVPS6KUu51UchFWVlB68VRVsc1IHX1rFo0TL3WkNVkmI96nWQMKixVx1FL1ooGJRmijFABS0lFIBaWkooAWikooAWnp92o6kT7tJgOoooqSgooooAKDRRQAUlLQaAEpRSUtACHilopKACiiigAooooAKWkpaAEpaDSd6AFo7UUUAFFJS9qACkpaSgBaSgUUALR9aSigBaSgUUALSUUUALRSUUALQKKSgBaKKTNAC0UneloASlpKKACloooAKKKKACiiigAooooAKSlpKAClpKKACilooAKKSigBaSlooASlopKAClFJRQAYooooAM0UtJQAUtFFABRRSUALRRRQAUUUlABS0lLQAgpaBRQAUlHeigBaSiigAooo7UAFFFFABRRRQAUUUUAFJRRTEFFGaKACikzRmgAozSZpM0wFJppakJpuc07CFJptFBYAcnFMQCkJxUD3cYOAc0xrtMcGqsxXLBcDqarzXIHAqs9wXb2pC6Y681SRNyVZdwyaZIwNRqyhutSsUI4NMRFinDAFAIB5p3B6UwGbual8sSIR60og3Dipo4do5pXHY4vX9HkWbz4lJx1FUrOQqmDXoE8KSJhhmsW50OGRiUGD6itY1NLMzcNdDkbx9z1FAcNxW7deGZHPyOaba+FZw/zynHsK0542J5WY1zIQK1/DdtK03mvGQp6ZFblr4WgjIaQbj6tWultFbptRQKzlUVrIpQd9SMkYxTCjH7uT9KlcAc02O/iTKng+1ZGg2MuGwRVkE9uKyG1Hdek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VVzQAzYWNSLHjrUirilNK4DcUdqKQ0gCiilApgIBTgKAKcBSATFOAoApwFIYmKcBRinYoGJSgUoFOAqRiYpRS0tACVIv3aZT16VLGFFLRSGJRS0UAJRS0UAJRS0UAJRS0UAJ3pKdRigBMUlOooASilooASilooASilooASilooASilooASilooAQUlOooASilooASilooASilooASilooASilooASilooATtRS0UAJ1o70tFACUUtFACUUtFACUUtFACUZpaKAEopaKAEopaKAEopaKAEopaKAEopaKAG0tLRQA2inUUANop2KMUANop2BRgUANop1GKAG0U6jFADKKfijA9KAGUU/A9KTA9KYhtFPwPSjA9KAGUU/A9KMD0oAZRT8D0owPSgBlJUmB6UmB6UAR0lS7R6UbR6UXAhpMVMVHpSbRnpTuIhIpCKm2jPSjaPSncRAaaRU5UelJtHpTuBXIpCKn2jHSk2j0piK+KQirBUelIUX0piIMUY5qXaPSjaPSmBHigrmpgo9KXaPSkBVaMEVCUK+4q+VHpTSi+lMCmB6UpWpnRQcgU4qvpTAgxSgVMqr6Uu1fSkBDt4pNuKn2jHSl2r6UAV8VGRzVvYvpTNo9KYFcLk1etVASogoz0q3bKNnSpk9BojZfnJqMjvVplG7pTCoz0qUxsqN1puMc1Y2KW5FNRQXORVEkaxkjJ/KnhQKsBF9KQqPSlcdiE8U0mpio9KbtGOlMCImkqXaPSlVV9KYiLFKBUu0elOCj0pDIgKdipAo9KdtXPSlcCKlAqTaPSnbR6UrjIhTsVJsX0p20elK4yOlFP2j0pdo9KVwGUU/A9KMD0pXGMp69KXA9KKBn/2Q==';

        // ikona pred paralelnym procesom: dve protismerne sipky (biele), vlozene ako SVG
        const IKONA_PROCESU = "url(\"data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' " +
            "fill='none' stroke='white' stroke-width='2.3' stroke-linecap='round' stroke-linejoin='round'>" +
            "<path d='M4 8.5h14.5M15 5l3.5 3.5L15 12'/><path d='M20 15.5H5.5M9 12l-3.5 3.5L9 19'/></svg>\")";

        function injectStyles() {
            if (document.getElementById(STYLE_ID)) return;
            const st = document.createElement('style');
            st.id = STYLE_ID;
            // spodna lista okna (dialog) aj rozbalovacieho okienka (popover) - rovnake tlacidla
            const LISTA = (x) => [B + ' .sapMDialog > footer', B + ' .sapMPopoverFooter']
                .map((z) => z + x).join(',\n');
            st.textContent = `
/* ================= paleta 3J: svetlomodre sklo namiesto bielej =================
   Plochy nie su biele, ale svetlomodre a polopriehladne (ako v navrhu) - pozadie
   cez ne jemne presvita. --nd-karta su velke karty, --nd-vnutro polozky v nich. */
html body.pda-nd { --nd-karta:rgba(231,237,246,.83); --nd-vnutro:rgba(244,247,251,.93);
  --nd-zebra:rgba(233,238,246,.94); --nd-okraj:rgba(255,255,255,.82);
  --nd-tien:0 8px 26px rgba(20,60,120,.10); }

${B} .pda-left-box, ${B} #__pda_hf_menu__, ${B} #__pda_detail_rightcol__,
${B} #WorkcenterDetail--Order_Status_Flexbox, ${B} #__pda_opis_button__,
${B} #WorkcenterDetail--Main_SimpleForm .sapUiForm, ${B} #WorkcenterDetail--Main_SimpleForm--Form {
  background:var(--nd-karta) !important; border:1px solid var(--nd-okraj) !important;
  box-shadow:var(--nd-tien) !important; -webkit-backdrop-filter:blur(10px); backdrop-filter:blur(10px); }
${B} #WorkcenterDetail--Order_Status_Flexbox { border-radius:18px !important; }
/* vseobecne panely (napr. na inych obrazovkach) - nie biele, ale svetlomodre */
${B} .sapMPanel { background:rgba(229,238,250,.72) !important; border-color:var(--nd-okraj) !important; }
/* horny pruh a patka priehladne */
${B} .nd-topbar { background:rgba(229,238,250,.685) !important; border-bottom:1px solid var(--nd-okraj) !important;
  -webkit-backdrop-filter:blur(10px); backdrop-filter:blur(10px); }
${B} #__pda_nd_footer__ { background:rgba(229,238,250,.72) !important; border-top:1px solid var(--nd-okraj) !important; }
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
  padding:11px 16px 9px !important; margin:0 -16px 12px !important; border-radius:16px 16px 0 0 !important;
  width:calc(100% + 32px) !important; max-width:none !important; flex:0 0 calc(100% + 32px) !important;
  box-sizing:border-box !important; }
/* paralelne procesy: hlavicka je hlavicka zoznamu appky, uz sedi cez celu sirku */
${B} #WorkcenterDetail--OrderStatus_List-header, ${B} #WorkcenterDetail--Order_Info_Buttons_FlexBox .sapMListHdrText {
  background:var(--nd-pas) !important; border-bottom:1px solid var(--nd-pas-ciara) !important;
  padding:11px 16px 9px !important; border-radius:16px 16px 0 0 !important; }
/* popis operacie: hlavicka a "Cely text" su dve bunky mriezky, pas je preto
   samostatny pseudo-prvok pod nimi */
${B} #__pda_opis_button__ { position:relative !important; overflow:hidden !important;
  padding-top:11px !important; row-gap:18px !important; }
${B} #__pda_opis_button__::before { content:''; position:absolute; left:0; right:0; top:0; height:40px;
  background:var(--nd-pas); border-bottom:1px solid var(--nd-pas-ciara); pointer-events:none;
  border-radius:14px 14px 0 0; }
${B} #__pda_opis_button__ > * { position:relative; z-index:1; }
/* pravy panel PDA */
${B} #__pda_hf_menu__ .hf-nadpis { background:var(--nd-pas) !important; border-bottom:1px solid var(--nd-pas-ciara) !important;
  margin:-12px -12px 8px !important; padding:12px 14px 10px !important; text-align:left !important;
  border-radius:17px 17px 0 0 !important; }
/* pracovny zoznam: hlavicka panela appky ako zaobleny pas */
${B} .pda-left-box .sapMPanelHdr, ${B} .pda-left-box .sapMPanelHeaderTB {
  background:var(--nd-pas) !important; border-bottom:1px solid var(--nd-pas-ciara) !important;
  border-radius:14px 14px 0 0 !important; }

/* ================= ZAKAZKA A MATERIAL: riadky oddelene ciarami =================
   Popisok vlavo v stlpci, hodnota vpravo, medzi riadkami tenka ciara, pekne pismo.
   Dvojbodku za popiskom pridaval dizajn (nie appka), v novom rozlozeni ju netreba. */
${B} #WorkcenterDetail--Main_SimpleForm .sapUiForm { gap:0 !important; padding:0 16px 10px !important; }
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
  box-shadow:var(--nd-tien) !important; padding:0 16px 16px !important;
  gap:12px !important; align-items:stretch !important; }
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
${B} ${CASY} .nd-pct { z-index:2 !important; transform:translate(-50%, calc(-50% + var(--k3-dy, 0px))) !important;
  transition:transform .2s ease !important; }
/* velkost cisla rastie s kolacom (platno ma sirku clamp(110px, 10.5vw, 180px)), aby sa vzdy zmestilo do otvoru */
${B} ${CASY} .nd-pct .c { font-size:calc(clamp(110px, 10.5vw, 180px) * .145) !important; font-weight:800 !important; }
${B} ${CASY} .nd-pct .h { display:block !important; font-size:max(9px, calc(clamp(110px, 10.5vw, 180px) * .07)) !important;
  margin-top:2px !important; }
/* legenda pod kolacom */
${B} ${CASY} .nd-legenda { display:flex !important; order:4 !important; flex-wrap:wrap !important;
  gap:6px 16px !important; justify-content:center !important; margin-top:2px !important;
  font-size:12px !important; font-weight:600 !important; color:#2a3d5c !important; }
${B} ${CASY} .nd-legenda i { width:11px !important; height:11px !important; }
/* hotova cast koláča je zelena - bodka v legende rovnako */
${B} ${CASY} .nd-legenda .a i { background:#2fa65a !important; }

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

/* ================= HLAVICKY: pas cez CELU sirku karty (posledne pravidla) =================
   Vseobecne pravidlo pre .nd-nadpis pas neroztiahlo vsade: pri Dokumentacii ho
   prebilo starsie pravidlo s ID (okraj 6 px vlavo), v SAP casoch a Zakazke malo
   ine odsadenie karty nez 16 px, s ktorymi pocita zaporny okraj pasu. Tu je preto
   pre kazdu kartu pravidlo s jej ID (vyssia specificita) a na konci stylu
   (vyhrava aj poradim). Zaoblenie pasu = zaoblenie karty minus 1 px okraja. */
${B} ${CASY}, ${B} #WorkcenterDetail--Main_SimpleForm .sapUiForm, ${B} #WorkcenterDetail--Order_Status_Flexbox,
${B} #__pda_detail_rightcol__, ${B} #__pda_left_box_graf__ {
  padding-top:0 !important; padding-left:16px !important; padding-right:16px !important; }
${B} ${CASY} > [data-nd-nadpis],
${B} #WorkcenterDetail--Main_SimpleForm .sapUiForm [data-nd-nadpis],
${B} #WorkcenterDetail--Order_Status_Flexbox > [data-nd-nadpis],
${B} #__pda_detail_rightcol__ > [data-nd-nadpis],
${B} #__pda_left_box_graf__ > [data-nd-nadpis] {
  margin:0 -16px 12px !important; width:calc(100% + 32px) !important; max-width:none !important;
  box-sizing:border-box !important; padding:11px 16px 9px !important;
  background:var(--nd-pas) !important; border-bottom:1px solid var(--nd-pas-ciara) !important; }
/* v riadkovych kartach (flex vedla seba) sa sirka riesi cez flex-basis,
   v stlpcovej karte zakazky by flex-basis znamenal vysku - tam ostava auto */
${B} ${CASY} > [data-nd-nadpis], ${B} #WorkcenterDetail--Order_Status_Flexbox > [data-nd-nadpis],
${B} #__pda_detail_rightcol__ > [data-nd-nadpis] { flex:0 0 calc(100% + 32px) !important; }
${B} #WorkcenterDetail--Main_SimpleForm .sapUiForm [data-nd-nadpis] { flex:0 0 auto !important; }
${B} ${CASY} > [data-nd-nadpis], ${B} #WorkcenterDetail--Order_Status_Flexbox > [data-nd-nadpis] {
  border-radius:17px 17px 0 0 !important; }
${B} #WorkcenterDetail--Main_SimpleForm .sapUiForm [data-nd-nadpis], ${B} #__pda_detail_rightcol__ > [data-nd-nadpis] {
  border-radius:15px 15px 0 0 !important; }
${B} #__pda_left_box_graf__ > [data-nd-nadpis] { border-radius:13px 13px 0 0 !important; }
/* paralelne procesy: produkcia styluje hlavicku zoznamu selektorom s dvomi ID,
   preto aj tu dve - inak ostala v starom malom rozpalcovanom pisme a bez pasu */
${B} #WorkcenterDetail--Order_Info_Buttons_FlexBox #WorkcenterDetail--OrderStatus_List-header,
${B} #WorkcenterDetail--Order_Info_Buttons_FlexBox .sapMListHdrText {
  font-family:"Segoe UI",-apple-system,Roboto,sans-serif !important; font-size:16px !important;
  font-weight:600 !important; letter-spacing:.03em !important; text-transform:uppercase !important;
  color:#2d4b73 !important; line-height:1.3 !important; height:auto !important;
  background:var(--nd-pas) !important; border-bottom:1px solid var(--nd-pas-ciara) !important;
  padding:11px 16px 9px !important; border-radius:15px 15px 0 0 !important; }

/* ================= uvodna obrazovka: panel Pracoviska stale rozbaleny =================
   Sipka na zbalenie sa skryje a obsah panela je vzdy viditelny. Len CSS - nic sa
   neklika ani neprepina; keby UI5 panel vnutorne "zbalil", obsah ostane aj tak
   zobrazeny (display/height s !important prebiju inline styl, ktory pise UI5). */
${B} #Main--Workcenter_Panel .sapMPanelExpandableIcon,
${B} #Main--Workcenter_Panel [id$="-expandButton"],
${B} #Main--Workcenter_Panel [id$="-CollapsedImg"] { display:none !important; }
${B} #Main--Workcenter_Panel-content,
${B} #Main--Workcenter_Panel > .sapMPanelContent { display:block !important; height:auto !important; }
${B} #Main--Workcenter_Panel .sapMPanelHdrExpandable,
${B} #Main--Workcenter_Panel .sapMPanelWrappingDivTb,
${B} #Main--Workcenter_Panel .sapMPanelWrappingDiv { cursor:default !important; }

/* ================= uvodna obrazovka: oblasti pracovisk vo farbe =================
   Kazda oblast (Assembly, Welding, ...) ma svoju farbu: pruh vlavo, dlazdica ikony
   aj ikona. Farbu dava JS cez premennu --ov-farba (znama oblast pevnu, nova oblast
   automaticky dalsiu z nahradnych farieb). Rozbalena oblast ma jemny nadych svojej
   farby a ram v nej. */
/* 3.2.1: vyraznejsie - pruh vlavo 8 px, cely ram karty jemne vo farbe oblasti
   a tien s nadychom tej istej farby */
${B} #__pda_overview__ .ov-g { --ov-farba:#6b7c95;
  border:3px solid color-mix(in srgb, var(--ov-farba) 45%, white) !important;   /* 3.2.4: 2x hrubsi ram */
  border-left:8px solid var(--ov-farba) !important; border-radius:12px !important;
  background:rgba(255,255,255,.9) !important;
  box-shadow:0 3px 12px color-mix(in srgb, var(--ov-farba) 16%, transparent) !important;
  margin-bottom:10px !important; transition:box-shadow .15s, border-color .15s; }
${B} #__pda_overview__ .ov-g:hover { border-color:color-mix(in srgb, var(--ov-farba) 70%, white) !important;
  border-left-color:var(--ov-farba) !important;
  box-shadow:0 6px 18px color-mix(in srgb, var(--ov-farba) 26%, transparent) !important; }
${B} #__pda_overview__ .ov-g.open { border-color:color-mix(in srgb, var(--ov-farba) 75%, white) !important;
  border-left-color:var(--ov-farba) !important;
  background:color-mix(in srgb, var(--ov-farba) 7%, white) !important; }
${B} #__pda_overview__ .ov-g.open .ov-gh { background:transparent !important; }
${B} #__pda_overview__ .ov-ic { width:44px !important; height:44px !important; border-radius:11px !important;
  background:color-mix(in srgb, var(--ov-farba) 15%, white) !important; color:var(--ov-farba) !important;
  font-size:22px !important; }
${B} #__pda_overview__ .ov-ic svg { width:24px; height:24px; display:block; }
/* pracoviska v rozbalenej oblasti: tenky ram vo farbe oblasti (1 px, jemnejsi
   nez 3 px ram oblasti). Premenna --ov-farba sa dedi z karty oblasti .ov-g. */
${B} #__pda_overview__ .ov-g .ov-chip { border:1px solid color-mix(in srgb, var(--ov-farba) 42%, white) !important; }
${B} #__pda_overview__ .ov-g .ov-chip:hover { border-color:var(--ov-farba) !important;
  background:color-mix(in srgb, var(--ov-farba) 6%, white) !important;
  box-shadow:0 4px 12px color-mix(in srgb, var(--ov-farba) 20%, transparent) !important; }

/* ================= ZAKAZKA A MATERIAL: biela karta so strojarskou grafikou =================
   Len tato karta je plne biela (nie polopriehladna) a ma vlastne pozadie -
   fotografiu obrobku vpravo dole na bielom (3.2.9; predtym technicka kresba). Hlavicka ostava rovnaka ako
   na ostatnych kartach: jej polopriehladny pas by na bielom vysiel svetlejsi,
   preto tu ma plnu farbu, ktora vyzera rovnako - a grafika pod nou nepresvita. */
${B} #WorkcenterDetail--Main_SimpleForm .sapUiForm {
  background:linear-gradient(rgba(255,255,255,.25), rgba(255,255,255,.25)),
    url(${KARTA_ZAKAZKA}) right bottom / cover no-repeat, #ffffff !important;   /* 3.2.11: o 25 % jemnejsi */
  -webkit-backdrop-filter:none !important; backdrop-filter:none !important; }
${B} #WorkcenterDetail--Main_SimpleForm .sapUiForm [data-nd-nadpis] { background:#ccdaed !important; }

/* ================= POPIS OPERACIE: tiez plne biela karta =================
   Rovnako ako Zakazka: biele pozadie bez priehladnosti; pas hlavicky (pseudo-prvok
   ::before) ma plnu farbu, aby vyzeral ako na ostatnych kartach. Pod mysou len
   jemny modry nadych, nech je vidiet, ze sa na kartu da kliknut. */
${B} #__pda_opis_button__ { background:#ffffff !important;
  -webkit-backdrop-filter:none !important; backdrop-filter:none !important; }
${B} #__pda_opis_button__:hover { background:#f6f9ff !important; }
${B} #__pda_opis_button__::before { background:#ccdaed !important; }
/* Nadpis a ukazka su vnutri obalu .stred, ktory ma display:contents - pravidlo
   '> * { z-index:1 }' ich preto nezasiahne a pas hlavicky (::before, absolutne
   polohovany) ich prekryl. Kym bol pas polopriehladny, text presvital; s plnou
   farbou (3.2.3) nadpis POPIS OPERACIE zmizol. Tu ich preto dvihneme nad pas. */
${B} #__pda_opis_button__ .nadpis, ${B} #__pda_opis_button__ .ukazka { position:relative !important; z-index:1 !important; }

/* ================= RAM KARIET: tmavomodry ako hlavicka Dokumentacia (3.2.11) =================
   Vsetky samostatne karty (okna) maju ram #13315c - rovnaka farba ako hlavicka
   Dokumentacie, takze karty su jasne ohranicene aj na rusnom pozadi. Vnutorne
   polozky (jednotlive SAP casy, zaznamy v zoznamoch) ostavaju bez neho. */
${B} .pda-left-box, ${B} #__pda_hf_menu__, ${B} #__pda_detail_rightcol__,
${B} #WorkcenterDetail--Order_Status_Flexbox, ${B} #__pda_opis_button__,
${B} #WorkcenterDetail--Main_SimpleForm .sapUiForm, ${B} ${CASY},
${B} #WorkcenterDetail--Order_Info_Buttons_FlexBox .sapMList,
${B} .sapMPanel.pda-panel-tesny, ${B} #Main--Workcenter_Panel, ${B} #__pda_search_sidebar__ {
  border:1.5px solid #13315c !important; }

/* ================= VYHLADAT ZAKAZKU: pole ako pilulka, 450 px (3.2.13) =================
   Pole bolo cez celu sirku obrazovky (3.2.11 ho skratilo na 20 znakov). Teraz ma
   450 px a tvar pilulky: zaoblene rohy a tenky tmavomodry ram dookola namiesto
   spodnej ciary temy (tu kresli tema ako obrazok v pozadi formulara). Ikonky
   zrusit / hladat ostavaju na konci pola. */
${B} #__pda_custom_search_ui__ .sapMListInfoTBar { justify-content:flex-start !important; }
${B} #__pda_custom_search_ui__ .sapMSF { width:450px !important; max-width:100% !important;
  flex:0 0 auto !important; }
${B} #__pda_custom_search_ui__ .sapMSF .sapMSFF, ${B} #__pda_custom_search_ui__ .sapMSF:hover .sapMSFF {
  background:#fff !important; background-image:none !important; border:1px solid #13315c !important;
  border-radius:999px !important; padding-left:16px !important; overflow:hidden !important;
  box-shadow:0 1px 3px rgba(16,36,63,.10) !important; }
${B} #__pda_custom_search_ui__ .sapMSF.sapMFocus .sapMSFF, ${B} #__pda_custom_search_ui__ .sapMSF .sapMSFF:focus-within {
  border-color:#1c478a !important; box-shadow:0 0 0 3px rgba(28,71,138,.16) !important; }
${B} #__pda_custom_search_ui__ .sapMSFB { border-radius:50% !important; color:#13315c !important; }
${B} #__pda_custom_search_ui__ .sapMSFB:last-of-type { margin-right:3px !important; }

/* ================= PRIHLASOVACIA STRANKA =================
   Len vzhlad dvoch prvkov appky (Login--Login_Button = prihlasenie cez Azure,
   Login--Link = prihlasenie menom a heslom); spravanie prihlasovania sa nemeni.
   Biely text odkazu sa na svetlom pozadi takmer stracal - odkaz je teraz biele
   tlacidlo s tmavym textom. Azure dostalo vyraznejsie tlacidlo a logo Microsoft
   (pseudo-prvok, do vnutra tlacidla UI5 sa nesiaha). */
${B} #Login--Login_Button, ${B} .pda-login-azure { height:auto !important; }
${B} #Login--Login_Button .sapMBtnInner, ${B} .pda-login-azure .sapMBtnInner { display:inline-flex !important; align-items:center !important;
  justify-content:center !important; height:50px !important; min-height:50px !important; padding:0 32px !important;
  background:linear-gradient(180deg,#2f86ea 0%,#1a5fc4 100%) !important; border:0 !important;
  border-radius:12px !important; box-shadow:0 8px 22px rgba(26,95,196,.38) !important;
  transition:transform .13s ease, box-shadow .13s ease !important; }
${B} #Login--Login_Button .sapMBtnInner::before, ${B} .pda-login-azure .sapMBtnInner::before { content:''; flex:0 0 auto; width:18px; height:18px;
  margin-right:12px; background:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 20 20'><rect x='0' y='0' width='9.4' height='9.4' fill='%23f25022'/><rect x='10.6' y='0' width='9.4' height='9.4' fill='%237fba00'/><rect x='0' y='10.6' width='9.4' height='9.4' fill='%2300a4ef'/><rect x='10.6' y='10.6' width='9.4' height='9.4' fill='%23ffb900'/></svg>") center / contain no-repeat; }
${B} #Login--Login_Button:hover .sapMBtnInner, ${B} .pda-login-azure:hover .sapMBtnInner { transform:translateY(-2px);
  box-shadow:0 12px 28px rgba(26,95,196,.46) !important; }
${B} #Login--Login_Button .sapMBtnContent, ${B} #Login--Login_Button bdi,
${B} .pda-login-azure .sapMBtnContent, ${B} .pda-login-azure bdi {
  color:#fff !important; font-size:16px !important; font-weight:700 !important; letter-spacing:.02em !important; }
${B} #Login--Link, ${B} [id^="Login--"] .sapMLnk, ${B} .sapMLnk[id^="Login--"], ${B} .pda-login-link {
  display:inline-block !important; margin-top:16px !important; padding:10px 22px !important;
  background:rgba(255,255,255,.92) !important; color:#13315c !important;
  border:1px solid #c9d7ea !important; border-radius:10px !important;
  font-size:14px !important; font-weight:600 !important; text-decoration:none !important; text-shadow:none !important;
  box-shadow:0 4px 14px rgba(16,36,63,.14) !important; transition:background .13s, color .13s !important; }
${B} #Login--Link *, ${B} [id^="Login--"] .sapMLnk *, ${B} .pda-login-link * { color:inherit !important; text-shadow:none !important; }
${B} #Login--Link:hover, ${B} [id^="Login--"] .sapMLnk:hover, ${B} .pda-login-link:hover { background:#fff !important; color:#1a5fc4 !important;
  text-decoration:underline !important; }

/* ================= DOKUMENTACIA podla vzoru (3.2.6) =================
   Len vzhlad. VYKRES je nas vlastny prvok (tlacidlo modulu vykresu), Components/BOM
   je tlacidlo appky - na to ide len CSS a pseudo-prvky, do jeho vnutra sa nesiaha.
   VYKRES si nechava cislo vykresu a reviziu (funkcna informacia) namiesto
   vseobecneho podnadpisu. */
/* hlavicka: tmavomodry pas s ikonou dokumentu */
${B} #__pda_detail_rightcol__ > [data-nd-nadpis] { background:#13315c !important; color:#fff !important;
  border-bottom:0 !important; display:flex !important; align-items:center !important; gap:9px !important; }
${B} #__pda_detail_rightcol__ > [data-nd-nadpis]::before { content:''; flex:0 0 16px; width:16px; height:16px;
  background:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='white' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z'/><path d='M14 2v4a2 2 0 0 0 2 2h4'/><path d='M16 13H8M16 17H8M10 9H8'/></svg>") center / contain no-repeat; }
${B} #__pda_detail_rightcol__ { align-items:stretch !important; column-gap:14px !important; }

/* --- dlazdica VYKRES --- */
${B} #__pda_detail_rightcol__ #__pda_order_drawing_wrapper__ { flex:0 1 340px !important; min-width:260px !important;
  width:auto !important; margin:0 !important; display:flex !important; align-items:stretch !important; }
${B} #__pda_detail_rightcol__ #__pda_order_drawing_button__ {
  display:grid !important; grid-template-columns:54px minmax(0,1fr) 16px !important;
  grid-template-rows:auto auto auto !important; column-gap:14px !important; row-gap:1px !important;
  align-content:center !important; align-items:center !important; justify-items:start !important;
  width:100% !important; min-height:80px !important; padding:12px 16px !important; margin:0 !important;
  background:#e6f0fc !important; border:1.5px solid #9fc3ef !important; border-radius:12px !important;
  box-shadow:0 2px 8px rgba(16,36,63,.06) !important; text-align:left !important; cursor:pointer !important;
  transition:background .13s, border-color .13s !important; }
${B} #__pda_detail_rightcol__ #__pda_order_drawing_button__:hover { background:#dbe9fc !important; border-color:#6fa6ea !important; }
${B} #__pda_detail_rightcol__ #__pda_order_drawing_button__::before { content:''; grid-column:1; grid-row:1 / span 3;
  width:54px; height:54px; border-radius:50%; background:#cfe2fa url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%231e5aa8' stroke-width='1.8' stroke-linecap='round' stroke-linejoin='round'><rect x='2.5' y='4' width='19' height='16' rx='2'/><path d='M2.5 8h19'/><circle cx='12' cy='14' r='3.2'/><path d='M12 9.6v1.2M12 17.2v1.2M7.6 14h1.2M15.2 14h1.2'/></svg>") center / 28px 28px no-repeat; }
${B} #__pda_detail_rightcol__ #__pda_order_drawing_button__::after { content:'›'; grid-column:3; grid-row:1 / span 3;
  justify-self:end; font-size:26px; line-height:1; color:#13315c; }
/* prvy riadok: nadpis "Vykres" - povodny text "VYKRES" sa len vizualne nahradi */
${B} #__pda_detail_rightcol__ #__pda_order_drawing_button__ > span:first-child { grid-column:2 !important; grid-row:1 !important;
  font-size:0 !important; letter-spacing:0 !important; }
${B} #__pda_detail_rightcol__ #__pda_order_drawing_button__ > span:first-child::after { content:'Výkres'; font-size:18px;
  font-weight:800; color:#13315c; letter-spacing:0; }
${B} #__pda_detail_rightcol__ #__pda_order_drawing_value__ { grid-column:2 !important; grid-row:2 !important;
  font-size:14px !important; font-weight:700 !important; color:#1e3e6b !important; }
${B} #__pda_detail_rightcol__ #__pda_order_drawing_revision__ { grid-column:2 !important; grid-row:3 !important;
  font-size:12px !important; color:#5b6b83 !important; }

/* --- dlazdica Components / BOM (tlacidlo appky) --- */
${B} #__pda_detail_rightcol__ #WorkcenterDetail--BoM_Button, ${B} #__pda_detail_rightcol__ .nd-bom {
  flex:0 1 340px !important; width:auto !important; min-width:260px !important; position:relative !important; }
${B} #__pda_detail_rightcol__ #WorkcenterDetail--BoM_Button .sapMBtnInner, ${B} #__pda_detail_rightcol__ .nd-bom .sapMBtnInner {
  display:grid !important; grid-template-columns:54px minmax(0,1fr) 16px !important; grid-template-rows:auto auto !important;
  column-gap:14px !important; row-gap:2px !important; align-content:center !important; align-items:center !important;
  justify-items:start !important; min-height:80px !important; padding:12px 16px !important;
  background:#fff !important; border:1.5px solid #dfe7f2 !important; border-radius:12px !important; text-align:left !important; }
${B} #__pda_detail_rightcol__ #WorkcenterDetail--BoM_Button:hover .sapMBtnInner,
${B} #__pda_detail_rightcol__ .nd-bom:hover .sapMBtnInner { background:#f6f9ff !important; border-color:#9fc3ef !important; }
${B} #__pda_detail_rightcol__ #WorkcenterDetail--BoM_Button .sapMBtnInner::before,
${B} #__pda_detail_rightcol__ .nd-bom .sapMBtnInner::before { content:''; grid-column:1; grid-row:1 / span 2;
  width:54px; height:54px; border-radius:50%; background:#e6f0fc url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%231e5aa8' stroke-width='1.8' stroke-linecap='round' stroke-linejoin='round'><rect x='7.5' y='2.5' width='9' height='4' rx='1'/><path d='M16.5 4.5h1.5a2 2 0 0 1 2 2V11'/><path d='M7.5 4.5H6a2 2 0 0 0-2 2v13a2 2 0 0 0 2 2h6'/><path d='M8 11h6M8 15h4'/><circle cx='18' cy='18' r='2.2'/><path d='M18 14.6v1.2M18 20.2v1.2M14.6 18h1.2M20.2 18h1.2'/></svg>") center / 28px 28px no-repeat; }
${B} #__pda_detail_rightcol__ #WorkcenterDetail--BoM_Button .sapMBtnContent,
${B} #__pda_detail_rightcol__ .nd-bom .sapMBtnContent { grid-column:2 !important; grid-row:1 !important;
  font-size:18px !important; font-weight:800 !important; color:#13315c !important; }
${B} #__pda_detail_rightcol__ #WorkcenterDetail--BoM_Button .sapMBtnInner::after,
${B} #__pda_detail_rightcol__ .nd-bom .sapMBtnInner::after { content:'Zoznam komponentov'; grid-column:2; grid-row:2;
  font-size:12.5px; font-weight:500; color:#5b6b83; }
${B} #__pda_detail_rightcol__ #WorkcenterDetail--BoM_Button::after,
${B} #__pda_detail_rightcol__ .nd-bom::after { content:'›'; position:absolute; right:16px; top:50%; transform:translateY(-50%);
  font-size:26px; line-height:1; color:#13315c; pointer-events:none; }
/* povodna mala ikonka UI5 je nahradena velkou v kruhu */
${B} #__pda_detail_rightcol__ #WorkcenterDetail--BoM_Button .sapMBtnIcon,
${B} #__pda_detail_rightcol__ .nd-bom .sapMBtnIcon { display:none !important; }

/* --- zvisla deliaca ciara pred prepinacom Stroj --- */
${B} #__pda_detail_rightcol__ .pda-machine { border-left:1px solid rgba(110,140,185,.35) !important;
  padding-left:22px !important; margin-left:auto !important; align-self:stretch !important; }

/* ================= RAMY KARIET NA UVODNEJ OBRAZOVKE (3.2.7) =================
   Osobny stav, Vyhladat zakazku a Pracoviska maju rovnaky ram, plochu a tien ako
   ostatne karty. Hlavne pocas nacitavania boli bez ramu - appka ich prekryje
   "cakacim" zavojom a vyzerali ako vyblednute obdlzniky bez okraja. */
${B} .sapMPanel.pda-panel-tesny, ${B} #Main--Workcenter_Panel, ${B} #__pda_search_sidebar__ {
  background:var(--nd-karta) !important; border:1.5px solid #13315c !important;
  border-radius:18px !important; box-shadow:var(--nd-tien) !important;
  -webkit-backdrop-filter:blur(10px); backdrop-filter:blur(10px); }
${B} #__pda_search_sidebar__ { margin:10px 12px !important; width:calc(100% - 24px) !important;
  box-sizing:border-box !important; padding:4px 6px !important; }
/* cakaci zavoj appky (busy indikator): jemnejsi a so zaoblenim karty, aby karta
   pocas nacitavania neostala biela a bez tvaru */
${B} .sapUiLocalBusyIndicator { background-color:rgba(229,238,250,.38) !important; border-radius:inherit !important; }
${B} .sapMPanel.sapUiLocalBusy, ${B} .sapMPanel .sapUiLocalBusy { border-radius:inherit; }

/* ================= OKNA APPKY: jeden system pre vsetky (3.2.12) =================
   Novy dizajn robi vsetky nastrojove listy UI5 (.sapMTB) priehladne, aby cez ne
   presvitalo pozadie. Spodna lista okna s tlacidlami je tiez taka lista a samotne
   okno UI5 plochu nema (kresli ju zvlast hlavicka, obsah a lista) - cez listu tak
   presvitala stranka pod oknom a tlacidla Zavriet / Nacitat vsetko sa stracali.
   Kazde okno appky (BOM, potvrdenia, zaznamy, spravy, zmena pouzivatela, chyby...)
   ma teraz plnu bielu plochu, tenky tmavomodry ram ako karty, tmavomodru hlavicku
   s bielym nadpisom (ako Dokumentacia) a svetlu spodnu listu so zretelnymi
   tlacidlami. Vysky hlavicky a listy sa NEMENIA - UI5 im v okne vyhradzuje presne
   2.75rem (lista je polohovana absolutne), inak by prekryli obsah. */
${B} .sapMDialog:not(.sapMBusyDialog-Light) { background:#fff !important;
  border:1.5px solid #13315c !important; border-radius:16px !important;
  box-shadow:0 24px 70px rgba(16,36,63,.38) !important; }
${B} .sapMDialog.sapMDialogStretched { border-width:0 !important; border-radius:0 !important; }
/* hlavicka: tmavomodry pas, biely nadpis rovnakym pismom ako hlavicky kariet */
${B} .sapMDialog > header > .sapMDialogTitleGroup,
${B} .sapMDialog > header > .sapMDialogTitleGroup > .sapMIBar {
  background:#13315c !important; background-image:none !important; border:0 !important; box-shadow:none !important; }
${B} .sapMDialog > header > .sapMDialogTitleGroup .sapMTitle,
${B} .sapMDialog > header > .sapMDialogTitleGroup .sapMTitle > span {
  color:#fff !important; font-family:"Segoe UI",-apple-system,Roboto,sans-serif !important; font-size:16px !important;
  font-weight:600 !important; letter-spacing:.03em !important; text-transform:uppercase !important; text-shadow:none !important; }
${B} .sapMDialog > header > .sapMDialogTitleGroup .sapMLabel,
${B} .sapMDialog > header > .sapMDialogTitleGroup .sapMText,
${B} .sapMDialog > header > .sapMDialogTitleGroup .sapUiIcon { color:#fff !important; }
/* tlacidla vo vlastnej hlavicke okna (napr. krizik) - biele na tmavomodrom */
${B} .sapMDialog > header > .sapMDialogTitleGroup .sapMBtn .sapMBtnInner {
  background:transparent !important; border-color:rgba(255,255,255,.4) !important; }
${B} .sapMDialog > header > .sapMDialogTitleGroup .sapMBtn:hover .sapMBtnInner { background:rgba(255,255,255,.14) !important; }
${B} .sapMDialog > header > .sapMDialogTitleGroup .sapMBtn .sapMBtnContent,
${B} .sapMDialog > header > .sapMDialogTitleGroup .sapMBtn bdi,
${B} .sapMDialog > header > .sapMDialogTitleGroup .sapMBtn .sapUiIcon { color:#fff !important; }
/* okna s hlasenim (chyba, varovanie, uspech, informacia): farebna ikona a pruh pod hlavickou */
${B} .sapMDialogError > header > .sapMDialogTitleGroup > .sapMIBar, ${B} .sapMMessageBoxError > header > .sapMDialogTitleGroup > .sapMIBar { box-shadow:inset 0 -3px #ff6b6b !important; }
${B} .sapMDialogWarning > header > .sapMDialogTitleGroup > .sapMIBar, ${B} .sapMMessageBoxWarning > header > .sapMDialogTitleGroup > .sapMIBar { box-shadow:inset 0 -3px #ffb347 !important; }
${B} .sapMDialogSuccess > header > .sapMDialogTitleGroup > .sapMIBar, ${B} .sapMMessageBoxSuccess > header > .sapMDialogTitleGroup > .sapMIBar { box-shadow:inset 0 -3px #5fd08a !important; }
${B} .sapMDialogInformation > header > .sapMDialogTitleGroup > .sapMIBar, ${B} .sapMMessageBoxInformation > header > .sapMDialogTitleGroup > .sapMIBar { box-shadow:inset 0 -3px #7fb8ff !important; }
${B} .sapMDialogError > header > .sapMDialogTitleGroup .sapMDialogIcon, ${B} .sapMMessageBoxError > header > .sapMDialogTitleGroup .sapMDialogIcon { color:#ff8a8a !important; }
${B} .sapMDialogWarning > header > .sapMDialogTitleGroup .sapMDialogIcon, ${B} .sapMMessageBoxWarning > header > .sapMDialogTitleGroup .sapMDialogIcon { color:#ffc56e !important; }
${B} .sapMDialogSuccess > header > .sapMDialogTitleGroup .sapMDialogIcon, ${B} .sapMMessageBoxSuccess > header > .sapMDialogTitleGroup .sapMDialogIcon { color:#7ee2a4 !important; }
${B} .sapMDialogInformation > header > .sapMDialogTitleGroup .sapMDialogIcon, ${B} .sapMMessageBoxInformation > header > .sapMDialogTitleGroup .sapMDialogIcon { color:#a9d0ff !important; }
/* spodna lista: svetly pas s deliacou ciarou */
${B} .sapMDialog > footer, ${B} .sapMDialog > footer .sapMIBar {
  background:#eef3fa !important; background-image:none !important; }
${B} .sapMDialog > footer .sapMIBar { border-top:1px solid #c5d3e6 !important; box-shadow:none !important; }
/* tlacidla v spodnej liste (okno aj rozbalovacie okienko): biele s tmavomodrym ramom */
${LISTA(' .sapMBtn .sapMBtnInner')} {
  background:#fff !important; background-image:none !important; border:1px solid #13315c !important;
  border-radius:10px !important; box-shadow:0 1px 3px rgba(16,36,63,.14) !important;
  transition:background-color .12s ease, box-shadow .12s ease, transform .12s ease !important; }
${LISTA(' .sapMBtn .sapMBtnInner .sapMBtnContent')},
${LISTA(' .sapMBtn .sapMBtnInner bdi')},
${LISTA(' .sapMBtn .sapMBtnInner .sapUiIcon')} { color:#13315c !important; text-shadow:none !important; }
${LISTA(' .sapMBtn .sapMBtnInner .sapMBtnContent')} { font-weight:600 !important; }
/* pri prechode mysou sa tlacidlo vyplni a jemne nadvihne */
${LISTA(' .sapMBtn:not(.sapMBtnDisabled):hover .sapMBtnInner')} {
  background:#13315c !important; transform:translateY(-1px); box-shadow:0 5px 12px rgba(16,36,63,.26) !important; }
${LISTA(' .sapMBtn:not(.sapMBtnDisabled):hover .sapMBtnInner .sapMBtnContent')},
${LISTA(' .sapMBtn:not(.sapMBtnDisabled):hover .sapMBtnInner bdi')},
${LISTA(' .sapMBtn:not(.sapMBtnDisabled):hover .sapMBtnInner .sapUiIcon')} { color:#fff !important; }
${LISTA(' .sapMBtn .sapMBtnInner.sapMBtnActive')} { background:#0b2447 !important; transform:none; }
/* hlavna akcia (Emphasized) plne tmavomodra, suhlas zeleny, odmietnutie cervene */
${LISTA(' .sapMBtn .sapMBtnInner.sapMBtnEmphasized')} { background:#13315c !important; }
${LISTA(' .sapMBtn .sapMBtnInner.sapMBtnAccept')},
${LISTA(' .sapMBtn .sapMBtnInner.sapMBtnSuccess')} { background:#1d7a46 !important; border-color:#1d7a46 !important; }
${LISTA(' .sapMBtn .sapMBtnInner.sapMBtnReject')},
${LISTA(' .sapMBtn .sapMBtnInner.sapMBtnNegative')} { background:#b3261e !important; border-color:#b3261e !important; }
${LISTA(' .sapMBtn .sapMBtnInner.sapMBtnEmphasized *')},
${LISTA(' .sapMBtn .sapMBtnInner.sapMBtnAccept *')},
${LISTA(' .sapMBtn .sapMBtnInner.sapMBtnSuccess *')},
${LISTA(' .sapMBtn .sapMBtnInner.sapMBtnReject *')},
${LISTA(' .sapMBtn .sapMBtnInner.sapMBtnNegative *')} { color:#fff !important; }
${LISTA(' .sapMBtn:not(.sapMBtnDisabled):hover .sapMBtnInner.sapMBtnEmphasized')} { background:#1c478a !important; }
${LISTA(' .sapMBtn:not(.sapMBtnDisabled):hover .sapMBtnInner.sapMBtnAccept')},
${LISTA(' .sapMBtn:not(.sapMBtnDisabled):hover .sapMBtnInner.sapMBtnSuccess')} { background:#16613a !important; }
${LISTA(' .sapMBtn:not(.sapMBtnDisabled):hover .sapMBtnInner.sapMBtnReject')},
${LISTA(' .sapMBtn:not(.sapMBtnDisabled):hover .sapMBtnInner.sapMBtnNegative')} { background:#8f1d17 !important; }
/* rozbalovacie okienka (vyber zo zoznamu, datum, cas): rovnaky tenky tmavomodry obrys.
   Obrys je tien (tak ho kresli aj tema), takze sa nemeni ziadny rozmer. */
${B} .sapMPopover { box-shadow:0 0 0 1px #13315c, 0 10px 30px rgba(16,36,63,.28) !important; }
${B} .sapMPopoverArr::after { box-shadow:0 0 0 1px #13315c !important; }

/* vlastne okna skriptu: rovnaky ram a hlavicka ako okna appky */
${B} #__pda_opis_overlay__ .karta, ${B} #__pda_hf_overlay__ .karta, ${B} #__pda_settings_pass__ > div,
${B} .pda-set-box { border:1.5px solid #13315c !important; border-radius:16px !important; }
${B} .pda-set-head { background:#13315c !important; }
/* okno Vykresy (zoznam z PDM): tmavomodra hlavicka a akcenty namiesto svetlomodrej */
${B} #__pda_pdm_overlay__ > div { border:1.5px solid #13315c !important; border-radius:16px !important;
  box-shadow:0 24px 70px rgba(16,36,63,.38) !important; }
${B} #__pda_pdm_overlay__ > div > div:first-child { background:#13315c !important; padding:14px 18px !important; }
${B} #__pda_pdm_overlay__ > div > div:first-child > b { font-family:"Segoe UI",-apple-system,Roboto,sans-serif !important;
  font-size:16px !important; font-weight:600 !important; letter-spacing:.03em !important; }
${B} #__pda_pdm_overlay__ th { border-bottom-color:#13315c !important; color:#13315c !important; }
${B} #__pda_pdm_overlay__ [data-pda-live] { background:#13315c !important; }

/* ================= RAMY VNUTORNYCH KARIET: jemna tmavomodra (3.2.12) =================
   Aj jednotlive karty vnutri sekcii (SAP casy, paralelne procesy, dlazdice
   Dokumentacie, tlacidla panela PDA, Graf / Tabulka, stavove tlacidla) maju tenky
   tmavomodry ram. Je 1 px - sekcie maju 1.5 px, aby bolo vidno, co je v com. */
${B} ${CASY} > .sapMVBox, ${B} ${ZOZNAM} .sapMLIB, ${B} ${ZOZNAM} .pda-aktivita,
${B} #__pda_hf_menu__ .hf-btn { border:1px solid #13315c !important; }
${B} #__pda_detail_rightcol__ #__pda_order_drawing_button__,
${B} #__pda_detail_rightcol__ #__pda_order_drawing_button__:hover,
${B} #__pda_detail_rightcol__ #WorkcenterDetail--BoM_Button .sapMBtnInner,
${B} #__pda_detail_rightcol__ #WorkcenterDetail--BoM_Button:hover .sapMBtnInner,
${B} #__pda_detail_rightcol__ .nd-bom .sapMBtnInner, ${B} #__pda_detail_rightcol__ .nd-bom:hover .sapMBtnInner,
${B} #__pda_left_box_graf__ #WorkcenterDetail--Buttons_FlexBox .sapMBtnInner,
${B} #__pda_left_box_graf__ #WorkcenterDetail--Buttons_FlexBox .sapMBtn:hover .sapMBtnInner {
  border:1px solid #13315c !important; }
${B} #WorkcenterDetail--Order_Status_Flexbox .sapMBtn, ${B} .statusBtn { border-color:#13315c !important; }

/* ================= SAP CASY: priestorovy kolac a pohyb pod mysou (3.2.12) =================
   Karta s kolacom sa pri prechode mysou nadvihne ako tlacidla a kolac sa este
   kusok zdvihne nad svoj tien. Kolac je nas obrazok (nakloneny prstenec s
   vyskou) - kolac appky pod nim ostava, len je neviditelny. Pohyb je len
   transform/opacity, ziadne filtre (tie v 1.20 mrazili appku). */
${B} ${CASY} > .sapMVBox { transition:transform .18s ease, box-shadow .18s ease !important; }
${B} ${CASY} > .sapMVBox:hover { transform:translateY(-3px) !important;
  box-shadow:0 10px 20px rgba(16,36,63,.20) !important; }
${B} ${CASY} canvas.pda-3j-skryty { opacity:0 !important; filter:none !important; }
${B} ${CASY} .pda-3j-kolac { position:absolute !important; z-index:1 !important; pointer-events:none !important; }
${B} ${CASY} .pda-3j-kolac::before, ${B} ${CASY} .pda-3j-kolac::after { content:''; position:absolute;
  pointer-events:none; transition:transform .2s ease, opacity .2s ease; }
/* tien na podlozke (mäkka elipsa, bez filtra) */
${B} ${CASY} .pda-3j-kolac::before { left:9%; width:82%; top:74%; height:14%; border-radius:50%;
  background:radial-gradient(closest-side, rgba(16,36,63,.32), rgba(16,36,63,.12) 62%, rgba(16,36,63,0)); }
${B} ${CASY} .pda-3j-kolac::after { inset:0; background:var(--k3-img) center / 100% 100% no-repeat; }
${B} ${CASY} > .sapMVBox:hover .pda-3j-kolac::after { transform:translateY(-6px); }
${B} ${CASY} > .sapMVBox:hover .pda-3j-kolac::before { transform:scale(.88); opacity:.65; }
${B} ${CASY} > .sapMVBox:hover .nd-pct { transform:translate(-50%, calc(-50% + var(--k3-dy, 0px) - 6px)) !important; }
/* ================= STAVOVE TLACIDLA: vyraznejsie obrazky v pozadi (3.2.13) =================
   Obrazky v pozadi (Stretnutia, Prestavka, Cakanie, Vyroba, Udrzba, ...) su o 30 %
   viditelnejsie: priehladnost .5 -> .65. Tlacidlo samotne sa nemeni. */
${B} .statusBtn[data-pda-obraz]::after { opacity:.65 !important; }
/* ================= UVOD: bez tlacidla "Hladat vyrobny prikaz" (3.2.13) =================
   Na pokyn pouzivatela sa tlacidlo v hlavicke karty Pracoviska nezobrazuje (zakazky sa
   hladaju v karte Vyhladat zakazku). Len sa skryje - v appke ostava, nic sa nemaze. */
${B} #Main--Button_SearchProductionOrder { display:none !important; }
/* ================= PANELY OSOBNY STAV A PRACOVISKO: nedaju sa zbalit (3.2.13) =================
   Rovnako ako panel Pracoviska: sipka na zbalenie je skryta a obsah panela je vzdy
   zobrazeny, aj keby ho UI5 po kliknuti na hlavicku vnutorne zbalilo (display/height
   s !important prebiju inline styl UI5). Len CSS - nic sa neklika ani neprepina. */
${B} #Main--Tasks_Panel .sapMPanelExpandableIcon,
${B} #WorkcenterDetail--Tasks_Panel .sapMPanelExpandableIcon,
${B} #WorkcenterDetail--Planned_WorkList .sapMPanelExpandableIcon,
${B} #Main--Tasks_Panel [id$="-expandButton"],
${B} #WorkcenterDetail--Tasks_Panel [id$="-expandButton"],
${B} #WorkcenterDetail--Planned_WorkList [id$="-expandButton"],
${B} #Main--Tasks_Panel [id$="-CollapsedImg"],
${B} #WorkcenterDetail--Tasks_Panel [id$="-CollapsedImg"],
${B} #WorkcenterDetail--Planned_WorkList [id$="-CollapsedImg"],
/* panel s obycajnou hlavickou (Osobny stav) ma sipku ako samostatnu ikonu v obale hlavicky */
${B} #Main--Tasks_Panel > .sapMPanelWrappingDiv > .sapUiIcon,
${B} #WorkcenterDetail--Tasks_Panel > .sapMPanelWrappingDiv > .sapUiIcon,
${B} #WorkcenterDetail--Planned_WorkList > .sapMPanelWrappingDiv > .sapUiIcon { display:none !important; }
${B} #Main--Tasks_Panel-content,
${B} #WorkcenterDetail--Tasks_Panel-content,
${B} #WorkcenterDetail--Planned_WorkList-content,
${B} #Main--Tasks_Panel > .sapMPanelContent,
${B} #WorkcenterDetail--Tasks_Panel > .sapMPanelContent,
${B} #WorkcenterDetail--Planned_WorkList > .sapMPanelContent { display:block !important; height:auto !important; }
${B} #Main--Tasks_Panel .sapMPanelHdrExpandable,
${B} #WorkcenterDetail--Tasks_Panel .sapMPanelHdrExpandable,
${B} #WorkcenterDetail--Planned_WorkList .sapMPanelHdrExpandable,
${B} #Main--Tasks_Panel .sapMPanelWrappingDivTb,
${B} #WorkcenterDetail--Tasks_Panel .sapMPanelWrappingDivTb,
${B} #WorkcenterDetail--Planned_WorkList .sapMPanelWrappingDivTb { cursor:default !important; }

/* ================= DETAIL: nadpis pracoviska len cislo a nazov (3.2.13) =================
   Namiesto "Arbeitsplatz: 5388 - PORTALKA FG 3010 CNC" je len "5388 - PORTALKA FG 3010 CNC",
   o 50 % vacsie (16 -> 24 px) a tucne. Text appky ostava netknuty (len sa nezobrazi),
   nas text je pseudo-prvok z atributu data-pda-nazov, ktory dopisuje JS. */
${B} #WorkcenterDetail--Planned_Worklist_Title.pda-nazov-pracoviska,
${B} #WorkcenterDetail--Planned_Worklist_Title.pda-nazov-pracoviska > * { font-size:0 !important; }
${B} #WorkcenterDetail--Planned_Worklist_Title.pda-nazov-pracoviska::after { content:attr(data-pda-nazov);
  font-family:"Segoe UI",-apple-system,Roboto,sans-serif; font-size:24px; font-weight:700; line-height:1.3;
  color:#13315c; letter-spacing:0; }
/* ================= VYKRES a COMPONENTS / BOM: pohyb pod mysou ako tlacidla (3.2.13) =================
   Obe dlazdice su tlacidla, preto sa pri prechode mysou nadvihnu a dostanu vyraznejsi
   tien rovnako ako ostatne tlacidla (napr. Popis operacie). Len transform a tien. */
${B} #__pda_detail_rightcol__ #__pda_order_drawing_button__,
${B} #__pda_detail_rightcol__ #WorkcenterDetail--BoM_Button, ${B} #__pda_detail_rightcol__ .nd-bom {
  transition:transform .13s ease, box-shadow .13s ease, background-color .13s, border-color .13s !important; }
${B} #__pda_detail_rightcol__ #WorkcenterDetail--BoM_Button .sapMBtnInner,
${B} #__pda_detail_rightcol__ .nd-bom .sapMBtnInner { transition:box-shadow .13s ease, background-color .13s !important; }
${B} #__pda_detail_rightcol__ #__pda_order_drawing_button__:hover,
${B} #__pda_detail_rightcol__ #WorkcenterDetail--BoM_Button:hover, ${B} #__pda_detail_rightcol__ .nd-bom:hover {
  transform:translateY(-3px) !important; }
${B} #__pda_detail_rightcol__ #__pda_order_drawing_button__:hover,
${B} #__pda_detail_rightcol__ #WorkcenterDetail--BoM_Button:hover .sapMBtnInner,
${B} #__pda_detail_rightcol__ .nd-bom:hover .sapMBtnInner { box-shadow:0 10px 20px rgba(16,36,63,.22) !important; }
${B} #__pda_detail_rightcol__ #__pda_order_drawing_button__:active,
${B} #__pda_detail_rightcol__ #WorkcenterDetail--BoM_Button:active, ${B} #__pda_detail_rightcol__ .nd-bom:active {
  transform:translateY(-1px) !important; }
`;
            document.head.appendChild(st);
        }

        // Velke panely stranky (Osobny stav, pracovisko) su len obaly - bez vlastnej
        // plochy, aby medzi kartami presvitalo pozadie. Karty v nich ostavaju biele.
        /*
         * Priehladny obal (bez vlastnej plochy) je uz len panel pracoviska v detaile -
         * v nom su vsetky karty. Panel Osobny stav je od 3.2.7 samostatna karta s ramom
         * (pocas nacitavania bez ramu vyzeral ako vyblednuty obdlznik), preto sa mu
         * priehladnost z predoslych verzii odoberie.
         */
        function priehladnePanely() {
            const left = document.getElementById('WorkcenterDetail--LeftColumn_FlexBox');
            const detail = left && left.closest('.sapMPanel');
            if (detail && !detail.classList.contains('pda-3j-priehladny')) detail.classList.add('pda-3j-priehladny');
            document.querySelectorAll('.pda-panel-tesny.pda-3j-priehladny').forEach((p) => {
                if (p !== detail) p.classList.remove('pda-3j-priehladny');
            });
        }

        /*
         * Uvodna obrazovka: farba a ikona pre kazdu oblast pracovisk.
         * Styri zname oblasti maju pevnu farbu; oblast, ktora sa objavi navyse,
         * dostane dalsiu nahradnu farbu v poradi, v akom je v prehlade - dve nove
         * oblasti tak nikdy nemaju rovnaku farbu. Ikony su jednoduche obrysy
         * (styl Lucide), kreslene farbou oblasti (currentColor).
         * Nastavuje sa len premenna --ov-farba a obsah dlazdice ikony; na
         * prepis sa siaha, len ked sa oblast zmenila (data-v3j).
         */
        const FARBY_OBLASTI = {
            'Assembly': '#1b9bd7',
            'Welding': '#e3782a',
            'Machining': '#2c5fc6',
            'Quality Control': '#189a6e',
        };
        const NAHRADNE_FARBY = ['#8b5cd6', '#d64f7a', '#c79a1e', '#3b8fa0', '#b0602f', '#6b7c95'];

        const SVG = (vnutro) => '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" ' +
            'stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' + vnutro + '</svg>';
        const IKONY_OBLASTI = {
            // klúč
            'Assembly': SVG('<path d="M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77a6 6 0 0 1-7.94 7.94' +
                'l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94l-3.76 3.76z"/>'),
            // plamen (vyplneny)
            'Welding': SVG('<path fill="currentColor" stroke="none" d="M8.5 14.5A2.5 2.5 0 0 0 11 12c0-1.38-.5-2-1-3' +
                '-1.07-2.14-.22-4.05 2-6 .5 2.5 2 4.9 4 6.5 2 1.6 3 3.5 3 5.5a7 7 0 1 1-14 0c0-1.15.43-2.29 1-3' +
                'a2.5 2.5 0 0 0 2.5 2.5z"/>'),
            // ozubene koleso (vyplnene, stred priesvitny)
            'Machining': SVG('<path fill="currentColor" stroke="none" fill-rule="evenodd" d="M12.22 2h-.44a2 2 0 0 0-2 2' +
                'v.18a2 2 0 0 1-1 1.73l-.43.25a2 2 0 0 1-2 0l-.15-.08a2 2 0 0 0-2.73.73l-.22.38a2 2 0 0 0 .73 2.73' +
                'l.15.1a2 2 0 0 1 1 1.72v.51a2 2 0 0 1-1 1.74l-.15.09a2 2 0 0 0-.73 2.73l.22.38a2 2 0 0 0 2.73.73' +
                'l.15-.08a2 2 0 0 1 2 0l.43.25a2 2 0 0 1 1 1.73V20a2 2 0 0 0 2 2h.44a2 2 0 0 0 2-2v-.18a2 2 0 0 1 1-1.73' +
                'l.43-.25a2 2 0 0 1 2 0l.15.08a2 2 0 0 0 2.73-.73l.22-.39a2 2 0 0 0-.73-2.73l-.15-.08a2 2 0 0 1-1-1.74' +
                'v-.5a2 2 0 0 1 1-1.74l.15-.09a2 2 0 0 0 .73-2.73l-.22-.38a2 2 0 0 0-2.73-.73l-.15.08a2 2 0 0 1-2 0' +
                'l-.43-.25a2 2 0 0 1-1-1.73V4a2 2 0 0 0-2-2zM12 15.2a3.2 3.2 0 1 0 0-6.4 3.2 3.2 0 0 0 0 6.4z"/>'),
            // lupa
            'Quality Control': SVG('<circle cx="11" cy="11" r="7"/><path d="m21 21-5-5" stroke-width="2.6"/>'),
        };
        // ina oblast: krabica
        const IKONA_INA = SVG('<path d="M21 8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73' +
            'l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16Z"/><path d="m3.3 7 8.7 5 8.7-5"/><path d="M12 22V12"/>');

        function farbyPrehladu() {
            const skupiny = document.querySelectorAll('#__pda_overview__ .ov-g');
            if (!skupiny.length) return;
            let nahradna = 0;
            skupiny.forEach((g) => {
                const b = g.querySelector('.ov-gt b');
                if (!b) return;
                const meno = (b.textContent || '').trim();
                const farba = FARBY_OBLASTI[meno] || NAHRADNE_FARBY[nahradna++ % NAHRADNE_FARBY.length];
                if (g.style.getPropertyValue('--ov-farba') !== farba) g.style.setProperty('--ov-farba', farba);
                const ic = g.querySelector('.ov-ic');
                if (ic && ic.dataset.v3j !== meno) {
                    ic.innerHTML = IKONY_OBLASTI[meno] || IKONA_INA;
                    ic.dataset.v3j = meno;
                }
            });
        }

        /*
         * Prihlasovacia stranka: tlacidlo Azure a odkaz "meno a heslo" sa oznacia
         * triedou aj podla TEXTU a toho, ze patria stranke Login - nielen podla ID.
         * Keby mali na stranke ine ID nez v zdrojakoch, styl sa aj tak chyti.
         * Len trieda (atribut), nic sa neklika ani nepresuva.
         */
        function oznacPrihlasenie() {
            document.querySelectorAll('.sapMBtn').forEach((b) => {
                if (b.classList.contains('pda-login-azure')) return;
                const naLogine = (b.id || '').indexOf('Login--') === 0 || b.closest('[id^="Login--"]');
                if (naLogine && /azure/i.test(b.textContent || '')) b.classList.add('pda-login-azure');
            });
            document.querySelectorAll('.sapMLnk').forEach((a) => {
                if (a.classList.contains('pda-login-link')) return;
                const naLogine = (a.id || '').indexOf('Login') === 0 || a.closest('[id^="Login"]');
                if (naLogine) a.classList.add('pda-login-link');
            });
        }

        /*
         * SAP CASY: priestorovy kolac s vyskou (3.2.12).
         *
         * Kolac appky (Chart.js platno) ostava v stranke a appka ho dalej kresli,
         * je len neviditelny. Nad nim lezi nas obrazok: nakloneny prstenec s bocnou
         * stenou (vyska), tienovanim a leskom - farby sedia s legendou pod kolacom
         * (Hotovo zelena, Zostava sivomodra). Percento sa pocita z textu casov appky
         * rovnako ako cislo v strede kolaca, takze data sa nemenia.
         *
         * Obrazok je SVG vlozene ako pozadie <div>, NIE prvok <svg>: ine moduly
         * hladaju kolac ako prvy <canvas>/<svg> v boxe a nas prvok ich nesmie
         * zmiast. Prekresluje sa len pri zmene percenta (cele percenta); ziadne
         * filtre - drop-shadow v 1.20 mrazil celu appku.
         */
        const KOLAC_KOTVY = ['WorkcenterDetail--SetupTime_Text', 'WorkcenterDetail--MachineTime_Text',
                             'WorkcenterDetail--LaborTime_Text'];
        // tvar v jednotkach obrazka 200 x 200: stred hornej plochy, polomery, sklon (k) a vyska steny (h)
        const K3 = { cx: 100, cy: 83, R: 88, r: 57, k: 0.68, h: 18 };
        const K3_FARBY = [
            { hore: '#2fa65a', stena: '#1f7c41', dnu: '#17602f' },   // hotovo (zelena)
            { hore: '#c9d4e2', stena: '#9eaec4', dnu: '#8193ad' },   // zostava
        ];

        function k3Sekundy(t) {
            t = String(t || '').trim();
            if (/:/.test(t)) return t.split(':').map(Number).reduce((a, b) => a * 60 + (isNaN(b) ? 0 : b), 0);
            const n = parseFloat(t.replace(',', '.'));
            return isNaN(n) ? 0 : n;
        }

        function kolacSvg(pct) {
            const { cx, cy, R, r, k, h } = K3;
            const PI = Math.PI, ZAC = -PI / 2, CELY = 2 * PI;
            // bod na elipse (uhol a: 0 = vpravo, PI/2 = vpredu dole), dy = posun nadol
            const bod = (rad, a, dy) => (cx + rad * Math.cos(a)).toFixed(2) + ' ' +
                (cy + rad * k * Math.sin(a) + (dy || 0)).toFixed(2);
            const oblukDo = (rad, a, b, dy, smer) => 'A' + rad + ' ' + (rad * k).toFixed(2) + ' 0 ' +
                (Math.abs(b - a) > PI ? 1 : 0) + ' ' + smer + ' ' + bod(rad, b, dy);
            // horna plocha useku medzikruzia; cely kruh po polovicach (jeden oblúk ho nenakresli)
            const plocha = (a, b) => {
                if (b - a > CELY - 1e-4) return plocha(a, a + PI) + plocha(a + PI, b);
                return 'M' + bod(R, a) + oblukDo(R, a, b, 0, 1) + 'L' + bod(r, b) + oblukDo(r, b, a, 0, 0) + 'Z';
            };
            // stena useku [a, b], len jej viditelna cast [od, po]
            const stena = (rad, a, b, od, po) => {
                const x = Math.max(a, od), y = Math.min(b, po);
                if (y - x < 1e-4) return '';
                return 'M' + bod(rad, x) + oblukDo(rad, x, y, 0, 1) + 'L' + bod(rad, y, h) +
                    oblukDo(rad, y, x, h, 0) + 'Z';
            };
            const podiel = Math.max(0, Math.min(100, pct)) / 100;
            const hranica = ZAC + CELY * podiel;
            const useky = [];
            if (podiel > 0.0005) useky.push([ZAC, hranica, K3_FARBY[0]]);
            if (podiel < 0.9995) useky.push([hranica, ZAC + CELY, K3_FARBY[1]]);

            let t = '';
            // 1. vonkajsia stena - vidno len prednu polovicu (uhly 0 az PI)
            useky.forEach(([a, b, f]) => {
                const d = stena(R, a, b, 0, PI);
                if (d) t += '<path d="' + d + '" fill="' + f.stena + '"/>';
            });
            t += '<path d="' + stena(R, 0, PI, 0, PI) + '" fill="url(#t)"/>';
            // 2. vnutorna stena - cez otvor vidno jej zadnu polovicu
            useky.forEach(([a, b, f]) => {
                const d = stena(r, a, b, ZAC, 0) + stena(r, a, b, PI, ZAC + CELY);
                if (d) t += '<path d="' + d + '" fill="' + f.dnu + '"/>';
            });
            // 3. horna plocha s jemnym leskom
            useky.forEach(([a, b, f]) => { t += '<path d="' + plocha(a, b) + '" fill="' + f.hore + '"/>'; });
            t += '<path d="' + plocha(ZAC, ZAC + CELY) + '" fill="url(#g)"/>';
            // 4. biele deliace ciary medzi usekmi (ako okraje kolaca appky)
            if (useky.length > 1) {
                [ZAC, hranica].forEach((a) => {
                    t += '<path d="M' + bod(r, a) + 'L' + bod(R, a) + '" stroke="#fff" stroke-width="1.6"/>';
                    if (Math.sin(a) > 0.02) {
                        t += '<path d="M' + bod(R, a) + 'L' + bod(R, a, h) + '" stroke="#fff" stroke-opacity=".7" stroke-width="1.2"/>';
                    }
                });
            }
            // 5. hrany: svetla horna, tmavsia spodna
            t += '<ellipse cx="' + cx + '" cy="' + cy + '" rx="' + R + '" ry="' + (R * k).toFixed(2) +
                '" fill="none" stroke="#fff" stroke-opacity=".55" stroke-width="1"/>';
            t += '<ellipse cx="' + cx + '" cy="' + cy + '" rx="' + r + '" ry="' + (r * k).toFixed(2) +
                '" fill="none" stroke="#0b2447" stroke-opacity=".18" stroke-width="1"/>';
            t += '<path d="M' + bod(R, 0, h) + oblukDo(R, 0, PI, h, 1) + '" fill="none" stroke="#0b2447" stroke-opacity=".22" stroke-width="1"/>';

            return '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 200"><defs>' +
                '<linearGradient id="t" x1="0" x2="1" y1="0" y2="0">' +
                '<stop offset="0" stop-color="#000" stop-opacity=".30"/><stop offset=".32" stop-color="#000" stop-opacity=".04"/>' +
                '<stop offset=".56" stop-color="#fff" stop-opacity=".12"/><stop offset="1" stop-color="#000" stop-opacity=".32"/>' +
                '</linearGradient>' +
                '<linearGradient id="g" x1="0" x2="0" y1="0" y2="1">' +
                '<stop offset="0" stop-color="#fff" stop-opacity=".42"/><stop offset=".5" stop-color="#fff" stop-opacity=".08"/>' +
                '<stop offset="1" stop-color="#000" stop-opacity=".06"/>' +
                '</linearGradient></defs>' + t + '</svg>';
        }

        function kolace3D() {
            KOLAC_KOTVY.forEach((id) => {
                const text = document.getElementById(id);
                if (!text) return;
                const box = text.closest('.sapMVBox') || text.closest('.sapMFlexBox') || text.parentElement;
                if (!box || !box.closest(CASY)) return;
                let canvas = null;
                box.querySelectorAll('canvas').forEach((g) => {
                    if (!canvas && g.id !== 'ResourceDetails' && g.id !== 'DialogChart') canvas = g;
                });
                if (!canvas || !canvas.parentElement) return;
                const par = canvas.parentElement;
                let el = null;
                for (const ch of par.children) { if (ch.classList.contains('pda-3j-kolac')) { el = ch; break; } }

                const casti = (text.textContent || '').split('/');
                const w = canvas.offsetWidth, v = canvas.offsetHeight;
                if (casti.length < 2 || w < 60 || v < 60) {
                    // neda sa precitat - ostane kolac appky
                    if (el) el.remove();
                    canvas.classList.remove('pda-3j-skryty');
                    return;
                }
                const plan = k3Sekundy(casti[1]);
                const pct = plan > 0 ? k3Sekundy(casti[0]) / plan * 100 : 0;

                if (par.style.position !== 'relative') par.style.position = 'relative';
                if (!el) {
                    el = document.createElement('div');
                    el.className = 'pda-3j-kolac';
                    el.setAttribute('aria-hidden', 'true');
                    par.insertBefore(el, canvas.nextSibling);
                }
                if (!canvas.classList.contains('pda-3j-skryty')) canvas.classList.add('pda-3j-skryty');

                const kluc = String(Math.round(Math.max(0, Math.min(100, pct))));
                if (el.dataset.k !== kluc) {
                    el.dataset.k = kluc;
                    el.style.setProperty('--k3-img', 'url("data:image/svg+xml,' + encodeURIComponent(kolacSvg(+kluc)) + '")');
                }
                const poloha = { left: canvas.offsetLeft + 'px', top: canvas.offsetTop + 'px', width: w + 'px', height: v + 'px' };
                Object.keys(poloha).forEach((q) => { if (el.style[q] !== poloha[q]) el.style[q] = poloha[q]; });

                // cislo v strede patri do viditelnej casti otvoru nakloneneho kolaca - pod zadnou
                // vnutornou stenou, teda o kusok inde nez je stred platna
                let lbl = null;
                for (const ch of par.children) { if (ch.classList.contains('nd-pct')) { lbl = ch; break; } }
                if (lbl) {
                    const dy = ((K3.cy + K3.h / 2 - 100) / 200 * v).toFixed(1) + 'px';
                    if (lbl.style.getPropertyValue('--k3-dy') !== dy) lbl.style.setProperty('--k3-dy', dy);
                }
            });
        }

        // detail: nadpis pracoviska bez "Arbeitsplatz:" - len cislo a nazov (zobrazuje ho CSS z atributu)
        function nadpisPracoviska() {
            const t = document.getElementById('WorkcenterDetail--Planned_Worklist_Title');
            if (!t) return;
            const cisty = (t.textContent || '').replace(/^\s*Arbeitsplatz\s*:\s*/i, '').replace(/\s+/g, ' ').trim();
            if (!cisty) return;
            if (t.getAttribute('data-pda-nazov') !== cisty) t.setAttribute('data-pda-nazov', cisty);
            if (!t.classList.contains('pda-nazov-pracoviska')) t.classList.add('pda-nazov-pracoviska');
        }

        // v patke nech je jasne, ze bezi verzia 3J
        function patka() {
            const r = document.querySelector('#__pda_nd_footer__ .nd-f-r');
            if (!r) return;
            let v = '';
            try { v = GM_info && GM_info.script ? GM_info.script.version : ''; } catch (e) { /* ignore */ }
            const t = 'PDA App Extension · verzia 3J' + (v ? ' (' + v + ')' : '') + ' · HF Slovakia';
            if (r.textContent !== t) r.textContent = t;
        }

        function apply() {
            injectStyles();
            [priehladnePanely, patka, farbyPrehladu, oznacPrihlasenie, kolace3D, nadpisPracoviska].forEach((f) => {
                try { f(); } catch (e) { console.warn(LOG, 'verzia 3J:', f.name, e); }
            });
        }

        DomWatch.add(apply);
        onReady(apply);
        // pri zmene velkosti okna sa kolac presunie spolu s platnom appky
        W.addEventListener('resize', () => { try { kolace3D(); } catch (e) { /* ignore */ } });
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
