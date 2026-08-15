/**
 * Finance engine: rule matching, statement splicing, workbook read/write.
 *
 * Deliberately free of DOM code so the dashboard can use the same categorisation
 * the sync page uses.  Mirrors tools/make_dummy_data.py - if one changes, so does
 * the other.
 */
const Engine = (function () {
    'use strict';

    const MONTHS = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN',
                    'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];

    // ---------------------------------------------------------------- text

    function normalise(s) {
        return String(s === null || s === undefined ? '' : s)
            .toUpperCase().split(/\s+/).filter(Boolean).join(' ');
    }

    // ---------------------------------------------------------------- dates

    /** Accepts Date, Excel serial, "10-Aug-26", "10/08/2026", "2026-08-10". */
    function parseDate(v) {
        if (v === null || v === undefined || v === '') return null;
        if (v instanceof Date) return isNaN(v) ? null : v;

        if (typeof v === 'number') {                    // Excel serial
            const ms = Math.round((v - 25569) * 86400 * 1000);
            const d = new Date(ms);
            return isNaN(d) ? null : new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
        }

        const s = String(v).trim();
        let m = s.match(/^(\d{1,2})[-\/\s]([A-Za-z]{3})[a-z]*[-\/\s](\d{2,4})$/);
        if (m) {
            const mon = MONTHS.indexOf(m[2].toUpperCase());
            let year = parseInt(m[3], 10);
            if (year < 100) year += 2000;
            if (mon >= 0) return new Date(year, mon, parseInt(m[1], 10));
        }
        m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
        if (m) return new Date(+m[1], +m[2] - 1, +m[3]);

        m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
        if (m) {                                        // day first - UK
            let year = parseInt(m[3], 10);
            if (year < 100) year += 2000;
            return new Date(year, +m[2] - 1, +m[1]);
        }
        const d = new Date(s);
        return isNaN(d) ? null : d;
    }

    function fmtDate(d) {
        if (!d) return '';
        const day = String(d.getDate()).padStart(2, '0');
        const mon = MONTHS[d.getMonth()];
        return day + '-' + mon.charAt(0) + mon.slice(1).toLowerCase() + '-' +
               String(d.getFullYear()).slice(2);
    }

    /**
     * Times come back three ways and all three must compare equal:
     *   "11:02"                     text, as a CSV gives it
     *   Date 1899-12-30T11:02       Excel time cell, via SheetJS cellDates
     *   0.4597...                   Excel time cell as a raw fraction of a day
     * Everything is reduced to "HH:MM".
     */
    function fmtTime(v) {
        if (v === null || v === undefined || v === '') return '';
        const pad = function (n) { return String(n).padStart(2, '0'); };
        if (v instanceof Date) {
            // A time-only cell sits on Excel's 1899-12-30 epoch; read it as UTC
            // so a timezone can never shift it. Real datetimes read locally.
            return v.getFullYear() < 1901
                ? pad(v.getUTCHours()) + ':' + pad(v.getUTCMinutes())
                : pad(v.getHours()) + ':' + pad(v.getMinutes());
        }
        if (typeof v === 'number') {
            const mins = Math.round((v % 1) * 24 * 60);
            return pad(Math.floor(mins / 60) % 24) + ':' + pad(mins % 60);
        }
        const m = String(v).trim().match(/^(\d{1,2}):(\d{2})/);
        return m ? pad(+m[1]) + ':' + pad(+m[2]) : String(v).trim();
    }

    function dayKey(d) {
        return d ? d.getFullYear() * 10000 + (d.getMonth() + 1) * 100 + d.getDate() : 0;
    }

    function num(v) {
        if (typeof v === 'number') return v;
        if (v === null || v === undefined || v === '') return null;
        const n = parseFloat(String(v).replace(/[£$€,\s]/g, '').replace(/^\((.*)\)$/, '-$1'));
        return isNaN(n) ? null : n;
    }

    // ---------------------------------------------------------------- rules

    function wildcardToRegex(pattern) {
        let out = '^';
        for (const ch of pattern) {
            if (ch === '*') out += '[\\s\\S]*';
            else if (ch === '?') out += '[\\s\\S]';
            else out += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        }
        return new RegExp(out + '$');
    }

    /** How much real text a rule pins down - this is what specificity means. */
    function literalLength(pattern, kind) {
        if (kind === 'wildcard') return pattern.replace(/[*?]/g, '').length;
        return pattern.length;
    }

    /**
     * Higher wins.  Sheet order is never consulted.
     *   exact     1,000,000
     *   wildcard  literals, +1 per anchored end (no leading/trailing *)
     *   contains  literals (identical to the fully open wildcard *X*)
     */
    function ruleScore(rule) {
        const kind = rule.type;
        const m = normalise(rule.match);
        if (kind === 'exact') return 1000000;
        let score = literalLength(m, kind);
        if (kind === 'wildcard') {
            score += m.startsWith('*') ? 0 : 1;
            score += m.endsWith('*') ? 0 : 1;
        }
        return score;
    }

    function compileRule(rule) {
        const r = Object.assign({}, rule);
        r.score = ruleScore(r);
        r._norm = normalise(r.match);
        if (r.type === 'wildcard') r._re = wildcardToRegex(r._norm);
        r.from = r.from ? parseDate(r.from) : null;
        r.to = r.to ? parseDate(r.to) : null;
        return r;
    }

    function compileRules(rules) { return rules.map(compileRule); }

    function ruleApplies(rule, desc, date) {
        if (rule.from && date && dayKey(date) < dayKey(rule.from)) return false;
        if (rule.to && date && dayKey(date) > dayKey(rule.to)) return false;
        const d = normalise(desc);
        switch (rule.type) {
            case 'exact':    return d === rule._norm;
            case 'contains': return rule._norm !== '' && d.indexOf(rule._norm) !== -1;
            case 'wildcard': return rule._re ? rule._re.test(d) : false;
            default:         return false;
        }
    }

    /**
     * Highest-scoring applicable rule, the runner-up it beat, and any
     * equal-scoring disagreement.  The runner-up is kept so the UI can say WHY
     * a category was chosen - opacity is the single biggest complaint levelled
     * at every rules engine of this kind.
     */
    function bestRule(desc, date, rules) {
        let best = null, runnerUp = null, conflict = null;
        for (const rule of rules) {
            if (!ruleApplies(rule, desc, date)) continue;
            if (!best || rule.score > best.score) {
                runnerUp = best; best = rule; conflict = null;
            } else if (rule.score === best.score && rule.category !== best.category) {
                conflict = rule;
            } else if (!runnerUp || rule.score > runnerUp.score) {
                runnerUp = rule;
            }
        }
        return { rule: best, runnerUp: runnerUp, conflict: conflict };
    }

    /**
     * One rule, no tuned constants: a description you have never banked before
     * gets a look, whatever matched it.  Everything else is trusted.
     */
    function confidenceOf(rule, desc, seenBefore) {
        if (!rule) return 'none';
        if (rule.type === 'exact') return 'confirmed';
        return seenBefore ? 'high' : 'review';
    }

    /**
     * Categorise every transaction in chronological order.
     *
     * `reviewFrom` is the index where newly imported rows begin: anything before
     * it is treated as already reviewed, however it was matched.  Without this,
     * opening a workbook cold would flag the first-ever sighting of every
     * merchant you have ever paid - hundreds of rows you settled long ago.
     */
    function classifyAll(transactions, rules, reviewFrom) {
        const seen = new Set();
        const from = reviewFrom === undefined ? transactions.length : reviewFrom;
        return transactions.map(function (tx, i) {
            const found = bestRule(tx.description, tx.date, rules);
            const key = normalise(tx.description);
            const conf = confidenceOf(found.rule, tx.description, seen.has(key) || i < from);
            seen.add(key);
            return {
                category: found.rule ? found.rule.category : 'Uncategorized',
                rule: found.rule || null,
                runnerUp: found.runnerUp || null,
                conflict: found.conflict || null,
                confidence: conf,
                score: found.rule ? found.rule.score : -1
            };
        });
    }

    /**
     * What would adding this rule do?  Compares against the current winners, so
     * it is exact rather than an approximation - and O(transactions), because
     * only the candidate needs scoring.
     */
    function previewRule(candidate, transactions, current, batchFrom) {
        const c = compileRule(candidate);
        const out = { matched: 0, inBatch: 0, fromUncategorised: 0,
                      changes: [], changesBefore: 0 };
        transactions.forEach(function (tx, i) {
            if (!ruleApplies(c, tx.description, tx.date)) return;
            const now = current[i];
            if (c.score < now.score) return;            // an existing rule still wins
            out.matched++;
            const inBatch = batchFrom !== undefined && i >= batchFrom;
            if (inBatch) out.inBatch++;
            if (!now.rule) out.fromUncategorised++;
            else if (now.rule.category !== c.category) {
                // Re-deciding a row in the batch you are reviewing is the job.
                // Re-deciding an older one is the thing to be warned about.
                if (!inBatch) out.changesBefore++;
                out.changes.push({ index: i, tx: tx, from: now.rule.category,
                                   rule: now.rule.match, past: !inBatch });
            }
        });
        return out;
    }

    /** Rule pairs where one match string contains another - the old ordering trap. */
    function collisions(rules) {
        const out = [];
        for (let i = 0; i < rules.length; i++) {
            for (let j = 0; j < rules.length; j++) {
                if (i === j) continue;
                const a = normalise(rules[i].match), b = normalise(rules[j].match);
                if (a && b && a !== b && b.indexOf(a) !== -1 &&
                    rules[i].category !== rules[j].category) {
                    out.push({ shorter: rules[i], longer: rules[j] });
                }
            }
        }
        return out;
    }

    // ------------------------------------------------------------ workbook

    const ROLE_HEADERS = {
        sheet: 'Sheet', label: 'Label', date: 'Date', time: 'Time',
        type: 'Transaction Type', description: 'Description',
        amount: 'Amount', balance: 'Balance'
    };

    function sheetRows(wb, name) {
        const ws = wb.Sheets[name];
        if (!ws) return [];
        return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
    }

    function indexOfHeader(headers, name) {
        if (!name) return -1;
        const want = normalise(name);
        for (let i = 0; i < headers.length; i++) {
            if (normalise(headers[i]) === want) return i;
        }
        return -1;
    }

    function readWorkbook(arrayBuffer) {
        const wb = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
        const model = { wb: wb, accounts: [], rules: [], banks: {}, legacy: [] };

        // _Accounts
        const acc = sheetRows(wb, '_Accounts');
        if (acc.length < 2) throw new Error('No _Accounts sheet, or it has no rows.');
        const accHead = acc[0];
        for (let r = 1; r < acc.length; r++) {
            const row = acc[r];
            if (!row || !row[0]) continue;
            const a = { cols: {} };
            Object.keys(ROLE_HEADERS).forEach(function (role) {
                const i = indexOfHeader(accHead, ROLE_HEADERS[role]);
                const v = i === -1 ? null : row[i];
                if (role === 'sheet' || role === 'label') a[role] = v ? String(v) : '';
                else a.cols[role] = v ? String(v) : null;
            });
            if (!a.label) a.label = a.sheet;
            model.accounts.push(a);
        }

        // Rules
        const rl = sheetRows(wb, 'Rules');
        if (rl.length) {
            const h = rl[0];
            const col = function (n) { return indexOfHeader(h, n); };
            const ci = { match: col('Match'), type: col('Type'), category: col('Category'),
                         from: col('From'), to: col('To'),
                         comment: col('Comment'), where: col('Where') };
            for (let r = 1; r < rl.length; r++) {
                const row = rl[r];
                if (!row || row[ci.match] === null || row[ci.match] === undefined ||
                    String(row[ci.match]).trim() === '') continue;
                model.rules.push({
                    match: String(row[ci.match]),
                    type: (row[ci.type] ? String(row[ci.type]) : 'contains').toLowerCase().trim(),
                    category: row[ci.category] ? String(row[ci.category]) : 'Uncategorized',
                    from: ci.from === -1 ? null : row[ci.from],
                    to: ci.to === -1 ? null : row[ci.to],
                    comment: ci.comment === -1 || !row[ci.comment] ? '' : String(row[ci.comment]),
                    where: ci.where === -1 || !row[ci.where] ? '' : String(row[ci.where])
                });
            }
        }

        // Bank sheets, verbatim
        model.accounts.forEach(function (a) {
            const rows = sheetRows(wb, a.sheet);
            if (!rows.length) throw new Error('Sheet "' + a.sheet + '" named in _Accounts is empty or missing.');
            model.banks[a.sheet] = { headers: rows[0].map(function (h) { return h === null ? '' : String(h); }),
                                     rows: rows.slice(1) };
        });

        // _Legacy (optional, migration aid): Description + Category from the old sheet
        const leg = sheetRows(wb, '_Legacy');
        if (leg.length > 1) {
            const dI = indexOfHeader(leg[0], 'Transaction Description') !== -1
                ? indexOfHeader(leg[0], 'Transaction Description') : indexOfHeader(leg[0], 'Description');
            const cI = indexOfHeader(leg[0], 'Category');
            for (let r = 1; r < leg.length; r++) {
                if (leg[r] && leg[r][dI]) {
                    model.legacy.push({ description: String(leg[r][dI]),
                                        category: cI === -1 ? '' : String(leg[r][cI] || '') });
                }
            }
        }
        return model;
    }

    /** Flatten every bank sheet into one chronological list of transactions. */
    function buildTransactions(model) {
        const txs = [];
        model.accounts.forEach(function (a) {
            const bank = model.banks[a.sheet];
            if (!bank) return;
            const idx = {};
            Object.keys(a.cols).forEach(function (role) {
                idx[role] = indexOfHeader(bank.headers, a.cols[role]);
            });
            bank.rows.forEach(function (row, i) {
                if (!row || row.every(function (c) { return c === null || c === ''; })) return;
                txs.push({
                    account: a.label,
                    sheet: a.sheet,
                    row: i,
                    date: idx.date === -1 ? null : parseDate(row[idx.date]),
                    time: idx.time === -1 ? '' : fmtTime(row[idx.time]),
                    type: idx.type === -1 ? '' : (row[idx.type] === null ? '' : String(row[idx.type])),
                    description: idx.description === -1 ? '' : String(row[idx.description] || ''),
                    amount: idx.amount === -1 ? null : num(row[idx.amount]),
                    balance: idx.balance === -1 ? null : num(row[idx.balance]),
                    raw: row
                });
            });
        });
        txs.sort(function (x, y) {
            const dk = dayKey(x.date) - dayKey(y.date);
            if (dk) return dk;
            return String(x.time).localeCompare(String(y.time));
        });
        return txs;
    }

    /**
     * Does each row's balance equal the one before it plus the amount?
     *
     * The bank's own running balance is the only independent check available on
     * an append-only import: a missing or duplicated row breaks the chain, and
     * nothing else would notice.  Accounts with no balance column are skipped.
     */
    function balanceBreaks(model) {
        const out = [];
        model.accounts.forEach(function (a) {
            const bank = model.banks[a.sheet];
            if (!bank || !a.cols.balance) return;
            const bi = indexOfHeader(bank.headers, a.cols.balance);
            const ai = indexOfHeader(bank.headers, a.cols.amount);
            const di = indexOfHeader(bank.headers, a.cols.date);
            const si = indexOfHeader(bank.headers, a.cols.description);
            if (bi === -1 || ai === -1) return;

            let prev = null;
            bank.rows.forEach(function (row, i) {
                const bal = num(row[bi]), amt = num(row[ai]);
                if (bal === null || amt === null) { prev = bal; return; }
                if (prev !== null) {
                    const expected = Math.round((prev + amt) * 100) / 100;
                    const diff = Math.round((bal - expected) * 100) / 100;
                    if (Math.abs(diff) >= 0.01) {
                        out.push({
                            account: a.label, sheet: a.sheet, row: i,
                            date: di === -1 ? null : parseDate(row[di]),
                            description: si === -1 ? '' : String(row[si] || ''),
                            expected: expected, actual: bal, diff: diff
                        });
                    }
                }
                prev = bal;
            });
        });
        return out;
    }

    // ------------------------------------------------------------ statement

    /** Minimal CSV reader: quotes, escaped quotes, CRLF. Everything stays text. */
    function parseCSV(text) {
        const rows = [];
        let row = [], field = '', quoted = false;
        for (let i = 0; i < text.length; i++) {
            const ch = text[i];
            if (quoted) {
                if (ch === '"') {
                    if (text[i + 1] === '"') { field += '"'; i++; }
                    else quoted = false;
                } else field += ch;
            } else if (ch === '"') quoted = true;
            else if (ch === ',') { row.push(field); field = ''; }
            else if (ch === '\n') { row.push(field); rows.push(row); row = []; field = ''; }
            else if (ch !== '\r') field += ch;
        }
        if (field !== '' || row.length) { row.push(field); rows.push(row); }
        return rows.filter(function (r) { return r.some(function (c) { return String(c).trim() !== ''; }); });
    }

    function readStatement(text, filename) {
        const rows = parseCSV(text);
        let headerAt = -1;
        for (let i = 0; i < Math.min(rows.length, 8); i++) {
            const filled = rows[i].filter(function (c) { return String(c).trim() !== ''; });
            if (filled.length >= 3) { headerAt = i; break; }
        }
        if (headerAt === -1) throw new Error('Could not find a header row in ' + filename);
        return {
            filename: filename,
            banner: headerAt > 0 ? String(rows[0][0] || '') : '',
            headers: rows[headerAt].map(function (h) { return String(h).trim(); }),
            rows: rows.slice(headerAt + 1)
        };
    }

    /**
     * Route a statement to an account by HEADER SIGNATURE - never by filename or
     * dates, which change on every download.  Returns candidates best-first.
     */
    function detectAccount(statement, model) {
        const want = statement.headers.map(normalise).filter(Boolean).join('|');
        const scored = model.accounts.map(function (a) {
            const bank = model.banks[a.sheet];
            const have = bank.headers.map(normalise).filter(Boolean).join('|');
            let score = 0;
            if (have === want) score = 100;
            else {
                const set = new Set(statement.headers.map(normalise));
                const hit = bank.headers.filter(function (h) { return set.has(normalise(h)); }).length;
                score = Math.round((hit / Math.max(bank.headers.length, 1)) * 90);
            }
            return { account: a, score: score };
        });
        scored.sort(function (x, y) { return y.score - x.score; });
        return scored;
    }

    /** Identity of a row for splice purposes: date, time, description, amount. */
    function rowKey(values) {
        return [values.date ? dayKey(values.date) : '', fmtTime(values.time),
                normalise(values.description),
                values.amount === null ? '' : values.amount.toFixed(2)].join('~');
    }

    function statementRowValues(statement, account, row) {
        const idx = {};
        Object.keys(account.cols).forEach(function (role) {
            idx[role] = indexOfHeader(statement.headers, account.cols[role]);
        });
        return {
            date: idx.date === -1 ? null : parseDate(row[idx.date]),
            time: idx.time === -1 ? '' : fmtTime(row[idx.time]),
            description: idx.description === -1 ? '' : (row[idx.description] || ''),
            amount: idx.amount === -1 ? null : num(row[idx.amount])
        };
    }

    /**
     * Where does the statement carry on from what is stored?  Finds the LAST
     * stored transaction inside the file and returns the index just past it.
     * null = the file does not contain it: a gap, or the wrong file.
     */
    function splicePoint(storedKeys, incomingKeys) {
        if (!storedKeys.length) return 0;
        const last = storedKeys[storedKeys.length - 1];
        for (let i = incomingKeys.length - 1; i >= 0; i--) {
            if (incomingKeys[i] === last) return i + 1;
        }
        return null;
    }

    function planImport(statement, account, model) {
        const bank = model.banks[account.sheet];
        const bankIdx = {};
        Object.keys(account.cols).forEach(function (role) {
            bankIdx[role] = indexOfHeader(bank.headers, account.cols[role]);
        });
        const storedKeys = bank.rows.map(function (row) {
            return rowKey({
                date: bankIdx.date === -1 ? null : parseDate(row[bankIdx.date]),
                time: bankIdx.time === -1 ? '' : fmtTime(row[bankIdx.time]),
                description: bankIdx.description === -1 ? '' : (row[bankIdx.description] || ''),
                amount: bankIdx.amount === -1 ? null : num(row[bankIdx.amount])
            });
        });
        const incomingKeys = statement.rows.map(function (row) {
            return rowKey(statementRowValues(statement, account, row));
        });
        const at = splicePoint(storedKeys, incomingKeys);

        // When the splice fails, say WHY: show the last stored row next to the
        // closest rows in the file, field by field, so the mismatch is visible.
        let diagnosis = null;
        if (at === null && bank.rows.length) {
            const lastRow = bank.rows[bank.rows.length - 1];
            const lastValues = {
                date: bankIdx.date === -1 ? null : parseDate(lastRow[bankIdx.date]),
                time: bankIdx.time === -1 ? '' : fmtTime(lastRow[bankIdx.time]),
                description: bankIdx.description === -1 ? '' : (lastRow[bankIdx.description] || ''),
                amount: bankIdx.amount === -1 ? null : num(lastRow[bankIdx.amount])
            };
            const scored = statement.rows.map(function (row, i) {
                const v = statementRowValues(statement, account, row);
                let hits = 0;
                if (v.date && lastValues.date && dayKey(v.date) === dayKey(lastValues.date)) hits++;
                if (v.time === lastValues.time) hits++;
                if (normalise(v.description) === normalise(lastValues.description)) hits++;
                if (v.amount !== null && lastValues.amount !== null &&
                    v.amount.toFixed(2) === lastValues.amount.toFixed(2)) hits++;
                return { index: i, values: v, hits: hits };
            }).filter(function (c) { return c.hits > 0; });
            scored.sort(function (a, b) { return b.hits - a.hits; });
            diagnosis = { last: lastValues, candidates: scored.slice(0, 3) };
        }

        // Re-order each incoming row into the bank sheet's own column order, and
        // make numeric columns actually numeric so Excel keeps treating them so.
        const mapped = at === null ? [] : statement.rows.slice(at).map(function (row) {
            return bank.headers.map(function (h) {
                const i = indexOfHeader(statement.headers, h);
                if (i === -1) return null;
                const v = row[i];
                if (v === '' || v === null || v === undefined) return null;
                const n = num(v);
                return (n !== null && /^[-+]?[\d.,\s£$€()]+$/.test(String(v))) ? n : String(v);
            });
        });
        let from = null, to = null;
        statement.rows.forEach(function (row) {
            const d = statementRowValues(statement, account, row).date;
            if (!d) return;
            if (!from || dayKey(d) < dayKey(from)) from = d;
            if (!to || dayKey(d) > dayKey(to)) to = d;
        });

        return { at: at, rows: mapped, skipped: at === null ? 0 : at,
                 total: statement.rows.length, diagnosis: diagnosis,
                 periodFrom: from ? fmtDate(from) : '', periodTo: to ? fmtDate(to) : '' };
    }

    // ------------------------------------------------------------ writing

    const MERGED_HEADERS = ['Account', 'Date', 'Time', 'Transaction Type',
        'Transaction Description', 'Amount', 'Currency', 'Balance',
        'Category', 'Matched Rule', 'Confidence'];

    const CONFIDENCE_LABEL = { confirmed: 'Confirmed', high: 'High', review: 'Needs review', none: 'None' };

    function writeWorkbook(model, transactions, results, importLog) {
        const wb = model.wb;

        model.accounts.forEach(function (a) {
            const bank = model.banks[a.sheet];
            const aoa = [bank.headers].concat(bank.rows);
            wb.Sheets[a.sheet] = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true });
        });

        const ruleRows = model.rules.map(function (r) {
            return [r.match, r.type, r.category,
                    r.from ? (r.from instanceof Date ? fmtDate(r.from) : r.from) : '',
                    r.to ? (r.to instanceof Date ? fmtDate(r.to) : r.to) : '',
                    r.comment || '', r.where || '', ''];
        });
        const rulesWs = XLSX.utils.aoa_to_sheet(
            [['Match', 'Type', 'Category', 'From', 'To', 'Comment', 'Where', 'Claimed']]
                .concat(ruleRows));
        // Claimed is a live formula counting what each rule actually won - unlike
        // the old SUMPRODUCT/SEARCH, which counted rows other rules had claimed.
        // The cached value has to be written too: SheetJS drops formula cells
        // that carry no value, and it keeps the count readable without Excel.
        const claimed = new Map();
        results.forEach(function (r) {
            if (r.rule) claimed.set(r.rule.match, (claimed.get(r.rule.match) || 0) + 1);
        });
        rulesWs['!ref'] = XLSX.utils.encode_range({
            s: { r: 0, c: 0 }, e: { r: ruleRows.length, c: 7 }
        });
        model.rules.forEach(function (r, i) {
            const row = i + 2;
            // COUNTIF treats * and ? in its criteria as wildcards, so a rule like
            // ZETTLE_* or *BAKERY* would over-count.  Escape them Excel-style (~).
            const criteria = 'SUBSTITUTE(SUBSTITUTE(SUBSTITUTE($A' + row +
                             ',"~","~~"),"*","~*"),"?","~?")';
            rulesWs['H' + row] = { t: 'n', v: claimed.get(r.match) || 0,
                                   f: 'COUNTIF(Merged!$J:$J,' + criteria + ')' };
        });
        wb.Sheets['Rules'] = rulesWs;

        const mergedRows = transactions.map(function (tx, i) {
            const res = results[i];
            return [tx.account, fmtDate(tx.date), tx.time, tx.type, tx.description,
                    tx.amount, 'GBP', tx.balance === null ? '' : tx.balance,
                    res.category, res.rule ? res.rule.match : '',
                    CONFIDENCE_LABEL[res.confidence] || res.confidence];
        });
        wb.Sheets['Merged'] = XLSX.utils.aoa_to_sheet([MERGED_HEADERS].concat(mergedRows));

        const impHeaders = ['Imported At', 'Account', 'File', 'Period From', 'Period To',
                            'Rows In File', 'Rows Added', 'Rows Skipped'];
        const existing = sheetRows(wb, '_Imports');
        const impRows = existing.length > 1 ? existing.slice(1) : [];
        if (importLog) impRows.push(importLog);
        wb.Sheets['_Imports'] = XLSX.utils.aoa_to_sheet([impHeaders].concat(impRows));

        ['Merged', '_Imports'].forEach(function (name) {
            if (wb.SheetNames.indexOf(name) === -1) wb.SheetNames.push(name);
        });
        if (wb.SheetNames.indexOf('_Legacy') !== -1) {   // migration aid, not kept
            wb.SheetNames.splice(wb.SheetNames.indexOf('_Legacy'), 1);
            delete wb.Sheets['_Legacy'];
        }
        return XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    }

    return {
        normalise: normalise, parseDate: parseDate, fmtDate: fmtDate,
        fmtTime: fmtTime, num: num,
        compileRules: compileRules, compileRule: compileRule, ruleScore: ruleScore,
        ruleApplies: ruleApplies, bestRule: bestRule, classifyAll: classifyAll,
        previewRule: previewRule, collisions: collisions, literalLength: literalLength,
        readWorkbook: readWorkbook, buildTransactions: buildTransactions,
        balanceBreaks: balanceBreaks,
        readStatement: readStatement, detectAccount: detectAccount,
        planImport: planImport, splicePoint: splicePoint, writeWorkbook: writeWorkbook,
        CONFIDENCE_LABEL: CONFIDENCE_LABEL
    };
})();

if (typeof module !== 'undefined' && module.exports) module.exports = Engine;
