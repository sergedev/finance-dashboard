/**
 * Review component: the transaction / rules / categories table, the rule editor,
 * the health panel and the save bar.
 *
 * Mounted by both pages — the import page (sync.html) and the dashboard
 * (index.html) — so there is exactly one implementation of the review UI.
 *
 *   Review.mount(rootElement)          inject the markup
 *   Review.load(model, filename)       take a parsed workbook and render
 *   Review.model()                     the live model, for the import splice
 *   Review.applyImport({...})          append imported rows and switch to them
 *   Review.undoImport()                take the last import back out
 *   Review.transactions()              { txs, results } for anything downstream
 *   Review.onChange(fn)                called whenever the data changes
 */
const Review = (function () {
    'use strict';

    const state = {
        model: null,          // parsed workbook
        rules: [],            // compiled rules
        txs: [],              // every transaction, chronological
        results: [],          // classification, parallel to txs
        newKeys: null,        // Set of "sheet:row" appended this session
        batchFrom: undefined, // index in txs of the earliest appended row
        applied: null,        // { sheet, count } of the last import, so it can be undone
        importLog: null,
        filter: 'needs',
        dismissed: null,      // descriptions nodded at this session (not persisted)
        undo: [],             // stack of snapshots, newest last
        expanded: null,       // rule indexes whose transactions are shown
        expandedCat: null,    // category names whose transactions are shown
        search: '',
        breakdownUnit: 'month',   // 'month' | 'year'
        sort: { tx: null, rules: null, categories: null, breakdown: null },
        dirty: false,
        editing: null,        // { txIndex, ruleIndex, existing, mode }
        filename: 'finance.xlsx'
    };

    let root = null;
    let changeHandlers = [];

    const $ = function (id) { return root.querySelector('#' + id); };

    function money(n) {
        if (n === null || n === undefined || isNaN(n)) return '';
        return (n < 0 ? '-£' : '£') + Math.abs(n).toFixed(2);
    }

    function esc(s) {
        return String(s === null || s === undefined ? '' : s)
            .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    /**
     * A rule's comment and its "where", in one string, comment first, joined
     * with a dot - the same shape used in the category detail rows.
     */
    function noteHtml(rule) {
        if (!rule) return '';
        const parts = [];
        if (rule.comment) parts.push(esc(rule.comment));
        if (rule.where) parts.push('<span class="where">' + esc(rule.where) + '</span>');
        return parts.join(' · ');
    }

    function isNew(tx) {
        return state.newKeys ? state.newKeys.has(tx.sheet + ':' + tx.row) : false;
    }

    function fireChange() {
        changeHandlers.forEach(function (fn) { fn(state.txs, state.results); });
    }

    /**
     * Undo is snapshot-based: rules and dismissals are small, so copying them
     * before every change is cheaper and far more reliable than trying to
     * reverse each operation.
     */
    function snapshot(label) {
        state.undo.push({
            label: label,
            rules: JSON.parse(JSON.stringify(state.model.rules)),
            dismissed: Array.from(state.dismissed || [])
        });
        if (state.undo.length > 30) state.undo.shift();
        refreshUndo();
    }

    function undo() {
        const snap = state.undo.pop();
        if (!snap) return;
        state.model.rules = snap.rules;
        state.dismissed = new Set(snap.dismissed);
        recompute();
        refreshUndo();
        markDirty('Undid: ' + snap.label);
        render();
    }

    function refreshUndo() {
        const btn = $('undoBtn');
        btn.hidden = state.undo.length === 0;
        if (state.undo.length) {
            btn.textContent = '↶ Undo ' + state.undo[state.undo.length - 1].label;
        }
    }

    function markDirty(note) {
        state.dirty = true;
        $('dirtyDot').hidden = false;
        $('saveBtn').disabled = false;
        $('saveNote').textContent = note || 'Unsaved changes';
    }

    function money(n) {
        if (n === null || n === undefined || isNaN(n)) return '';
        return (n < 0 ? '-£' : '£') + Math.abs(n).toFixed(2);
    }

    function esc(s) {
        return String(s === null || s === undefined ? '' : s)
            .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    /** Recompile rules, rebuild the transaction list, reclassify everything. */
    function recompute() {
        state.rules = Engine.compileRules(state.model.rules);
        state.txs = Engine.buildTransactions(state.model);
        if (state.newKeys && state.newKeys.size) {
            state.batchFrom = state.txs.length;
            state.txs.forEach(function (tx, i) {
                if (state.newKeys.has(tx.sheet + ':' + tx.row) && i < state.batchFrom) {
                    state.batchFrom = i;
                }
            });
        }
        // Rows that were in the workbook when it was opened count as reviewed
        state.results = Engine.classifyAll(state.txs, state.rules, state.batchFrom);
        fireChange();
    }


    // --------------------------------------------------------------- render

    /** Does this row still want a human? Dismissing is session-only. */
    function isFlagged(i) {
        const r = state.results[i];
        if (r.conflict) return true;
        if (!r.rule) return true;                       // nothing matched at all
        if (r.confidence !== 'review') return false;
        return !state.dismissed.has(Engine.normalise(state.txs[i].description));
    }

    /** Why is this row asking for attention? Shown on the row itself. */
    function flagReason(i) {
        const r = state.results[i];
        if (r.conflict) return 'two rules score the same and disagree';
        if (!r.rule) return 'no rule matches this';
        if (r.confidence === 'review') return 'first time you\'ve seen this description';
        return '';
    }

    function counts() {
        let needs = 0, batch = 0;
        state.txs.forEach(function (tx, i) {
            if (isFlagged(i)) needs++;
            if (isNew(tx)) batch++;
        });
        return { needs: needs, batch: batch, all: state.txs.length };
    }

    function matchesSearch(text) {
        if (!state.search) return true;
        return String(text).toUpperCase().indexOf(state.search.toUpperCase()) !== -1;
    }

    function visibleIndexes() {
        const out = [];
        state.txs.forEach(function (tx, i) {
            const r = state.results[i];
            if (state.search && !matchesSearch(tx.description) && !matchesSearch(r.category) &&
                !matchesSearch(r.rule ? r.rule.match : '')) return;
            if (state.filter === 'needs') {
                if (isFlagged(i)) out.push(i);
            } else if (state.filter === 'batch') {
                if (isNew(tx)) out.push(i);
            } else out.push(i);
        });
        if (state.filter === 'all') out.reverse();
        return out;
    }

    function render() {
        const c = counts();
        $('tabs').innerHTML = [
            tab('needs', 'Needs you', c.needs),
            c.batch ? tab('batch', 'This import', c.batch) : '',
            tab('all', 'Everything', c.all),
            tab('rules', 'Rules', state.model.rules.length),
            tab('categories', 'Categories', categorySummary().length),
            tab('breakdown', 'Breakdown', '')
        ].join('');
        Array.prototype.forEach.call(root.querySelectorAll('.tab'), function (el) {
            el.addEventListener('click', function () {
                state.filter = el.dataset.filter;
                render();
            });
        });

        const uncategorised = state.results.filter(function (r) { return !r.rule; }).length;
        const acceptable = state.txs.filter(function (_tx, i) {
            return isFlagged(i) && state.results[i].rule && !state.results[i].conflict;
        }).length;

        if (state.filter === 'breakdown') {
            $('reviewStats').innerHTML =
                '<div class="tabs small">' +
                    '<button class="tab' + (state.breakdownUnit === 'month' ? ' on' : '') +
                        '" data-bunit="month">Months</button>' +
                    '<button class="tab' + (state.breakdownUnit === 'year' ? ' on' : '') +
                        '" data-bunit="year">Years</button>' +
                '</div>';
            Array.prototype.forEach.call(root.querySelectorAll('[data-bunit]'), function (b) {
                b.addEventListener('click', function () {
                    state.breakdownUnit = b.dataset.bunit;
                    render();
                });
            });
            renderBreakdown();
        } else if (state.filter === 'categories') {
            const cats = categorySummary();
            $('reviewStats').innerHTML = '<span>' + cats.length + ' categories</span>';
            renderCategories(cats);
        } else if (state.filter === 'rules') {
            $('reviewStats').innerHTML = '<button class="mini" id="newRule">+ New rule</button>' +
                '<span>' + state.model.rules.length + ' rules</span>';
            $('newRule').addEventListener('click', function () { openEditor(null); });
            renderRules();
        } else {
            $('reviewStats').innerHTML =
                (acceptable ? '<button class="mini ok-btn" id="dismissAll">Looks right on all ' +
                    acceptable + '</button>' : '') +
                '<span>' + c.all + ' transactions</span>' +
                '<span>' + uncategorised + ' uncategorised</span>';
            if (acceptable) $('dismissAll').addEventListener('click', dismissAll);
            renderRows(visibleIndexes());
        }
        renderHealth();
        renderCategoryList();
    }

    function tab(key, label, n) {
        return '<button class="tab' + (state.filter === key ? ' on' : '') +
            '" data-filter="' + key + '">' + label + ' <span class="pill">' + n + '</span></button>';
    }


    // ------------------------------------------------------------- sorting

    /**
     * One sorting mechanism for every table.  Each view names its columns and
     * supplies a value accessor; clicking a header toggles direction, clicking a
     * third time clears it and restores that view's natural order.
     */
    function sortHeader(view, key, label, cls) {
        const s = state.sort[view];
        const on = s && s.key === key;
        return '<th class="sortable ' + (cls || '') + (on ? ' sorted' : '') +
            '" data-sort="' + view + ':' + key + '">' + label +
            '<span class="arrow">' + (on ? (s.dir === 'asc' ? '▲' : '▼') : '') + '</span></th>';
    }

    function wireSortHeaders() {
        Array.prototype.forEach.call($('txHead').querySelectorAll('[data-sort]'), function (th) {
            th.addEventListener('click', function () {
                // split on the FIRST colon only - period keys contain one too
                const spec = th.dataset.sort;
                const at = spec.indexOf(':');
                const view = spec.slice(0, at), key = spec.slice(at + 1);
                const s = state.sort[view];
                if (!s || s.key !== key) state.sort[view] = { key: key, dir: 'asc' };
                else if (s.dir === 'asc') state.sort[view] = { key: key, dir: 'desc' };
                else state.sort[view] = null;           // back to natural order
                render();
            });
        });
    }

    /** Stable sort by an accessor, nulls last, strings case-insensitive. */
    function sortBy(list, view, accessors) {
        const s = state.sort[view];
        if (!s || !accessors[s.key]) return list;
        const get = accessors[s.key];
        const dir = s.dir === 'asc' ? 1 : -1;
        return list.map(function (item, i) { return { item: item, i: i }; })
            .sort(function (a, b) {
                let x = get(a.item), y = get(b.item);
                if (typeof x === 'string' || typeof y === 'string') {
                    x = String(x === null || x === undefined ? '' : x).toUpperCase();
                    y = String(y === null || y === undefined ? '' : y).toUpperCase();
                    if (x === y) return a.i - b.i;
                    return x < y ? -dir : dir;
                }
                x = x === null || x === undefined || isNaN(x) ? -Infinity : x;
                y = y === null || y === undefined || isNaN(y) ? -Infinity : y;
                if (x === y) return a.i - b.i;
                return x < y ? -dir : dir;
            })
            .map(function (w) { return w.item; });
    }

    function txHead() {
        return '<tr>' +
            sortHeader('tx', 'date', 'Date') +
            sortHeader('tx', 'description', 'Description') +
            sortHeader('tx', 'amount', 'Amount', 'r') +
            sortHeader('tx', 'category', 'Category') +
            sortHeader('tx', 'rule', 'Matched by') +
            '<th></th></tr>';
    }

    function ruleHead() {
        return '<tr>' +
            sortHeader('rules', 'match', 'Match') +
            sortHeader('rules', 'type', 'Type') +
            sortHeader('rules', 'category', 'Category') +
            sortHeader('rules', 'scope', 'Scope') +
            sortHeader('rules', 'note', 'Comment · where') +
            sortHeader('rules', 'claimed', 'Claimed', 'r') +
            '<th></th></tr>';
    }

    /** The rules sheet, editable: this is the safety net for a bad edit. */
    function renderRules() {
        $('txHead').innerHTML = ruleHead();
        wireSortHeaders();
        $('emptyNote').hidden = true;

        const claimed = claimedCounts();
        const won = x => claimed.get(x.rule.match + '\u0000' + x.rule.type) || 0;

        let rows = state.model.rules
            .map(function (rule, i) { return { rule: rule, i: i }; })
            .filter(function (x) {
                return !state.search || matchesSearch(x.rule.match) ||
                    matchesSearch(x.rule.category) || matchesSearch(x.rule.comment || '') ||
                    matchesSearch(x.rule.where || '');
            });

        rows = sortBy(rows, 'rules', {
            match: x => x.rule.match,
            type: x => x.rule.type,
            category: x => x.rule.category,
            scope: x => (x.rule.from ? String(x.rule.from) : ''),
            note: x => [x.rule.comment, x.rule.where].filter(Boolean).join(' '),
            claimed: won
        });

        if (!rows.length) {
            $('txBody').innerHTML = '';
            $('emptyNote').hidden = false;
            $('emptyNote').textContent = 'No rules match "' + state.search + '".';
            return;
        }

        $('txBody').innerHTML = rows.map(function (x) {
            const r = x.rule;
            const n = claimed.get(r.match + '\u0000' + r.type) || 0;
            const open = state.expanded.has(x.i);
            const scope = [r.from ? 'from ' + Engine.fmtDate(Engine.parseDate(r.from)) : '',
                           r.to ? 'to ' + Engine.fmtDate(Engine.parseDate(r.to)) : '']
                          .filter(Boolean).join(' ');
            return '<tr class="rule-row ' + (n === 0 ? 'dead' : '') + (open ? ' open' : '') +
                    '" data-rule-row="' + x.i + '">' +
                '<td class="desc"><span class="chev">' + (open ? '▾' : '▸') + '</span>' +
                    '<code>' + esc(r.match) + '</code></td>' +
                '<td class="nowrap muted">' + esc(r.type) + '</td>' +
                '<td>' + esc(r.category) + '</td>' +
                '<td class="nowrap muted">' + esc(scope || '—') + '</td>' +
                '<td class="muted">' + (noteHtml(r) || '—') + '</td>' +
                '<td class="r nowrap' + (n === 0 ? ' zero' : '') + '">' + n + '</td>' +
                '<td class="r nowrap">' +
                    '<button class="mini" data-rule-edit="' + x.i + '">Edit</button>' +
                    '<button class="mini danger-btn" data-rule-del="' + x.i + '">Delete</button>' +
                '</td></tr>' +
                (open ? ruleDetail(r, n) : '');
        }).join('');

        Array.prototype.forEach.call($('txBody').querySelectorAll('[data-rule-row]'), function (row) {
            row.addEventListener('click', function (e) {
                if (e.target.closest('button')) return;
                toggleRule(+row.dataset.ruleRow);
            });
        });
        Array.prototype.forEach.call($('txBody').querySelectorAll('[data-rule-edit]'), function (b) {
            b.addEventListener('click', function () { openRuleEditor(+b.dataset.ruleEdit); });
        });
        Array.prototype.forEach.call($('txBody').querySelectorAll('[data-rule-del]'), function (b) {
            b.addEventListener('click', function () { deleteRule(+b.dataset.ruleDel); });
        });
    }

    /** The transactions a rule actually won, newest first. */
    function ruleDetail(rule, n) {
        const hits = [];
        state.results.forEach(function (res, i) {
            if (res.rule && res.rule.match === rule.match && res.rule.type === rule.type) {
                hits.push(i);
            }
        });
        hits.reverse();

        if (!hits.length) {
            return '<tr class="detail"><td colspan="7"><div class="detail-box muted">' +
                'This rule wins nothing. Either its keyword never appears, or a more specific ' +
                'rule claims every transaction it would match.</div></td></tr>';
        }

        const MAX = 200;
        const body = hits.slice(0, MAX).map(function (i) {
            const tx = state.txs[i];
            return '<tr><td class="nowrap">' + esc(Engine.fmtDate(tx.date)) + '</td>' +
                '<td>' + esc(tx.description) + '</td>' +
                '<td class="nowrap muted">' + esc(tx.account) + '</td>' +
                '<td class="r nowrap ' + (tx.amount < 0 ? 'out' : 'in') + '">' +
                    money(tx.amount) + '</td></tr>';
        }).join('');

        const total = hits.reduce(function (sum, i) {
            return sum + (state.txs[i].amount || 0);
        }, 0);

        const note = noteHtml(rule);
        return '<tr class="detail"><td colspan="7"><div class="detail-box">' +
            '<div class="detail-head">' + n + ' transaction' + (n === 1 ? '' : 's') +
            ' · ' + money(total) + ' total' +
            (hits.length > MAX ? ' · showing the newest ' + MAX : '') +
            (note ? ' · ' + note : '') + '</div>' +
            '<table class="detail-table">' + body + '</table>' +
            '</div></td></tr>';
    }

    function toggleRule(index) {
        if (state.expanded.has(index)) state.expanded.delete(index);
        else state.expanded.add(index);
        render();
    }

    function deleteRule(index) {
        const rule = state.model.rules[index];
        snapshot('delete "' + rule.match + '"');
        state.model.rules.splice(index, 1);
        recompute();
        markDirty('Deleted rule "' + rule.match + '"');
        render();
    }

    function renderRows(indexes) {
        $('txHead').innerHTML = txHead();
        wireSortHeaders();
        indexes = sortBy(indexes, 'tx', {
            date: i => (state.txs[i].date ? state.txs[i].date.getTime() : 0),
            description: i => state.txs[i].description,
            amount: i => state.txs[i].amount,
            category: i => state.results[i].category,
            rule: i => (state.results[i].rule ? state.results[i].rule.match : '')
        });
        const claimed = claimedCounts();
        const body = $('txBody');
        const note = $('emptyNote');
        if (!indexes.length) {
            body.innerHTML = '';
            note.hidden = false;
            note.textContent = state.filter === 'needs'
                ? 'Nothing needs your attention.'
                : (state.filter === 'batch'
                    ? 'No statement imported yet.'
                    : 'Nothing to show.');
            return;
        }
        note.hidden = true;

        const MAX = 600;
        const shown = indexes.slice(0, MAX);
        body.innerHTML = shown.map(function (i) {
            const tx = state.txs[i], r = state.results[i];
            const flagged = isFlagged(i);
            const dismissed = r.confidence === 'review' && !flagged;
            const sameDesc = state.dismissed && countSameInBatch(i);

            let badge;
            if (r.conflict) badge = '<span class="badge conflict">conflict</span>';
            else if (!r.rule) badge = '<span class="badge none">no rule</span>';
            else if (dismissed) badge = '<span class="badge ok">ok</span>';
            else if (r.confidence === 'review') badge = '<span class="badge review">new to you</span>';
            else if (r.confidence === 'confirmed') badge = '<span class="badge confirmed">exact</span>';
            else badge = '';

            const reason = flagged
                ? '<div class="why">' + esc(flagReason(i)) + '</div>' : '';

            // the rule, with its comment and where, without opening anything
            let matched = '<span class="muted">—</span>';
            if (r.rule) {
                const extra = noteHtml(r.rule);
                const beat = r.runnerUp
                    ? '<div class="rule-note">beat <code>' + esc(r.runnerUp.match) + '</code> → ' +
                      esc(r.runnerUp.category) + '</div>'
                    : '';
                const won = claimed.get(r.rule.match + '\u0000' + r.rule.type) || 0;
                matched = '<code>' + esc(r.rule.match) + '</code> <span class="muted">' +
                    esc(r.rule.type) + ' (' + won + ')</span>' +
                    (r.rule.from ? '<span class="muted"> · from ' + esc(Engine.fmtDate(r.rule.from)) + '</span>' : '') +
                    (extra ? '<div class="rule-note">' + extra + '</div>' : '') + beat;
            }

            let actions = '';
            if (flagged && r.rule && !r.conflict) {
                actions += '<button class="mini ok-btn" data-ok="' + i + '">Looks right' +
                    (sameDesc > 1 ? ' (' + sameDesc + ')' : '') + '</button>';
            }
            actions += '<button class="mini" data-edit="' + i + '">' +
                (r.rule ? 'Change rule' : 'Categorise') + '</button>';

            return '<tr class="' + (isNew(tx) ? 'is-new ' : '') +
                    (flagged ? (r.rule ? 'review' : 'none') : '') + '">' +
                '<td class="nowrap">' + esc(Engine.fmtDate(tx.date)) +
                    (tx.time ? '<span class="acct">' + esc(tx.time) + '</span>' : '') + '</td>' +
                '<td class="desc" title="' + esc(tx.account) + '">' + esc(tx.description) +
                    reason + '</td>' +
                '<td class="r nowrap ' + (tx.amount < 0 ? 'out' : 'in') + '">' + money(tx.amount) + '</td>' +
                '<td>' + esc(r.category) + ' ' + badge + '</td>' +
                '<td>' + matched + '</td>' +
                '<td class="r nowrap">' + actions + '</td></tr>';
        }).join('');

        if (indexes.length > MAX) {
            body.innerHTML += '<tr><td colspan="6" class="muted pad">Showing the first ' + MAX +
                ' of ' + indexes.length + '.</td></tr>';
        }

        Array.prototype.forEach.call(body.querySelectorAll('[data-edit]'), function (btn) {
            btn.addEventListener('click', function () { openEditor(+btn.dataset.edit); });
        });
        Array.prototype.forEach.call(body.querySelectorAll('[data-ok]'), function (btn) {
            btn.addEventListener('click', function () { dismissRow(+btn.dataset.ok); });
        });
    }

    /**
     * Categories are never configured anywhere - they are whatever the rules
     * currently produce, so this is derived on every render.
     */
    function categorySummary() {
        const by = new Map();
        state.results.forEach(function (r, i) {
            const name = r.category;
            if (!by.has(name)) by.set(name, { name: name, indexes: [], total: 0, rules: new Set() });
            const c = by.get(name);
            c.indexes.push(i);
            c.total += state.txs[i].amount || 0;
            if (r.rule) c.rules.add(r.rule.match);
        });
        return Array.from(by.values()).sort(function (a, b) {
            return b.indexes.length - a.indexes.length;
        });
    }

    function catHead() {
        return '<tr>' +
            sortHeader('categories', 'name', 'Category') +
            sortHeader('categories', 'count', 'Transactions', 'r') +
            sortHeader('categories', 'total', 'Total', 'r') +
            sortHeader('categories', 'rules', 'Rules', 'r') +
            '<th></th></tr>';
    }

    function renderCategories(cats) {
        $('txHead').innerHTML = catHead();
        wireSortHeaders();
        $('emptyNote').hidden = true;

        let rows = cats.filter(function (c) { return !state.search || matchesSearch(c.name); });
        rows = sortBy(rows, 'categories', {
            name: c => c.name,
            count: c => c.indexes.length,
            total: c => c.total,
            rules: c => c.rules.size
        });
        if (!rows.length) {
            $('txBody').innerHTML = '';
            $('emptyNote').hidden = false;
            $('emptyNote').textContent = 'No categories match "' + state.search + '".';
            return;
        }

        $('txBody').innerHTML = rows.map(function (c) {
            const open = state.expandedCat.has(c.name);
            return '<tr class="rule-row ' + (open ? 'open' : '') +
                    (c.name === 'Uncategorized' ? ' dead' : '') +
                    '" data-cat-row="' + esc(c.name) + '">' +
                '<td class="desc"><span class="chev">' + (open ? '▾' : '▸') + '</span>' +
                    esc(c.name) + '</td>' +
                '<td class="r nowrap">' + c.indexes.length + '</td>' +
                '<td class="r nowrap ' + (c.total < 0 ? 'out' : 'in') + '">' + money(c.total) + '</td>' +
                '<td class="r nowrap muted">' + (c.rules.size || '—') + '</td>' +
                '<td></td></tr>' +
                (open ? categoryDetail(c) : '');
        }).join('');

        Array.prototype.forEach.call($('txBody').querySelectorAll('[data-cat-row]'), function (row) {
            row.addEventListener('click', function () {
                const name = row.dataset.catRow;
                if (state.expandedCat.has(name)) state.expandedCat.delete(name);
                else state.expandedCat.add(name);
                render();
            });
        });
    }

    /** Newest 200 transactions in a category, with the rule that put them there. */
    function categoryDetail(c) {
        const MAX = 200;
        const hits = c.indexes.slice().reverse().slice(0, MAX);
        const body = hits.map(function (i) {
            const tx = state.txs[i], r = state.results[i];
            const note = noteHtml(r.rule);
            return '<tr><td class="nowrap">' + esc(Engine.fmtDate(tx.date)) + '</td>' +
                '<td>' + esc(tx.description) + '</td>' +
                '<td class="muted">' + (r.rule ? '<code>' + esc(r.rule.match) + '</code>' : '—') +
                    (note ? '<div class="rule-note">' + note + '</div>' : '') + '</td>' +
                '<td class="r nowrap ' + (tx.amount < 0 ? 'out' : 'in') + '">' + money(tx.amount) + '</td></tr>';
        }).join('');

        return '<tr class="detail"><td colspan="5"><div class="detail-box">' +
            '<div class="detail-head">' + c.indexes.length + ' transaction' +
            (c.indexes.length === 1 ? '' : 's') + ' · ' + money(c.total) + ' total · ' +
            (c.rules.size ? c.rules.size + ' rule' + (c.rules.size === 1 ? '' : 's') : 'no rules') +
            (c.indexes.length > MAX ? ' · showing the newest ' + MAX : '') + '</div>' +
            '<table class="detail-table">' + body + '</table>' +
            '</div></td></tr>';
    }


    /** Period keys across the whole ledger, oldest first. */
    function breakdownPeriods() {
        const keys = [];
        const seen = new Set();
        state.txs.forEach(function (tx) {
            if (!tx.date) return;
            const key = state.breakdownUnit === 'year'
                ? String(tx.date.getFullYear())
                : tx.date.getFullYear() + '-' + String(tx.date.getMonth() + 1).padStart(2, '0');
            if (!seen.has(key)) { seen.add(key); keys.push(key); }
        });
        return keys.sort();
    }

    function periodOf(tx) {
        return state.breakdownUnit === 'year'
            ? String(tx.date.getFullYear())
            : tx.date.getFullYear() + '-' + String(tx.date.getMonth() + 1).padStart(2, '0');
    }

    function periodLabel(key) {
        if (state.breakdownUnit === 'year') return key;
        const [y, m] = key.split('-');
        return new Date(+y, +m - 1, 1).toLocaleDateString('en-GB', { month: 'short', year: '2-digit' });
    }

    /** A tiny inline trend line - one per row, so nothing overlaps. */
    function sparkline(values) {
        if (values.length < 2) return '';
        const max = Math.max.apply(null, values.map(Math.abs)) || 1;
        const w = 60, h = 16;
        const pts = values.map(function (v, i) {
            const x = (i / (values.length - 1)) * w;
            const y = h - (Math.abs(v) / max) * (h - 2) - 1;
            return x.toFixed(1) + ',' + y.toFixed(1);
        }).join(' ');
        return '<svg class="spark" width="' + w + '" height="' + h + '" viewBox="0 0 ' + w + ' ' + h +
            '" preserveAspectRatio="none"><polyline points="' + pts + '"/></svg>';
    }

    /**
     * Category x period grid.
     *
     * The average divides by the periods a category was actually ACTIVE in, not
     * by every period in the table - so an occasional category reads as what it
     * costs when it happens.
     */
    function renderBreakdown() {
        const periods = breakdownPeriods();
        const rows = new Map();

        state.results.forEach(function (r, i) {
            const tx = state.txs[i];
            if (!tx.date) return;
            if (!rows.has(r.category)) {
                rows.set(r.category, { name: r.category, cells: {}, total: 0, indexes: [] });
            }
            const row = rows.get(r.category);
            const key = periodOf(tx);
            if (!row.cells[key]) row.cells[key] = { net: 0, in: 0, out: 0, n: 0 };
            const cell = row.cells[key];
            cell.net += tx.amount || 0;
            cell.n++;
            if ((tx.amount || 0) > 0) cell.in += tx.amount;
            else cell.out += tx.amount;
            row.total += tx.amount || 0;
            row.indexes.push(i);
        });

        let all = Array.from(rows.values());
        if (state.search) all = all.filter(function (r) { return matchesSearch(r.name); });

        // income and transfers sit below a divider: they would otherwise dominate
        // the shading and read oddly among the spending rows
        const isIncomeLike = function (r) {
            return r.total > 0 || /^transfers?$/i.test(r.name);
        };
        const spending = all.filter(function (r) { return !isIncomeLike(r); })
            .sort(function (a, b) { return a.total - b.total; });
        const income = all.filter(isIncomeLike)
            .sort(function (a, b) { return b.total - a.total; });

        $('txHead').innerHTML = '<tr>' +
            sortHeader('breakdown', 'name', 'Category') +
            periods.map(function (p) {
                return sortHeader('breakdown', 'p:' + p, esc(periodLabel(p)), 'r');
            }).join('') +
            '<th class="r">Trend</th>' +
            sortHeader('breakdown', 'avg', 'Avg', 'r') +
            sortHeader('breakdown', 'total', 'Total', 'r') + '</tr>';
        wireSortHeaders();

        // sorting applies within each group, so income stays below the divider
        const accessors = {
            name: r => r.name,
            total: r => r.total,
            avg: r => {
                const active = periods.filter(p => r.cells[p] && r.cells[p].n > 0).length;
                return active ? r.total / active : 0;
            }
        };
        periods.forEach(function (p) {
            accessors['p:' + p] = r => (r.cells[p] ? r.cells[p].net : 0);
        });
        const ordered = list => sortBy(list, 'breakdown', accessors);

        const renderRow = function (r) {
            const open = state.expandedCat.has(r.name);
            const cellsRaw = periods.map(function (p) {
                return r.cells[p] || { net: 0, in: 0, out: 0, n: 0 };
            });
            const values = cellsRaw.map(function (c) { return c.net; });
            // "active" counts periods with transactions, not periods with a
            // non-zero net - a month of £500 in and £500 out was not idle
            const active = cellsRaw.filter(function (c) { return c.n > 0; }).length;
            const avg = active ? r.total / active : 0;
            const max = Math.max.apply(null, values.map(Math.abs)) || 1;

            // A row that only ever moves one way needs no signs - its direction is
            // in the Total column.  A row that moves BOTH ways signs every cell,
            // because that is exactly where +42 hiding among spending matters.
            const mixed = values.some(function (v) { return v > 0; }) &&
                          values.some(function (v) { return v < 0; });

            const cells = cellsRaw.map(function (c) {
                const v = c.net;
                if (!c.n) return '<td class="r muted zero">·</td>';
                if (!v) {
                    // money moved both ways and cancelled: say so, don't imply silence
                    return '<td class="r muted zero" title="' +
                        money(c.in) + ' in, ' + money(c.out) + ' out">0</td>';
                }
                const weight = Math.min(1, Math.abs(v) / max);
                const positive = v > 0;
                const tint = positive
                    ? 'rgba(16,185,129,' + (0.06 + weight * 0.20).toFixed(3) + ')'
                    : 'rgba(37,99,235,' + (0.04 + weight * 0.16).toFixed(3) + ')';
                const sign = mixed ? (positive ? '+' : '\u2212') : '';
                return '<td class="r cell' + (positive ? ' pos' : '') +
                    '" style="background: ' + tint + '">' + sign +
                    Math.round(Math.abs(v)).toLocaleString('en-GB') + '</td>';
            }).join('');

            return '<tr class="rule-row' + (open ? ' open' : '') +
                    '" data-cat-row="' + esc(r.name) + '">' +
                '<td class="desc"><span class="chev">' + (open ? '▾' : '▸') + '</span>' +
                    esc(r.name) + '</td>' + cells +
                '<td class="r">' + sparkline(values) + '</td>' +
                '<td class="r nowrap" title="' + active + ' of ' + periods.length +
                    ' periods had activity">' + money(avg) + '</td>' +
                '<td class="r nowrap"><strong>' + money(r.total) + '</strong></td></tr>' +
                (open ? breakdownDetail(r, periods.length + 4) : '');
        };

        const divider = income.length && spending.length
            ? '<tr class="divider"><td colspan="' + (periods.length + 4) + '">Income &amp; transfers</td></tr>'
            : '';

        if (!all.length) {
            $('txBody').innerHTML = '';
            $('emptyNote').hidden = false;
            $('emptyNote').textContent = 'No categories match "' + state.search + '".';
            return;
        }
        $('emptyNote').hidden = true;
        $('txBody').innerHTML = ordered(spending).map(renderRow).join('') +
            divider + ordered(income).map(renderRow).join('');

        Array.prototype.forEach.call($('txBody').querySelectorAll('[data-cat-row]'), function (row) {
            row.addEventListener('click', function () {
                const name = row.dataset.catRow;
                if (state.expandedCat.has(name)) state.expandedCat.delete(name);
                else state.expandedCat.add(name);
                render();
            });
        });
    }

    function breakdownDetail(row, colspan) {
        const MAX = 200;
        const hits = row.indexes.slice().reverse().slice(0, MAX);
        const body = hits.map(function (i) {
            const tx = state.txs[i], r = state.results[i];
            return '<tr><td class="nowrap">' + esc(Engine.fmtDate(tx.date)) + '</td>' +
                '<td>' + esc(tx.description) + '</td>' +
                '<td class="muted">' + (r.rule ? '<code>' + esc(r.rule.match) + '</code>' : '—') + '</td>' +
                '<td class="r nowrap ' + (tx.amount < 0 ? 'out' : 'in') + '">' + money(tx.amount) + '</td></tr>';
        }).join('');
        return '<tr class="detail"><td colspan="' + colspan + '"><div class="detail-box">' +
            '<div class="detail-head">' + row.indexes.length + ' transaction' +
            (row.indexes.length === 1 ? '' : 's') + ' · ' + money(row.total) + ' total' +
            (row.indexes.length > MAX ? ' · showing the newest ' + MAX : '') + '</div>' +
            '<table class="detail-table">' + body + '</table></div></td></tr>';
    }

    /** How many transactions each rule actually won, keyed by match+type. */
    function claimedCounts() {
        const m = new Map();
        state.results.forEach(function (r) {
            if (!r.rule) return;
            const k = r.rule.match + '\u0000' + r.rule.type;
            m.set(k, (m.get(k) || 0) + 1);
        });
        return m;
    }

    /** How many rows in view share this description - they move together. */
    function countSameInBatch(index) {
        const key = Engine.normalise(state.txs[index].description);
        let n = 0;
        state.txs.forEach(function (tx, i) {
            if (Engine.normalise(tx.description) === key &&
                (state.batchFrom === undefined || i >= state.batchFrom)) n++;
        });
        return n;
    }

    function renderCategoryList() {
        const cats = Array.from(new Set(state.model.rules.map(function (r) { return r.category; })))
            .filter(Boolean).sort();
        $('categoryList').innerHTML = cats.map(function (c) {
            return '<option value="' + esc(c) + '">';
        }).join('');
    }

    function renderHealth() {
        const claimed = claimedCounts();
        const dead = state.model.rules.filter(function (r) {
            return !claimed.has(r.match + '\u0000' + r.type);
        });
        const uncategorised = state.results.filter(function (r) { return !r.rule; }).length;
        const total = state.txs.length;
        const sum = Array.from(claimed.values()).reduce(function (a, b) { return a + b; }, 0);

        const breaks = Engine.balanceBreaks(state.model);
        $('healthGrid').innerHTML =
            card(breaks.length === 0 ? 'balanced' : breaks.length + ' breaks',
                 "the bank's running balance adds up",
                 breaks.length ? 'bad' : 'good') +
            card(sum + ' + ' + uncategorised + ' = ' + total,
                 'Claimed + uncategorised = total',
                 sum + uncategorised === total ? 'good' : 'bad') +
            card(dead.length, 'rules that win nothing', dead.length ? 'warn' : 'good') +
            card(state.results.filter(function (r) { return r.confidence === 'review'; }).length,
                 'matches worth checking', '') +
            card(state.results.filter(function (r) { return r.conflict; }).length,
                 'scoring conflicts', state.results.some(function (r) { return r.conflict; }) ? 'bad' : 'good');

        $('balanceBox').hidden = !breaks.length;
        if (breaks.length) {
            $('balanceCount').textContent = breaks.length;
            $('balanceList').innerHTML = breaks.slice(0, 40).map(function (b) {
                return '<div><strong>' + esc(Engine.fmtDate(b.date)) + '</strong> ' +
                    esc(b.description) + ' — balance reads ' + money(b.actual) +
                    ', expected ' + money(b.expected) +
                    ' <span class="muted">(off by ' + money(b.diff) + ')</span></div>';
            }).join('') +
            '<div class="muted">A break means a row is missing or duplicated at that point — ' +
            'or the sheet was edited by hand there.</div>';
        }

        const cols = Engine.collisions(state.model.rules);
        $('collisionBox').hidden = !cols.length;
        if (cols.length) {
            $('collisionCount').textContent = cols.length;
            $('collisionList').innerHTML = cols.slice(0, 60).map(function (c) {
                return '<div><code>' + esc(c.shorter.match) + '</code> → ' + esc(c.shorter.category) +
                    '  <span class="muted">is inside</span>  <code>' + esc(c.longer.match) +
                    '</code> → ' + esc(c.longer.category) +
                    ' <span class="muted">(longer wins)</span></div>';
            }).join('');
        }

        renderLegacy();
    }

    /** Migration aid: diff the old workbook's categories against the new engine. */
    function renderLegacy() {
        const legacy = state.model.legacy;
        const box = $('legacyBox');
        if (!legacy || !legacy.length) { box.hidden = true; return; }

        const was = new Map();
        legacy.forEach(function (l) { was.set(Engine.normalise(l.description), l.category); });

        const diffs = [];
        const seen = new Set();
        state.txs.forEach(function (tx, i) {
            const key = Engine.normalise(tx.description);
            if (seen.has(key) || !was.has(key)) return;
            seen.add(key);
            const before = was.get(key), now = state.results[i].category;
            if (Engine.normalise(before) !== Engine.normalise(now)) {
                diffs.push({ desc: tx.description, before: before, now: now });
            }
        });

        box.hidden = false;
        $('legacyCount').textContent = diffs.length;
        $('legacyList').innerHTML = diffs.length
            ? diffs.map(function (d) {
                return '<div><code>' + esc(d.desc) + '</code> <span class="muted">was</span> ' +
                    esc(d.before || '—') + ' <span class="muted">→ now</span> <strong>' +
                    esc(d.now) + '</strong></div>';
            }).join('')
            : '<div class="muted">Nothing changed — every description resolves the same way it used to.</div>';
    }

    function card(value, label, tone) {
        return '<div class="hcard ' + tone + '"><div class="hval">' + esc(value) +
            '</div><div class="hlabel">' + esc(label) + '</div></div>';
    }

    // ---------------------------------------------------------- rule editor

    /** A sensible starting point for the match string: drop trailing ref codes. */
    function suggestMatch(description) {
        const words = Engine.normalise(description).split(' ');
        const kept = [];
        for (const w of words) {
            if (/^\d+$/.test(w) || /^[A-Z]*\d[A-Z0-9]*$/.test(w)) break;
            kept.push(w);
            if (kept.length === 3) break;
        }
        return (kept.length ? kept : words.slice(0, 2)).join(' ');
    }

    /**
     * Opened from a transaction (or with null, for a blank rule).
     *
     * When a rule already matched, the default is a NEW rule rather than an edit:
     * you are usually here because this transaction is wrong, not because the
     * whole rule is.  Editing the matched rule is one radio button away, and the
     * impact preview tells you what either choice does.
     */
    function openEditor(txIndex) {
        const tx = txIndex === null ? null : state.txs[txIndex];
        const res = txIndex === null ? { rule: null } : state.results[txIndex];
        const existing = res.rule ? indexOfRule(res.rule) : -1;
        state.editing = { txIndex: txIndex, ruleIndex: -1, existing: existing, mode: 'new' };

        $('modeWrap').hidden = existing === -1;
        if (existing !== -1) {
            $('modeExisting').textContent = res.rule.match + '  ·  ' + res.rule.type +
                '  →  ' + res.rule.category;
            document.querySelector('input[name="ruleMode"][value="new"]').checked = true;
        }

        $('ruleTitle').textContent = tx ? (res.rule ? 'This one is wrong' : 'Categorise this') : 'New rule';
        $('ruleForDesc').innerHTML = tx
            ? '<code>' + esc(tx.description) + '</code> · ' + esc(Engine.fmtDate(tx.date)) +
              ' · ' + money(tx.amount)
            : 'Applies to every transaction it matches.';

        $('ruleMatch').value = tx ? suggestMatch(tx.description) : '';
        $('ruleType').value = 'contains';
        $('ruleCategory').value = '';
        $('ruleComment').value = '';
        $('ruleWhere').value = '';
        $('ruleScope').checked = false;
        $('ruleConfirm').checked = false;

        const from = scopeDate();
        $('scopeWrap').hidden = !from;
        $('scopeLabel').textContent = from
            ? 'Apply from ' + Engine.fmtDate(from) + ' onwards only (leaves older transactions alone)'
            : '';

        $('ruleModal').hidden = false;
        $('ruleMatch').focus();
        updateImpact();
    }

    /** Opened from the Rules tab: always an edit of that exact rule. */
    function openRuleEditor(ruleIndex) {
        const r = state.model.rules[ruleIndex];
        state.editing = { txIndex: null, ruleIndex: ruleIndex, existing: ruleIndex, mode: 'edit' };

        $('modeWrap').hidden = true;
        $('ruleTitle').textContent = 'Edit rule';
        $('ruleForDesc').innerHTML = 'Changes apply everywhere this rule matches.';
        $('ruleMatch').value = r.match;
        $('ruleType').value = r.type;
        $('ruleCategory').value = r.category;
        $('ruleComment').value = r.comment || '';
        $('ruleWhere').value = r.where || '';
        $('ruleScope').checked = false;
        $('ruleConfirm').checked = false;
        $('scopeWrap').hidden = !scopeDate();
        $('ruleModal').hidden = false;
        $('ruleMatch').focus();
        updateImpact();
    }

    function indexOfRule(rule) {
        for (let i = 0; i < state.model.rules.length; i++) {
            if (state.model.rules[i].match === rule.match &&
                state.model.rules[i].type === rule.type) return i;
        }
        return -1;
    }

    /** Earliest date of the current batch, used by the "onwards only" scope. */
    function scopeDate() {
        if (state.batchFrom === undefined) return null;
        const tx = state.txs[state.batchFrom];
        return tx ? tx.date : null;
    }

    function setMode(mode) {
        if (!state.editing) return;
        state.editing.mode = mode;
        state.editing.ruleIndex = mode === 'edit' ? state.editing.existing : -1;
        if (mode === 'edit' && state.editing.existing !== -1) {
            const r = state.model.rules[state.editing.existing];
            $('ruleMatch').value = r.match;
            $('ruleType').value = r.type;
            $('ruleCategory').value = r.category;
            $('ruleComment').value = r.comment || '';
            $('ruleWhere').value = r.where || '';
        } else if (state.editing.txIndex !== null) {
            $('ruleMatch').value = suggestMatch(state.txs[state.editing.txIndex].description);
            $('ruleType').value = 'contains';
            $('ruleCategory').value = '';
        }
        updateImpact();
    }

    function draftRule() {
        const from = $('ruleScope').checked ? scopeDate() : null;
        return {
            match: $('ruleMatch').value.trim(),
            type: $('ruleType').value,
            category: $('ruleCategory').value.trim() || 'Uncategorized',
            from: from ? Engine.fmtDate(from) : null,
            to: null,
            comment: $('ruleComment').value.trim(),
            where: $('ruleWhere').value.trim()
        };
    }

    let impactTimer = null;
    function updateImpact() {
        clearTimeout(impactTimer);
        impactTimer = setTimeout(function () {
            const draft = draftRule();
            const box = $('impact');
            if (!draft.match) {
                box.innerHTML = '<p class="muted">Type something to match on.</p>';
                return;
            }

            // Exclude the rule being edited, so its own rows don't read as changes
            const others = state.rules.filter(function (_r, i) {
                return i !== state.editing.ruleIndex;
            });
            const baseline = state.editing.ruleIndex === -1
                ? state.results
                : Engine.classifyAll(state.txs, others, state.batchFrom);

            const p = Engine.previewRule(draft, state.txs, baseline, state.batchFrom);

            const rows = [
                line(p.matched, 'transactions match this rule', ''),
                state.batchFrom !== undefined ? line(p.inBatch, 'of them are in this batch', '') : '',
                line(p.fromUncategorised, 'currently uncategorised', ''),
                line(p.changesBefore, 'already-categorised transactions would CHANGE',
                     p.changesBefore ? 'bad' : 'good')
            ].join('');

            let detail = '';
            if (p.changes.length) {
                const sorted = p.changes.slice().sort(function (a, b) { return b.past - a.past; });
                detail = '<div class="changes">' + sorted.slice(0, 12).map(function (ch) {
                    return '<div class="' + (ch.past ? 'past' : '') + '"><code>' +
                        esc(ch.tx.description) + '</code> ' + esc(Engine.fmtDate(ch.tx.date)) +
                        ' · <strong>' + esc(ch.from) + '</strong> → ' + esc(draft.category) +
                        ' <span class="muted">(was ' + esc(ch.rule) +
                        (ch.past ? ', before this batch' : '') + ')</span></div>';
                }).join('') + (p.changes.length > 12
                    ? '<div class="muted">…and ' + (p.changes.length - 12) + ' more</div>' : '') + '</div>';
            }

            box.innerHTML = '<div class="impact-grid">' + rows + '</div>' + detail;
            $('confirmWrap').hidden = p.changesBefore === 0;
            validate();
        }, 120);
    }

    function line(n, label, tone) {
        return '<div class="iline ' + tone + '"><span class="inum">' + n + '</span>' + esc(label) + '</div>';
    }

    function validate() {
        const needsConfirm = !$('confirmWrap').hidden && !$('ruleConfirm').checked;
        const draft = draftRule();
        $('ruleSave').disabled = !draft.match || !draft.category ||
            draft.category === 'Uncategorized' || needsConfirm;
    }

    function closeEditor() {
        $('ruleModal').hidden = true;
        state.editing = null;
    }

    function saveRule() {
        const draft = draftRule();
        const editing = state.editing.ruleIndex >= 0;
        snapshot(editing ? 'edit "' + draft.match + '"' : 'add "' + draft.match + '"');
        if (editing) state.model.rules[state.editing.ruleIndex] = draft;
        else state.model.rules.push(draft);
        recompute();
        markDirty((editing ? 'Edited' : 'Added') + ' rule "' + draft.match + '"');
        closeEditor();
        render();
    }

    /**
     * "Looks right" acknowledges the row and touches NOTHING else.
     *
     * It deliberately writes no rule: once the workbook is saved these rows are
     * history, and history is never flagged again, so there is nothing to
     * persist.  All rows sharing the description clear together, because they
     * are one decision.
     */
    function dismissRow(txIndex) {
        snapshot('accept "' + state.txs[txIndex].description + '"');
        state.dismissed.add(Engine.normalise(state.txs[txIndex].description));
        render();
    }

    /** Nod at everything still flagged that a rule already handled. */
    function dismissAll() {
        snapshot('accept all');
        state.txs.forEach(function (tx, i) {
            if (isFlagged(i) && state.results[i].rule && !state.results[i].conflict) {
                state.dismissed.add(Engine.normalise(tx.description));
            }
        });
        render();
    }

    // ----------------------------------------------------------------- save

    /** finance.xlsx -> finance-2026-08-15.xlsx, so saves build up a history
     *  instead of overwriting the only copy.  Re-saving the same day replaces
     *  the date rather than stacking another one on. */
    function datedName(filename) {
        const base = filename.replace(/\.xlsx?$/i, '').replace(/-\d{4}-\d{2}-\d{2}$/, '');
        const d = new Date();
        const stamp = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') +
                      '-' + String(d.getDate()).padStart(2, '0');
        return base + '-' + stamp + '.xlsx';
    }

    function saveWorkbook() {
        const out = Engine.writeWorkbook(state.model, state.txs, state.results, state.importLog);
        const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = datedName(state.filename);
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(function () { URL.revokeObjectURL(url); }, 5000);

        state.dirty = false;
        state.importLog = null;
        $('dirtyDot').hidden = true;
        $('saveBtn').disabled = true;
        $('saveNote').textContent = 'Saved as ' + datedName(state.filename);
    }


    // ----------------------------------------------------------------- markup

    const MARKUP = [
        '<section class="review" id="review" hidden>',
        '  <div class="review-head">',
        '    <div class="tabs" id="tabs"></div>',
        '    <div class="review-tools">',
        '      <input type="search" id="searchBox" placeholder="Search…" autocomplete="off">',
        '      <div class="review-stats" id="reviewStats"></div>',
        '    </div>',
        '  </div>',
        '  <div class="table-scroll">',
        '    <table class="tx"><thead id="txHead"></thead><tbody id="txBody"></tbody></table>',
        '  </div>',
        '  <p class="empty" id="emptyNote" hidden></p>',
        '</section>',
        '<section class="health" id="health" hidden>',
        '  <h2>Rule health</h2>',
        '  <div class="health-grid" id="healthGrid"></div>',
        '  <details id="balanceBox" hidden>',
        '    <summary><span id="balanceCount"></span> breaks in the bank\'s running balance</summary>',
        '    <div class="collision-list" id="balanceList"></div>',
        '  </details>',
        '  <details id="collisionBox" hidden>',
        '    <summary><span id="collisionCount"></span> rule pairs where one match contains another</summary>',
        '    <div class="collision-list" id="collisionList"></div>',
        '  </details>',
        '  <details id="legacyBox" hidden>',
        '    <summary><span id="legacyCount"></span> descriptions categorised differently than your old workbook</summary>',
        '    <div class="collision-list" id="legacyList"></div>',
        '  </details>',
        '</section>',
        '<div class="savebar" id="savebar" hidden>',
        '  <span class="dirty" id="dirtyDot" hidden>●</span>',
        '  <span id="saveNote">No changes yet</span>',
        '  <button id="undoBtn" class="ghost" hidden>↶ Undo</button>',
        '  <button id="saveBtn" class="primary" disabled>Save workbook</button>',
        '</div>',
        '<div class="modal" id="ruleModal" hidden>',
        '  <div class="modal-card">',
        '    <h3 id="ruleTitle">New rule</h3>',
        '    <p class="modal-desc" id="ruleForDesc"></p>',
        '    <div class="mode" id="modeWrap" hidden>',
        '      <label><input type="radio" name="ruleMode" value="new" checked>',
        '        <span>Write a <strong>new</strong> rule for this transaction</span></label>',
        '      <label><input type="radio" name="ruleMode" value="edit">',
        '        <span>Edit the rule that matched: <code id="modeExisting"></code></span></label>',
        '    </div>',
        '    <div class="field-row">',
        '      <label class="field grow"><span>Match</span>',
        '        <input type="text" id="ruleMatch" autocomplete="off"></label>',
        '      <label class="field"><span>Type</span>',
        '        <select id="ruleType">',
        '          <option value="contains">contains</option>',
        '          <option value="exact">exact</option>',
        '          <option value="wildcard">wildcard (* ?)</option>',
        '        </select></label>',
        '    </div>',
        '    <div class="field-row">',
        '      <label class="field grow"><span>Category</span>',
        '        <input type="text" id="ruleCategory" list="categoryList" autocomplete="off">',
        '        <datalist id="categoryList"></datalist></label>',
        '      <label class="field"><span>Where</span>',
        '        <input type="text" id="ruleWhere" autocomplete="off"></label>',
        '    </div>',
        '    <label class="field"><span>Comment</span>',
        '      <input type="text" id="ruleComment" autocomplete="off"></label>',
        '    <label class="check" id="scopeWrap">',
        '      <input type="checkbox" id="ruleScope"><span id="scopeLabel"></span></label>',
        '    <div class="impact" id="impact"></div>',
        '    <label class="check danger" id="confirmWrap" hidden>',
        '      <input type="checkbox" id="ruleConfirm">',
        '      <span>I understand this changes transactions I already categorised</span></label>',
        '    <div class="modal-actions">',
        '      <button class="ghost" id="ruleCancel">Cancel</button>',
        '      <button class="primary" id="ruleSave">Save rule</button>',
        '    </div>',
        '  </div>',
        '</div>'
    ].join('\n');

    // ------------------------------------------------------------------ api

    function mount(rootElement) {
        root = rootElement;
        root.innerHTML = MARKUP;

        ['ruleMatch', 'ruleType', 'ruleCategory', 'ruleScope'].forEach(function (id) {
            $(id).addEventListener('input', updateImpact);
            $(id).addEventListener('change', updateImpact);
        });
        $('ruleConfirm').addEventListener('change', validate);
        Array.prototype.forEach.call(root.querySelectorAll('input[name="ruleMode"]'), function (radio) {
            radio.addEventListener('change', function () { setMode(radio.value); });
        });
        $('undoBtn').addEventListener('click', undo);
        $('searchBox').addEventListener('input', function () {
            state.search = $('searchBox').value.trim();
            render();
        });
        $('ruleCancel').addEventListener('click', closeEditor);
        $('ruleSave').addEventListener('click', saveRule);
        $('ruleModal').addEventListener('click', function (e) {
            if (e.target === $('ruleModal')) closeEditor();
        });
        $('saveBtn').addEventListener('click', saveWorkbook);

        document.addEventListener('keydown', function (e) {
            if (e.key !== 'Escape' || !root) return;
            if (!$('ruleModal').hidden) closeEditor();
            else if (state.search) { state.search = ''; $('searchBox').value = ''; render(); }
        });
        window.addEventListener('beforeunload', function (e) {
            if (state.dirty) { e.preventDefault(); e.returnValue = ''; }
        });
    }

    function load(model, filename) {
        state.model = model;
        state.filename = filename || 'finance.xlsx';
        state.newKeys = new Set();
        state.dismissed = new Set();
        state.undo = [];
        state.expanded = new Set();
        state.expandedCat = new Set();
        state.batchFrom = undefined;
        state.importLog = null;
        state.filter = 'needs';
        state.search = '';
        state.dirty = false;

        $('searchBox').value = '';
        $('review').hidden = false;
        $('health').hidden = false;
        $('savebar').hidden = false;
        document.body.classList.add('has-savebar');
        $('dirtyDot').hidden = true;
        $('saveBtn').disabled = true;
        $('saveNote').textContent = 'No changes yet';

        recompute();
        render();
    }

    /** Append imported rows, mark them as this session's batch, show them. */
    function applyImport(opts) {
        const bank = state.model.banks[opts.sheet];
        const offset = bank.rows.length;
        opts.rows.forEach(function (row, i) {
            bank.rows.push(row);
            state.newKeys.add(opts.sheet + ':' + (offset + i));
        });
        state.applied = { sheet: opts.sheet, count: opts.rows.length };
        state.importLog = opts.importLog || null;
        recompute();
        state.filter = 'batch';
        markDirty(opts.rows.length + ' transactions imported');
        render();
    }

    /** Take the last import back out, so the account can be re-picked. */
    function undoImport() {
        if (!state.applied) return;
        const bank = state.model.banks[state.applied.sheet];
        const from = bank.rows.length - state.applied.count;
        bank.rows.splice(from, state.applied.count);
        for (let i = from; i < from + state.applied.count; i++) {
            state.newKeys.delete(state.applied.sheet + ':' + i);
        }
        state.applied = null;
        state.batchFrom = undefined;
    }

    return {
        mount: mount,
        load: load,
        model: function () { return state.model; },
        applyImport: applyImport,
        undoImport: undoImport,
        transactions: function () { return { txs: state.txs, results: state.results }; },
        onChange: function (fn) { changeHandlers.push(fn); },
        isLoaded: function () { return !!state.model; }
    };
})();

if (typeof module !== 'undefined' && module.exports) module.exports = Review;
