/**
 * Import page: upload a workbook, drop a statement on it, hand both to the
 * shared Review component.  Everything after the import lives in review.js,
 * which the dashboard mounts too.
 */
(function () {
    'use strict';

    const $ = function (id) { return document.getElementById(id); };

    function esc(s) {
        return String(s === null || s === undefined ? '' : s)
            .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    function wireDrop(zoneId, inputId, buttonId, handler) {
        const zone = $(zoneId), input = $(inputId);
        $(buttonId).addEventListener('click', function (e) { e.preventDefault(); input.click(); });
        zone.addEventListener('click', function (e) { if (e.target === zone || e.target.tagName === 'P') input.click(); });
        input.addEventListener('change', function () { if (input.files[0]) handler(input.files[0]); });
        ['dragenter', 'dragover'].forEach(function (ev) {
            zone.addEventListener(ev, function (e) { e.preventDefault(); zone.classList.add('over'); });
        });
        ['dragleave', 'drop'].forEach(function (ev) {
            zone.addEventListener(ev, function (e) { e.preventDefault(); zone.classList.remove('over'); });
        });
        zone.addEventListener('drop', function (e) {
            if (e.dataTransfer.files[0]) handler(e.dataTransfer.files[0]);
        });
    }

    // ------------------------------------------------------------- workbook

    function loadWorkbook(file) {
        const reader = new FileReader();
        reader.onload = function () {
            try {
                const model = Engine.readWorkbook(reader.result);
                Review.load(model, file.name);
                $('workbookInfo').hidden = false;
                $('workbookInfo').innerHTML = renderWorkbookInfo(file, model);
                $('step2').classList.remove('disabled');
            } catch (err) {
                $('workbookInfo').hidden = false;
                $('workbookInfo').innerHTML = '<p class="error">Could not read that workbook: ' +
                    esc(err.message) + '</p>';
            }
        };
        reader.readAsArrayBuffer(file);
    }

    function renderWorkbookInfo(file, model) {
        const accounts = model.accounts.map(function (a) {
            const n = model.banks[a.sheet].rows.length;
            return '<li><strong>' + esc(a.label) + '</strong> <span class="muted">(' +
                esc(a.sheet) + ')</span> — ' + n + ' transactions</li>';
        }).join('');
        return '<p class="ok">Loaded <strong>' + esc(file.name) + '</strong></p><ul class="accounts">' +
            accounts + '</ul><p class="muted">' + model.rules.length +
            ' rules · every category recomputed from scratch</p>';
    }

    // ------------------------------------------------------------ statement

    function loadStatement(file) {
        const reader = new FileReader();
        reader.onload = function () {
            const box = $('statementInfo');
            box.hidden = false;
            try {
                const statement = Engine.readStatement(reader.result, file.name);
                const candidates = Engine.detectAccount(statement, Review.model());
                const best = candidates[0];
                if (!best || best.score < 40) {
                    box.innerHTML = '<p class="error">These columns don\'t match any account sheet.</p>' +
                        '<p class="muted">File has: ' + esc(statement.headers.join(' · ')) + '</p>';
                    return;
                }
                applyStatement(statement, best.account, candidates, file);
            } catch (err) {
                box.innerHTML = '<p class="error">' + esc(err.message) + '</p>';
            }
        };
        reader.readAsText(file);
    }

    function applyStatement(statement, account, candidates, file) {
        Review.undoImport();
        const plan = Engine.planImport(statement, account, Review.model());
        const box = $('statementInfo');

        if (plan.at === null) {
            box.innerHTML = '<p class="error">Couldn\'t find your last stored transaction in this file.</p>' +
                '<p class="muted">Nothing was imported. Below: your last stored row in <strong>' +
                esc(account.label) + '</strong>, and the closest rows in the file. A ✗ marks the ' +
                'field that stops them matching.</p>' + renderDiagnosis(plan.diagnosis);
            return;
        }

        Review.applyImport({
            sheet: account.sheet,
            rows: plan.rows,
            importLog: [new Date().toISOString().slice(0, 19).replace('T', ' '),
                        account.sheet, file.name, plan.periodFrom, plan.periodTo,
                        plan.total, plan.rows.length, plan.skipped]
        });

        const detected = candidates.map(function (c) {
            return '<option value="' + esc(c.account.sheet) + '"' +
                (c.account.sheet === account.sheet ? ' selected' : '') + '>' +
                esc(c.account.label) + ' (' + c.score + '% column match)</option>';
        }).join('');

        box.innerHTML =
            '<p class="ok">Loaded <strong>' + esc(file.name) + '</strong></p>' +
            (statement.banner ? '<p class="muted">' + esc(statement.banner) + '</p>' : '') +
            '<div class="route">Appended to <select id="routeSelect">' + detected + '</select></div>' +
            '<p><strong>' + plan.rows.length + '</strong> new transactions added · ' +
            plan.skipped + ' already stored, skipped</p>';

        $('routeSelect').addEventListener('change', function (e) {
            const picked = Review.model().accounts.filter(function (a) {
                return a.sheet === e.target.value;
            })[0];
            if (picked) applyStatement(statement, picked, candidates, file);
        });

    }

    /** Field-by-field comparison, so a formatting mismatch is obvious. */
    function renderDiagnosis(d) {
        if (!d) return '';
        const f = function (v, ok) {
            return '<td class="' + (ok === null ? '' : (ok ? 'match' : 'nomatch')) + '">' +
                (ok === null ? '' : (ok ? '✓ ' : '✗ ')) + esc(v === null || v === '' ? '(empty)' : v) + '</td>';
        };
        const L = d.last;
        let html = '<table class="diag"><tr><th></th><th>Date</th><th>Time</th>' +
            '<th>Description</th><th>Amount</th></tr>' +
            '<tr><th>stored</th>' + f(Engine.fmtDate(L.date), null) + f(L.time, null) +
            f(L.description, null) + f(L.amount === null ? null : L.amount.toFixed(2), null) + '</tr>';

        if (!d.candidates.length) {
            html += '<tr><td colspan="5" class="muted">No row in the file resembles it at all — ' +
                'likely the wrong file, or the wrong account.</td></tr>';
        }
        d.candidates.forEach(function (c) {
            const v = c.values;
            html += '<tr><th>file row ' + (c.index + 1) + '</th>' +
                f(Engine.fmtDate(v.date), Engine.fmtDate(v.date) === Engine.fmtDate(L.date)) +
                f(v.time, Engine.fmtTime(v.time) === Engine.fmtTime(L.time)) +
                f(v.description, Engine.normalise(v.description) === Engine.normalise(L.description)) +
                f(v.amount === null ? null : v.amount.toFixed(2),
                  v.amount !== null && L.amount !== null && v.amount.toFixed(2) === L.amount.toFixed(2)) +
                '</tr>';
        });
        return html + '</table>';
    }

    // ----------------------------------------------------------------- init

    Review.mount(document.getElementById('reviewRoot'));

    wireDrop('dropWorkbook', 'fileWorkbook', 'pickWorkbook', loadWorkbook);
    wireDrop('dropStatement', 'fileStatement', 'pickStatement', function (file) {
        if (!Review.isLoaded()) return;
        loadStatement(file);
    });
})();
