"use strict";

const learning = {view: null, poll: null, issues: [], suite: null, note: null};
const metricNames = {
  opening_revenue: 'Opening revenue', new: 'New', upsell: 'Upsell',
  downsell: 'Downsell', cross_sell: 'Cross-sell', churn: 'Churn',
  closing_revenue: 'Closing revenue', gross_retention: 'Gross retention',
  net_retention: 'Net retention', ttm_gross_retention: 'TTM gross retention'
};

function exitLearning() {
  learning.view = null;
  clearTimeout(learning.poll);
  $('lab-workspace').hidden = true;
  document.querySelector('h1').textContent = 'Understand the revenue story.';
  document.querySelector('.subtitle').textContent = 'Build the customer cube. Follow the evidence. Make the call.';
  document.querySelectorAll('.nav-item').forEach(b => b.classList.toggle('active', b.id === 'workspace'));
}

const analysisLoadRun = loadRun;
loadRun = async function(id) { exitLearning(); await analysisLoadRun(id); };
const analysisNewRun = newRun;
newRun = function() { exitLearning(); analysisNewRun(); };
$('new-run').onclick = newRun;
$('workspace').onclick = newRun;

async function openLearning(view) {
  try {
    state.loadVersion++;
    clearTimeout(state.poll);
    clearTimeout(learning.poll);
    learning.view = view;
    for (const id of ['setup', 'progress', 'empty', 'results']) $(id).hidden = true;
    $('lab-workspace').hidden = false;
    document.querySelectorAll('.lab-panel').forEach(p => p.hidden = p.id !== 'lab-' + view);
    document.querySelectorAll('.nav-item').forEach(b => b.classList.toggle('active', b.id === ({compare: 'compare-nav', issues: 'issues-nav', cases: 'cases-nav'}[view])));
    document.querySelector('h1').textContent = {compare: 'Measure the differences.', issues: 'Make failures useful.', cases: 'Keep the lessons.'}[view];
    document.querySelector('.subtitle').textContent = {compare: 'Same data. Same task. Inspectable results.', issues: 'Track each issue without rewriting the original answer.', cases: 'Your reviewed answers become the next test.'}[view];
    notice('');
    if (view === 'compare') {
      const selected = $('comparison-source').value || state.run?.id;
      const runs = await api('/api/runs');
      $('comparison-source').innerHTML = runs.filter(r => r.status === 'complete' && r.amount_mode === 'period_sum')
        .map(r => `<option value="${r.id}">${esc(r.filename)} · ${esc(r.mode)} · ${r.id.slice(0,6)}</option>`).join('');
      if (selected && runs.some(r => r.id === selected)) $('comparison-source').value = selected;
      if (!$('comparison-models').value && state.config?.ai.model) $('comparison-models').value = state.config.ai.model + ', ';
      await refreshComparisons();
    } else if (view === 'issues') {
      learning.issues = await api('/api/issues');
      renderIssues();
    } else {
      learning.suite = await api('/api/cases');
      renderCases();
    }
  } catch (e) { notice(e.message); }
}

$('compare-nav').onclick = () => openLearning('compare');
$('issues-nav').onclick = () => openLearning('issues');
$('cases-nav').onclick = () => openLearning('cases');
document.querySelectorAll('[data-lab-view]').forEach(b => b.onclick = () => openLearning(b.dataset.labView));

async function refreshComparisons() {
  const batches = await api('/api/comparisons');
  if (learning.view !== 'compare') return;
  $('comparisons-list').innerHTML = batches.length ? batches.map(batch => `
    <article class="card lab-item">
      <div class="section-heading"><div><h3>${esc(batch.filename)}</h3><p>Source ${batch.source_sha256.slice(0,12)} · ${esc(batch.status)} · ${esc(new Date(batch.created_at).toLocaleString())}</p></div>
      <span class="status ${batch.status === 'complete' ? 'pass' : 'review'}">${esc(batch.status)}</span></div>
      <div class="table-wrap"><table><thead><tr><th>Model</th><th>Status</th><th>Score / 100</th><th>Failures</th><th>Model seconds</th><th>Total seconds</th><th>Inspect</th></tr></thead><tbody>
      ${batch.candidates.map(c => `<tr><td>${esc(c.model)}</td><td>${esc(c.status)}</td><td>${c.score == null ? 'n.a.' : c.score.toFixed(2)}</td><td>${num(c.failures)}</td><td>${c.model_seconds ?? 'n.a.'}</td><td>${c.total_seconds ?? 'n.a.'}</td><td><button class="row-button" data-open-run="${c.id}">Open run ↗</button></td></tr>${c.error ? `<tr><td colspan="7">${esc(c.error)}</td></tr>` : ''}`).join('')}
      </tbody></table></div><p class="help">${batch.same_prompt === true ? 'Recorded prompt hashes match.' : batch.same_prompt === false ? 'Prompt hashes differ: inspect the run manifests before interpreting this comparison.' : 'Prompt hashes will appear when model calls start.'} ${esc(batch.caveat)}</p>
    </article>`).join('') : '<div class="empty-small">No comparisons yet. Select a saved analysis and enter two model IDs above.</div>';
  bindOpenRuns($('comparisons-list'));
  if (batches.some(b => b.status === 'running')) learning.poll = setTimeout(() => refreshComparisons().catch(e => notice(e.message)), 2000);
}

$('comparison-form').onsubmit = async event => {
  event.preventDefault();
  const button = event.submitter;
  try {
    button.disabled = true;
    const models = $('comparison-models').value.split(/[\n,]/).map(s => s.trim()).filter(Boolean);
    await api('/api/comparisons', {method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({source_run_id: $('comparison-source').value, models})});
    notice('Comparison started. Each model gets a separate saved run.');
    await refreshComparisons();
  } catch (e) { notice(e.message); } finally { button.disabled = false; }
};

function bindOpenRuns(container) {
  container.querySelectorAll('[data-open-run]').forEach(b => b.onclick = () => loadRun(b.dataset.openRun).catch(e => notice(e.message)));
}

function renderIssues() {
  const status = $('issue-status-filter').value;
  const category = $('issue-category-filter').value;
  const issues = learning.issues.filter(i => (status === 'all' || i.status === status) && (category === 'all' || i.category === category));
  $('issue-total').textContent = `${issues.length} issues · ${learning.issues.filter(i => !['resolved','dismissed'].includes(i.status)).length} open or investigating`;
  $('issues-list').innerHTML = issues.length ? issues.slice(0,200).map(issue => `
    <article class="issue-item"><div class="section-heading"><div><h3>${esc(issue.category)} · ${esc(issue.metric || keyText(issue.row_key))}</h3>
    <p>${esc(issue.model || 'Calculated analysis')} · ${esc(issue.filename)} · ${issue.run_id.slice(0,8)}</p></div>
    <span class="status ${['resolved','dismissed'].includes(issue.status) ? 'pass' : 'review'}">${esc(issue.status)}</span></div>
    <p>${esc(issue.detail)}</p>${issue.metric ? `<p class="help">${esc(keyText(issue.row_key))} · engine: ${esc(issue.expected)} · model: ${esc(issue.actual)}</p>` : ''}
    <button class="evidence-link" data-open-run="${issue.run_id}">Open original run ↗</button>
    ${issue.row_key ? `<button class="evidence-link" data-review-issue="${issue.id}">Review this output row ↗</button>` : ''}
    <details><summary>Update status / inspect history</summary><form data-issue-form="${issue.id}" class="lab-form"><label>Status<select name="status">${['open','investigating','resolved','dismissed'].map(s => `<option ${s === issue.status ? 'selected' : ''}>${s}</option>`).join('')}</select></label><label>Reason<textarea name="note" maxlength="4000" rows="2" required placeholder="What did you check or correct?"></textarea></label><button class="secondary" type="submit">Save issue update</button></form>
    ${issue.history.map(h => `<div class="issue-event"><strong>${esc(h.status)}</strong> · ${esc(new Date(h.created_at).toLocaleString())}<p>${esc(h.note)}</p></div>`).join('') || '<p class="help">No status changes yet.</p>'}</details></article>`).join('') : '<p class="empty-small">No issues match. Flag a review note, or run a model evaluation to collect numerical failures.</p>';
  if (issues.length > 200) $('issues-list').insertAdjacentHTML('beforeend','<p class="help">Showing the first 200. Export learning history for all issues.</p>');
  bindOpenRuns($('issues-list'));
  $('issues-list').querySelectorAll('[data-review-issue]').forEach(button => button.onclick = async () => {
    try {
      const issue = learning.issues.find(i => i.id === button.dataset.reviewIssue);
      await loadRun(issue.run_id);
      state.reviewKey = issue.row_key;
      renderNotes();
      showTab('review');
    } catch (e) { notice(e.message); }
  });
  $('issues-list').querySelectorAll('[data-issue-form]').forEach(form => form.onsubmit = async event => {
    event.preventDefault();
    const button = event.submitter;
    try {
      button.disabled = true;
      await api('/api/issues/' + form.dataset.issueForm, {method: 'POST', headers: {'Content-Type':'application/json'},
        body: JSON.stringify({status: form.elements.status.value, note: form.elements.note.value})});
      learning.issues = await api('/api/issues');
      renderIssues();
      notice('Issue history saved. Original model output and review notes are unchanged.');
    } catch (e) { notice(e.message); } finally { button.disabled = false; }
  });
}
for (const id of ['issue-status-filter','issue-category-filter']) $(id).onchange = renderIssues;

function renderCases() {
  const {cases, evaluations} = learning.suite;
  $('cases-list').innerHTML = cases.length ? cases.map(c => {
    const matches = evaluations.filter(e => e.case_id === c.id);
    const assertionLabel = c.assertion === 'row_absent' ? 'Row must be absent' : c.assertion === 'row_present' ? 'Row must appear exactly once' : metricNames[c.metric];
    return `<article class="card lab-item"><div class="section-heading"><div><h3>${esc(c.title)}</h3><p>${esc(keyText(c.row_key))} · ${esc(assertionLabel)}</p></div><span class="step-label">${c.id.slice(0,10)}</span></div>
      <p><strong>Expected: ${esc(c.metric ? c.expected === null ? 'null (unavailable)' : c.expected : assertionLabel)}</strong></p><p>${esc(c.rationale)}</p>
      <p class="help">Reviewer-entered answer · source ${c.source_sha256.slice(0,12)} · absolute / relative tolerance 1e-6 · ${esc(c.amount_mode)}</p>
      <form class="case-rerun" data-case-rerun="${c.id}"><label>Model ID<input name="model" value="${esc(state.config?.ai.model || '')}" maxlength="120" required></label><button class="secondary" type="submit">Rerun source against this test ↗</button></form>
      <p class="help">Starts one API call. The candidate does not receive this expected answer or your review notes.</p>
      <div class="table-wrap"><table><thead><tr><th>Run / model</th><th>Result</th><th>Actual</th><th>Reason</th><th>Inspect</th></tr></thead><tbody>
      ${matches.slice(0,20).map(e => `<tr><td>${esc(e.model || e.mode)} · ${e.run_id.slice(0,6)}</td><td><span class="status ${e.status === 'pass' ? 'pass' : e.status === 'fail' ? 'fail' : 'review'}">${esc(e.status)}</span></td><td>${esc(e.actual === null ? 'null / unavailable' : e.actual)}</td><td>${esc(e.reason)}</td><td><button class="row-button" data-open-run="${e.run_id}">Open ↗</button></td></tr>`).join('') || '<tr><td colspan="5">No matching runs yet.</td></tr>'}
      </tbody></table></div><p class="help">Calculated and AI-explanation runs show engine behavior; only model-calculation runs test independently generated numbers. The latest 20 matching runs are shown.</p></article>`;
  }).join('') : '<p class="empty-small">No tests yet. Inspect an output row, save a review note, and use its Create regression test button.</p>';
  bindOpenRuns($('cases-list'));
  $('cases-list').querySelectorAll('[data-case-rerun]').forEach(form => form.onsubmit = async event => {
    event.preventDefault();
    const button = event.submitter;
    try {
      button.disabled = true;
      const run = await api('/api/cases/' + form.dataset.caseRerun + '/run', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({model:form.elements.model.value.trim()})});
      await loadRun(run.id);
    } catch (e) { notice(e.message); } finally { button.disabled = false; }
  });
}

function openCaseForm(note) {
  learning.note = note;
  $('case-origin').textContent = keyText(note.row_key);
  $('case-title').value = note.note.slice(0,100);
  $('case-rationale').value = note.note;
  $('case-metric').innerHTML = Object.entries(metricNames).map(([k,v]) => `<option value="${k}" ${k === 'closing_revenue' ? 'selected' : ''}>${v}</option>`).join('');
  $('case-error').textContent = '';
  $('case-assertion').value = 'metric_equals';
  toggleAssertion();
  suggestExpected();
  $('case-dialog').showModal();
}
function suggestExpected() {
  const row = state.result.reference_rows.find(r => JSON.stringify(key(r)) === JSON.stringify(learning.note.row_key));
  $('case-expected').value = row ? String(row[$('case-metric').value]) : '';
}
$('case-metric').onchange = suggestExpected;
function toggleAssertion() {
  const numeric = $('case-assertion').value === 'metric_equals';
  $('case-metric-label').hidden = !numeric;
  $('case-expected-label').hidden = !numeric;
  $('case-expected').required = numeric;
}
$('case-assertion').onchange = toggleAssertion;
$('case-close').onclick = () => $('case-dialog').close();
$('case-form').onsubmit = async event => {
  event.preventDefault();
  const button = event.submitter;
  try {
    button.disabled = true;
    const raw = $('case-expected').value.trim();
    const assertion = $('case-assertion').value;
    const expected = assertion !== 'metric_equals' || raw === 'null' ? null : Number(raw);
    if (assertion === 'metric_equals' && (!raw || (expected !== null && !Number.isFinite(expected)))) throw Error('Enter a number or the word null.');
    await api('/api/cases', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({run_id:state.run.id,
      note_id:learning.note.id, title:$('case-title').value, assertion, metric:$('case-metric').value, expected, rationale:$('case-rationale').value})});
    $('case-dialog').close();
    await openLearning('cases');
    notice('Regression case saved. Your expected answer is kept separate from the candidate prompt.');
  } catch (e) { $('case-error').textContent = e.message; } finally { button.disabled = false; }
};

const originalRenderNotes = renderNotes;
renderNotes = function() {
  originalRenderNotes();
  const elements = $('notes').querySelectorAll('.note');
  state.feedback.slice().reverse().forEach((note, i) => {
    if (!note.row_key) return;
    const button = document.createElement('button');
    button.className = 'secondary';
    button.textContent = 'Create regression test ↗';
    button.onclick = () => openCaseForm(note);
    elements[i].appendChild(button);
  });
};
