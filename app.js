/* =====================================================
   Kiểm Phiếu Bầu Cử — Application Logic v2
   Tối ưu cho 4500+ phiếu
   ===================================================== */

// ============ DATA STORE ============
const appData = {
    quochoi: {
        label: 'Quốc Hội',
        totalCandidates: 5,
        electCount: 3,
        crossCount: 2,
        candidates: [
            'Trần Thị Phương Anh',
            'Nguyễn Trần Phương Hà',
            'Trương Vũ Yến Nhi',
            'Lê Minh Trí',
            'Vũ Hồng Văn'
        ],
        stacks: [{ id: 1, name: 'Xấp 1', ballots: [] }],
        currentStack: 0,
        started: false,
        _voteCounts: null,
        _allBallotsCache: null,
        _ballotsCacheDirty: true,
        _historyPage: 1
    },
    tinh: {
        label: 'HĐND Tỉnh/TP',
        totalCandidates: 5,
        electCount: 3,
        crossCount: 2,
        candidates: [
            'Châu Thị Chánh',
            'Nguyễn Đình Giáp',
            'Nguyễn Đức Hải',
            'Nguyễn Văn Phơi',
            'Trần Hoàng Tâm'
        ],
        stacks: [{ id: 1, name: 'Xấp 1', ballots: [] }],
        currentStack: 0,
        started: false,
        _voteCounts: null,
        _allBallotsCache: null,
        _ballotsCacheDirty: true,
        _historyPage: 1
    },
    phuong: {
        label: 'HĐND Phường/Xã',
        totalCandidates: 8,
        electCount: 5,
        crossCount: 3,
        candidates: [
            'Nguyễn Thị Ngọc Hòa',
            'Nguyễn Thị Ngọc Hương',
            'Lê Thị Kim Khánh',
            'Nguyễn Hữu Phước',
            'Lê Gia Quý',
            'Nguyễn Minh Tân',
            'Trương Minh Trung',
            'Nguyễn Vũ Trường'
        ],
        stacks: [{ id: 1, name: 'Xấp 1', ballots: [] }],
        currentStack: 0,
        started: false,
        _voteCounts: null,
        _allBallotsCache: null,
        _ballotsCacheDirty: true,
        _historyPage: 1
    }
};

const BALLOTS_PER_PAGE = 50;
let _saveTimeout = null;

// ============ INIT ============
document.addEventListener('DOMContentLoaded', () => {
    loadFromStorage();
    ['quochoi', 'tinh', 'phuong'].forEach(level => {
        initCandidatesList(level);
        updateConfig(level);
        if (appData[level].started) {
            restoreState(level);
        }
    });
    setupTabs();
});

// ============ TABS ============
function setupTabs() {
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const tab = btn.dataset.tab;
            document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
            document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
            btn.classList.add('active');
            document.getElementById(`content-${tab}`).classList.add('active');
        });
    });
}

// ============ CONFIG ============
function updateConfig(level) {
    const data = appData[level];
    const totalEl = document.getElementById(`totalCandidates-${level}`);
    const electEl = document.getElementById(`electCount-${level}`);
    const crossEl = document.getElementById(`crossCount-${level}`);

    data.totalCandidates = parseInt(totalEl.value) || 2;
    data.electCount = parseInt(electEl.value) || 1;

    if (data.electCount >= data.totalCandidates) {
        data.electCount = data.totalCandidates - 1;
        electEl.value = data.electCount;
    }

    data.crossCount = data.totalCandidates - data.electCount;
    crossEl.value = data.crossCount;

    data._voteCounts = null;
    initCandidatesList(level);
    debouncedSave();
}

function initCandidatesList(level) {
    const data = appData[level];
    const container = document.getElementById(`candidates-list-${level}`);
    const existingNames = data.candidates.slice();
    data.candidates = [];
    container.innerHTML = '';

    for (let i = 0; i < data.totalCandidates; i++) {
        const name = existingNames[i] || `Ứng cử viên ${i + 1}`;
        data.candidates.push(name);
        const item = document.createElement('div');
        item.className = 'candidate-item';
        item.innerHTML = `
            <span class="candidate-number">${i + 1}</span>
            <input type="text" class="candidate-input" value="${name}" 
                   data-level="${level}" data-index="${i}"
                   onchange="updateCandidateName(this)"
                   placeholder="Nhập tên ứng cử viên">
        `;
        container.appendChild(item);
    }
}

function updateCandidateName(input) {
    const level = input.dataset.level;
    const index = parseInt(input.dataset.index);
    appData[level].candidates[index] = input.value;
    debouncedSave();
    if (appData[level].started) updateResultsFromCache(level);
}

function toggleConfig(level) {
    const body = document.getElementById(`config-body-${level}`);
    const toggle = document.getElementById(`config-toggle-${level}`);
    if (body.style.display === 'none') {
        body.style.display = 'block';
        toggle.textContent = '▼ Thu gọn';
    } else {
        body.style.display = 'none';
        toggle.textContent = '▶ Mở rộng';
    }
}

// ============ START COUNTING ============
function startCounting(level) {
    const data = appData[level];
    const hasEmpty = data.candidates.some(c => !c.trim());
    if (hasEmpty) {
        showToast('Vui lòng nhập đầy đủ tên ứng cử viên!', 'warning');
        return;
    }

    data.started = true;
    initVoteCounts(level);

    document.getElementById(`input-panel-${level}`).style.display = 'block';
    document.getElementById(`results-panel-${level}`).style.display = 'block';

    updateStackSelect(level);
    updateResultsFromCache(level);
    renderBallotHistory(level);
    document.getElementById(`ballot-input-${level}`).focus();

    const body = document.getElementById(`config-body-${level}`);
    body.style.display = 'none';
    document.getElementById(`config-toggle-${level}`).textContent = '▶ Mở rộng';

    saveToStorage();
    showToast(`Bắt đầu kiểm phiếu ${data.label}!`, 'success');
}

function restoreState(level) {
    document.getElementById(`input-panel-${level}`).style.display = 'block';
    document.getElementById(`results-panel-${level}`).style.display = 'block';
    rebuildVoteCounts(level);
    updateStackSelect(level);
    updateResultsFromCache(level);
    renderBallotHistory(level);
    updateBallotCounter(level);

    const body = document.getElementById(`config-body-${level}`);
    body.style.display = 'none';
    document.getElementById(`config-toggle-${level}`).textContent = '▶ Mở rộng';
    document.getElementById(`totalCandidates-${level}`).value = appData[level].totalCandidates;
    document.getElementById(`electCount-${level}`).value = appData[level].electCount;
    document.getElementById(`crossCount-${level}`).value = appData[level].crossCount;
}

// ============ VOTE COUNTS CACHE (Incremental) ============
function initVoteCounts(level) {
    const data = appData[level];
    data._voteCounts = {
        total: 0, valid: 0, invalid: 0, none: 0,
        perCandidate: new Array(data.totalCandidates).fill(0)
    };
}

function rebuildVoteCounts(level) {
    initVoteCounts(level);
    getAllBallotsCached(level).forEach(b => addBallotToCache(level, b));
}

function addBallotToCache(level, ballot) {
    const data = appData[level];
    if (!data._voteCounts) initVoteCounts(level);
    const vc = data._voteCounts;
    vc.total++;
    if (ballot.status === 'valid') vc.valid++;
    else if (ballot.status === 'none') vc.none++;
    else vc.invalid++;

    if (ballot.status === 'valid' || ballot.status === 'none') {
        for (let i = 0; i < data.totalCandidates; i++) {
            if (!ballot.crossedOut.includes(i + 1)) {
                vc.perCandidate[i]++;
            }
        }
    }
}

function removeBallotFromCache(level, ballot) {
    const data = appData[level];
    if (!data._voteCounts) { rebuildVoteCounts(level); return; }
    const vc = data._voteCounts;
    vc.total--;
    if (ballot.status === 'valid') vc.valid--;
    else if (ballot.status === 'none') vc.none--;
    else vc.invalid--;

    if (ballot.status === 'valid' || ballot.status === 'none') {
        for (let i = 0; i < data.totalCandidates; i++) {
            if (!ballot.crossedOut.includes(i + 1)) {
                vc.perCandidate[i]--;
            }
        }
    }
}

// ============ CACHED getAllBallots ============
function getAllBallotsCached(level) {
    const data = appData[level];
    if (data._ballotsCacheDirty || !data._allBallotsCache) {
        data._allBallotsCache = data.stacks.reduce((arr, s) => arr.concat(s.ballots), []);
        data._ballotsCacheDirty = false;
    }
    return data._allBallotsCache;
}

function invalidateBallotCache(level) {
    appData[level]._ballotsCacheDirty = true;
    appData[level]._allBallotsCache = null;
}

// ============ STACK MANAGEMENT ============
function updateStackSelect(level) {
    const select = document.getElementById(`stack-select-${level}`);
    const data = appData[level];
    select.innerHTML = '';

    const allOpt = document.createElement('option');
    allOpt.value = 'all';
    allOpt.textContent = '📦 Tất cả xấp';
    select.appendChild(allOpt);

    data.stacks.forEach((stack, idx) => {
        const opt = document.createElement('option');
        opt.value = idx;
        opt.textContent = `${stack.name} (${stack.ballots.length})`;
        if (idx === data.currentStack) opt.selected = true;
        select.appendChild(opt);
    });
}

function switchStack(level) {
    const val = document.getElementById(`stack-select-${level}`).value;
    appData[level].currentStack = val === 'all' ? -1 : parseInt(val);
    appData[level]._historyPage = 1;
    renderBallotHistory(level);
    updateBallotCounter(level);
}

function addStack(level) {
    const data = appData[level];
    const newId = data.stacks.length + 1;
    data.stacks.push({ id: newId, name: `Xấp ${newId}`, ballots: [] });
    data.currentStack = data.stacks.length - 1;
    updateStackSelect(level);
    document.getElementById(`stack-select-${level}`).value = data.currentStack;
    data._historyPage = 1;
    renderBallotHistory(level);
    updateBallotCounter(level);
    debouncedSave();
    showToast(`Đã thêm Xấp ${newId}`, 'info');
}

// ============ BALLOT INPUT ============
function handleBallotInput(event, level) {
    if (event.key === 'Enter') {
        event.preventDefault();
        submitBallot(level);
    }
}

function submitBallot(level) {
    const data = appData[level];
    const input = document.getElementById(`ballot-input-${level}`);
    const rawValue = input.value.trim();

    let crossedOut = [];
    let status = 'valid';

    if (rawValue === '') {
        crossedOut = [];
        status = 'none';
    } else {
        let parts;
        // If input has separators (comma, space, semicolon), split normally
        if (/[\s,;.]/.test(rawValue)) {
            parts = rawValue.split(/[\s,;.]+/).filter(p => p.length > 0);
        } else {
            // No separators: split each character as individual digit
            // e.g. "13" -> ["1", "3"], "245" -> ["2", "4", "5"]
            parts = rawValue.split('').filter(p => p.length > 0);
        }
        crossedOut = parts.map(p => parseInt(p)).filter(n => !isNaN(n));

        const unique = [...new Set(crossedOut)];
        if (unique.length !== crossedOut.length) {
            showToast('Có số thứ tự bị trùng lặp!', 'warning');
            return;
        }
        crossedOut = unique;

        if (crossedOut.some(n => n < 1 || n > data.totalCandidates)) {
            showToast(`STT phải từ 1 đến ${data.totalCandidates}!`, 'error');
            return;
        }

        if (crossedOut.length !== data.crossCount) {
            status = 'invalid';
        }
    }

    let stackIdx = data.currentStack;
    if (stackIdx < 0) stackIdx = 0;
    if (stackIdx >= data.stacks.length) stackIdx = data.stacks.length - 1;

    const ballotNumber = getAllBallotsCached(level).length + 1;
    const ballot = {
        id: Date.now(),
        number: ballotNumber,
        crossedOut: crossedOut,
        status: status,
        stackId: data.stacks[stackIdx].id
    };

    data.stacks[stackIdx].ballots.push(ballot);
    invalidateBallotCache(level);
    addBallotToCache(level, ballot);

    input.value = '';
    input.focus();
    updateBallotCounter(level);
    updateStackSelect(level);
    data._historyPage = 1;
    renderBallotHistory(level);
    updateResultsFromCache(level);
    debouncedSave();

    if (typeof supabaseSync !== 'undefined' && supabaseSync.isConnected()) {
        supabaseSync.insertBallot(level, ballot, stackIdx);
    }

    if (status === 'valid') {
        showToast(`✓ Phiếu #${ballotNumber} — Hợp lệ (gạch: ${crossedOut.join(', ')})`, 'success');
    } else if (status === 'none') {
        showToast(`✓ Phiếu #${ballotNumber} — Bầu đủ (không gạch ai)`, 'info');
    } else {
        showToast(`⚠ Phiếu #${ballotNumber} — KHÔNG hợp lệ (gạch ${crossedOut.length}/${data.crossCount})`, 'warning');
    }
}

function undoLastBallot(level) {
    const data = appData[level];
    let stackIdx = data.currentStack;
    if (stackIdx < 0) {
        for (let i = data.stacks.length - 1; i >= 0; i--) {
            if (data.stacks[i].ballots.length > 0) { stackIdx = i; break; }
        }
    }
    if (stackIdx < 0 || !data.stacks[stackIdx] || data.stacks[stackIdx].ballots.length === 0) {
        showToast('Không có phiếu nào để hủy!', 'warning');
        return;
    }

    const removed = data.stacks[stackIdx].ballots.pop();
    invalidateBallotCache(level);
    removeBallotFromCache(level, removed);
    renumberBallots(level);
    renderBallotHistory(level);
    updateBallotCounter(level);
    updateStackSelect(level);
    updateResultsFromCache(level);
    debouncedSave();

    if (typeof supabaseSync !== 'undefined' && supabaseSync.isConnected()) {
        supabaseSync.deleteBallot(removed.id);
    }
    showToast(`Đã hủy phiếu #${removed.number}`, 'warning');
}

function deleteBallot(level, stackIdx, ballotIdx) {
    const removed = appData[level].stacks[stackIdx].ballots.splice(ballotIdx, 1)[0];
    invalidateBallotCache(level);
    removeBallotFromCache(level, removed);
    renumberBallots(level);
    renderBallotHistory(level);
    updateBallotCounter(level);
    updateStackSelect(level);
    updateResultsFromCache(level);
    debouncedSave();

    if (typeof supabaseSync !== 'undefined' && supabaseSync.isConnected()) {
        supabaseSync.deleteBallot(removed.id);
    }
    showToast('Đã xóa phiếu', 'warning');
}

function clearAllBallots(level) {
    if (!confirm('Bạn có chắc muốn xóa TẤT CẢ phiếu đã nhập?')) return;
    appData[level].stacks.forEach(s => s.ballots = []);
    invalidateBallotCache(level);
    initVoteCounts(level);
    appData[level]._historyPage = 1;
    renderBallotHistory(level);
    updateBallotCounter(level);
    updateStackSelect(level);
    updateResultsFromCache(level);
    debouncedSave();

    if (typeof supabaseSync !== 'undefined' && supabaseSync.isConnected()) {
        supabaseSync.clearBallots(level);
    }
    showToast('Đã xóa tất cả phiếu', 'warning');
}

function renumberBallots(level) {
    let count = 0;
    appData[level].stacks.forEach(stack => {
        stack.ballots.forEach(b => { count++; b.number = count; });
    });
}

function updateBallotCounter(level) {
    const all = getAllBallotsCached(level);
    document.getElementById(`current-ballot-${level}`).textContent = `Phiếu #${all.length + 1}`;
    document.getElementById(`ballot-count-${level}`).textContent = all.length;
}

// ============ BALLOT HISTORY (PAGINATED) ============
function renderBallotHistory(level) {
    const data = appData[level];
    const tbody = document.querySelector(`#history-table-${level} tbody`);
    tbody.innerHTML = '';

    const showAll = data.currentStack < 0;
    let ballots = [];

    if (showAll) {
        data.stacks.forEach((stack, sIdx) => {
            stack.ballots.forEach((b, bIdx) => {
                ballots.push({ ...b, stackName: stack.name, stackIdx: sIdx, ballotIdx: bIdx });
            });
        });
    } else {
        const stack = data.stacks[data.currentStack];
        if (stack) {
            stack.ballots.forEach((b, bIdx) => {
                ballots.push({ ...b, stackName: stack.name, stackIdx: data.currentStack, ballotIdx: bIdx });
            });
        }
    }

    // Reverse to show latest first
    ballots.reverse();

    const totalBallots = ballots.length;
    const totalPages = Math.max(1, Math.ceil(totalBallots / BALLOTS_PER_PAGE));
    const page = Math.min(data._historyPage || 1, totalPages);
    data._historyPage = page;

    const startIdx = (page - 1) * BALLOTS_PER_PAGE;
    const endIdx = Math.min(startIdx + BALLOTS_PER_PAGE, totalBallots);
    const pageBallots = ballots.slice(startIdx, endIdx);

    // Use DocumentFragment for batch DOM insert
    const fragment = document.createDocumentFragment();

    pageBallots.forEach(b => {
        const tr = document.createElement('tr');
        let statusText, statusClass;
        if (b.status === 'valid') { statusText = '✓ Hợp lệ'; statusClass = 'status-valid'; }
        else if (b.status === 'none') { statusText = '● Bầu đủ'; statusClass = 'status-none'; }
        else { statusText = '✗ Không hợp lệ'; statusClass = 'status-invalid'; }

        const crossedNames = b.crossedOut.length > 0
            ? b.crossedOut.map(n => `${n}. ${data.candidates[n - 1] || '?'}`).join(', ')
            : '(Không gạch ai)';

        tr.innerHTML = `
            <td>${b.number}</td>
            <td>${b.stackName}</td>
            <td>${crossedNames}</td>
            <td><span class="${statusClass}">${statusText}</span></td>
            <td><button class="btn-delete-ballot" onclick="deleteBallot('${level}', ${b.stackIdx}, ${b.ballotIdx})">🗑️</button></td>
        `;
        fragment.appendChild(tr);
    });
    tbody.appendChild(fragment);

    // Render pagination controls
    renderPagination(level, page, totalPages, totalBallots);
}

function renderPagination(level, currentPage, totalPages, totalBallots) {
    const container = document.getElementById(`pagination-${level}`);
    if (!container) return;

    if (totalPages <= 1) {
        container.innerHTML = `<span class="pagination-info">Hiển thị tất cả ${totalBallots} phiếu</span>`;
        return;
    }

    const startItem = (currentPage - 1) * BALLOTS_PER_PAGE + 1;
    const endItem = Math.min(currentPage * BALLOTS_PER_PAGE, totalBallots);

    let html = `<span class="pagination-info">Hiển thị ${startItem}–${endItem} / ${totalBallots} phiếu</span>`;
    html += '<div class="pagination-buttons">';

    if (currentPage > 1) {
        html += `<button class="btn btn-sm btn-outline" onclick="goToPage('${level}', ${currentPage - 1})">◀ Trước</button>`;
    }

    const maxVisible = 7;
    let startPage = Math.max(1, currentPage - 3);
    let endPage = Math.min(totalPages, startPage + maxVisible - 1);
    if (endPage - startPage < maxVisible - 1) startPage = Math.max(1, endPage - maxVisible + 1);

    if (startPage > 1) {
        html += `<button class="btn btn-sm btn-outline" onclick="goToPage('${level}', 1)">1</button>`;
        if (startPage > 2) html += '<span class="pagination-dots">...</span>';
    }

    for (let p = startPage; p <= endPage; p++) {
        html += p === currentPage
            ? `<button class="btn btn-sm btn-primary">${p}</button>`
            : `<button class="btn btn-sm btn-outline" onclick="goToPage('${level}', ${p})">${p}</button>`;
    }

    if (endPage < totalPages) {
        if (endPage < totalPages - 1) html += '<span class="pagination-dots">...</span>';
        html += `<button class="btn btn-sm btn-outline" onclick="goToPage('${level}', ${totalPages})">${totalPages}</button>`;
    }

    if (currentPage < totalPages) {
        html += `<button class="btn btn-sm btn-outline" onclick="goToPage('${level}', ${currentPage + 1})">Sau ▶</button>`;
    }

    html += '</div>';
    container.innerHTML = html;
}

function goToPage(level, page) {
    appData[level]._historyPage = page;
    renderBallotHistory(level);
}

// ============ RESULTS (FROM CACHE) ============
function updateResults(level) {
    rebuildVoteCounts(level);
    updateResultsFromCache(level);
}

function updateStackFilterOptions(level) {
    const data = appData[level];
    const select = document.getElementById(`stack-filter-${level}`);
    if (!select) return;
    const currentVal = select.value;
    select.innerHTML = '<option value="all">📦 Tất cả xấp</option>';
    data.stacks.forEach((stack, idx) => {
        const count = stack.ballots.length;
        const opt = document.createElement('option');
        opt.value = idx;
        opt.textContent = `${stack.name} (${count} phiếu)`;
        select.appendChild(opt);
    });
    // Restore selection if still valid
    if (currentVal !== 'all' && parseInt(currentVal) < data.stacks.length) {
        select.value = currentVal;
    } else {
        select.value = 'all';
    }
}

function filterResultsByStack(level) {
    renderResultsForFilter(level);
}

function calculateVotesForBallots(ballots, totalCandidates) {
    const vc = { total: 0, valid: 0, invalid: 0, none: 0, perCandidate: new Array(totalCandidates).fill(0) };
    ballots.forEach(b => {
        vc.total++;
        if (b.status === 'valid') vc.valid++;
        else if (b.status === 'invalid') vc.invalid++;
        else if (b.status === 'none') vc.none++;

        if (b.status === 'valid' || b.status === 'none') {
            for (let i = 0; i < totalCandidates; i++) {
                if (!b.crossedOut.includes(i + 1)) {
                    vc.perCandidate[i]++;
                }
            }
        }
    });
    return vc;
}

function renderResultsForFilter(level) {
    const data = appData[level];
    const select = document.getElementById(`stack-filter-${level}`);
    const filterVal = select ? select.value : 'all';

    let vc;
    let filterLabel = '';

    if (filterVal === 'all') {
        if (!data._voteCounts) rebuildVoteCounts(level);
        vc = data._voteCounts;
        filterLabel = '';
    } else {
        const stackIdx = parseInt(filterVal);
        const stack = data.stacks[stackIdx];
        if (!stack) return;
        vc = calculateVotesForBallots(stack.ballots, data.totalCandidates);
        filterLabel = ` — ${stack.name}`;
    }

    const countable = vc.valid + vc.none;

    document.getElementById(`stats-grid-${level}`).innerHTML = `
        <div class="stat-card total"><div class="stat-value">${vc.total}</div><div class="stat-label">Tổng số phiếu${filterLabel}</div></div>
        <div class="stat-card valid"><div class="stat-value">${countable}</div><div class="stat-label">Phiếu hợp lệ</div></div>
        <div class="stat-card invalid"><div class="stat-value">${vc.invalid}</div><div class="stat-label">Phiếu không hợp lệ</div></div>
        <div class="stat-card nomark"><div class="stat-value">${vc.none}</div><div class="stat-label">Bầu đủ (không gạch)</div></div>
    `;

    const candidateResults = data.candidates.map((name, idx) => ({
        stt: idx + 1, name,
        votes: vc.perCandidate[idx] || 0,
        percent: countable > 0 ? ((vc.perCandidate[idx] / countable) * 100).toFixed(1) : '0.0'
    }));

    const sorted = [...candidateResults].sort((a, b) => b.votes - a.votes);
    const maxVotes = Math.max(...vc.perCandidate, 1);

    const electedStt = new Set();
    if (filterVal === 'all') {
        sorted.slice(0, data.electCount).forEach(c => { if (c.votes > 0) electedStt.add(c.stt); });
    }

    const tbody = document.querySelector(`#results-table-${level} tbody`);
    tbody.innerHTML = '';
    const fragment = document.createDocumentFragment();
    candidateResults.forEach(c => {
        const isElected = electedStt.has(c.stt);
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td style="font-weight:600">${c.stt}</td>
            <td class="candidate-name-cell">${c.name}${isElected ? ' <span class="elected-badge">Trúng cử</span>' : ''}</td>
            <td class="vote-count">${c.votes}</td>
            <td class="vote-percent">${c.percent}%</td>
            <td><div class="bar-container"><div class="bar-fill ${isElected ? 'elected' : ''}" style="width: ${maxVotes > 0 ? (c.votes / maxVotes * 100) : 0}%"></div></div></td>
        `;
        fragment.appendChild(tr);
    });
    tbody.appendChild(fragment);
}

function updateResultsFromCache(level) {
    updateStackFilterOptions(level);
    renderResultsForFilter(level);
}

// ============ EXPORT CSV ============
function exportCSV(level) {
    const data = appData[level];
    const select = document.getElementById(`stack-filter-${level}`);
    const filterVal = select ? select.value : 'all';

    let vc, ballots, filterLabel;

    if (filterVal === 'all') {
        ballots = getAllBallotsCached(level);
        if (!data._voteCounts) rebuildVoteCounts(level);
        vc = data._voteCounts;
        filterLabel = '';
    } else {
        const stackIdx = parseInt(filterVal);
        const stack = data.stacks[stackIdx];
        if (!stack || stack.ballots.length === 0) {
            showToast('Xấp này chưa có phiếu nào!', 'warning');
            return;
        }
        ballots = stack.ballots;
        vc = calculateVotesForBallots(stack.ballots, data.totalCandidates);
        filterLabel = ` - ${stack.name}`;
    }

    if (ballots.length === 0) { showToast('Chưa có phiếu nào để xuất!', 'warning'); return; }

    const countable = vc.valid + vc.none;

    let csv = '\uFEFF';
    csv += `Kết quả kiểm phiếu ${data.label}${filterLabel} - Nhiệm kỳ 2026-2031\n\n`;
    csv += `Tổng số phiếu,${vc.total}\nPhiếu hợp lệ,${countable}\nPhiếu không hợp lệ,${vc.invalid}\n\n`;
    csv += 'STT,Họ và tên,Số phiếu bầu,Tỷ lệ\n';

    data.candidates.forEach((name, idx) => {
        const votes = vc.perCandidate[idx] || 0;
        const percent = countable > 0 ? ((votes / countable) * 100).toFixed(1) : '0.0';
        csv += `${idx + 1},"${name}",${votes},${percent}%\n`;
    });

    csv += '\n--- Chi tiết phiếu ---\nSTT Phiếu,Xấp,Người bị gạch,Trạng thái\n';
    ballots.forEach(b => {
        const stackName = data.stacks.find(s => s.id === b.stackId)?.name || '?';
        const crossed = b.crossedOut.length > 0 ? b.crossedOut.join(' ') : 'Không gạch';
        const status = b.status === 'valid' ? 'Hợp lệ' : b.status === 'none' ? 'Bầu đủ' : 'Không hợp lệ';
        csv += `${b.number},"${stackName}","${crossed}","${status}"\n`;
    });

    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `kiem-phieu-${level}-${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
    showToast('Đã xuất file CSV!', 'success');
}

// ============ EXPORT EXCEL ============
function exportExcel(level) {
    const data = appData[level];
    const allBallots = getAllBallotsCached(level);
    if (allBallots.length === 0) { showToast('Chưa có phiếu nào để xuất!', 'warning'); return; }

    if (typeof XLSX === 'undefined') {
        showToast('❌ Đang tải thư viện Excel, vui lòng thử lại sau 3 giây...', 'warning');
        return;
    }

    const vc = data._voteCounts;
    const countable = vc.valid + vc.none;
    const today = new Date().toLocaleDateString('vi-VN');

    // ====== SHEET 1: Kết quả tổng hợp ======
    const sheet1Data = [];

    // Header rows
    sheet1Data.push(['CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM']);
    sheet1Data.push(['Độc lập - Tự do - Hạnh phúc']);
    sheet1Data.push([]);
    sheet1Data.push([`BIÊN BẢN KIỂM PHIẾU BẦU CỬ ${data.label.toUpperCase()}`]);
    sheet1Data.push([`Nhiệm kỳ 2026 - 2031`]);
    sheet1Data.push([`Ngày lập: ${today}`]);
    sheet1Data.push([]);

    // Summary stats
    sheet1Data.push(['THỐNG KÊ TỔNG HỢP']);
    sheet1Data.push(['Tổng số phiếu phát ra', vc.total]);
    sheet1Data.push(['Tổng số phiếu thu về', vc.total]);
    sheet1Data.push(['Số phiếu hợp lệ', countable]);
    sheet1Data.push(['Số phiếu không hợp lệ', vc.invalid]);
    sheet1Data.push(['Số phiếu bầu đủ (không gạch)', vc.none]);
    sheet1Data.push([]);

    // Candidate results
    sheet1Data.push(['KẾT QUẢ KIỂM PHIẾU']);
    sheet1Data.push(['STT', 'Họ và tên ứng cử viên', 'Số phiếu bầu', 'Tỷ lệ (%)', 'Kết quả']);

    const candidateResults = data.candidates.map((name, idx) => ({
        stt: idx + 1, name,
        votes: vc.perCandidate[idx] || 0,
        percent: countable > 0 ? ((vc.perCandidate[idx] / countable) * 100).toFixed(1) : '0.0'
    }));
    const sorted = [...candidateResults].sort((a, b) => b.votes - a.votes);
    const electedStt = new Set();
    sorted.slice(0, data.electCount).forEach(c => { if (c.votes > 0) electedStt.add(c.stt); });

    candidateResults.forEach(c => {
        sheet1Data.push([
            c.stt,
            c.name,
            c.votes,
            parseFloat(c.percent),
            electedStt.has(c.stt) ? 'TRÚNG CỬ' : ''
        ]);
    });

    sheet1Data.push([]);
    sheet1Data.push(['', '', '', '', `Ngày ${new Date().getDate()} tháng ${new Date().getMonth() + 1} năm ${new Date().getFullYear()}`]);
    sheet1Data.push(['TỔ TRƯỞNG TỔ BẦU CỬ', '', '', '', 'THƯ KÝ']);

    const ws1 = XLSX.utils.aoa_to_sheet(sheet1Data);

    // Set column widths
    ws1['!cols'] = [
        { wch: 6 },   // STT
        { wch: 35 },  // Tên
        { wch: 16 },  // Số phiếu
        { wch: 12 },  // Tỷ lệ
        { wch: 18 }   // Kết quả
    ];

    // Merge cells for headers
    ws1['!merges'] = [
        { s: { r: 0, c: 0 }, e: { r: 0, c: 4 } },  // Cộng hòa...
        { s: { r: 1, c: 0 }, e: { r: 1, c: 4 } },  // Độc lập...
        { s: { r: 3, c: 0 }, e: { r: 3, c: 4 } },  // Biên bản...
        { s: { r: 4, c: 0 }, e: { r: 4, c: 4 } },  // Nhiệm kỳ...
        { s: { r: 5, c: 0 }, e: { r: 5, c: 4 } },  // Ngày lập...
        { s: { r: 7, c: 0 }, e: { r: 7, c: 4 } },  // Thống kê...
    ];

    // ====== SHEET 2: Chi tiết phiếu ======
    const sheet2Data = [];
    sheet2Data.push([`CHI TIẾT PHIẾU BẦU — ${data.label.toUpperCase()}`]);
    sheet2Data.push([]);
    sheet2Data.push(['STT Phiếu', 'Xấp', 'Người bị gạch (STT)', 'Tên người bị gạch', 'Trạng thái']);

    allBallots.forEach(b => {
        const stackName = data.stacks.find(s => s.id === b.stackId)?.name || '?';
        const crossedSTT = b.crossedOut.length > 0 ? b.crossedOut.join(', ') : 'Không gạch';
        const crossedNames = b.crossedOut.length > 0
            ? b.crossedOut.map(n => data.candidates[n - 1] || '?').join(', ')
            : 'Không gạch ai';
        const status = b.status === 'valid' ? 'Hợp lệ' : b.status === 'none' ? 'Bầu đủ' : 'Không hợp lệ';
        sheet2Data.push([b.number, stackName, crossedSTT, crossedNames, status]);
    });

    const ws2 = XLSX.utils.aoa_to_sheet(sheet2Data);
    ws2['!cols'] = [
        { wch: 10 },  // STT
        { wch: 10 },  // Xấp
        { wch: 22 },  // STT gạch
        { wch: 40 },  // Tên gạch
        { wch: 16 }   // Trạng thái
    ];
    ws2['!merges'] = [
        { s: { r: 0, c: 0 }, e: { r: 0, c: 4 } }
    ];

    // ====== Create workbook ======
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws1, 'Kết quả');
    XLSX.utils.book_append_sheet(wb, ws2, 'Chi tiết phiếu');

    const fileName = `bao-cao-kiem-phieu-${level}-${new Date().toISOString().slice(0, 10)}.xlsx`;
    XLSX.writeFile(wb, fileName);
    showToast(`✓ Đã xuất file Excel: ${fileName}`, 'success');
}

// ============ PRINT ============
function printResults(level) {
    document.querySelectorAll('.results-panel').forEach(p => p.style.display = 'none');
    document.getElementById(`results-panel-${level}`).style.display = 'block';
    window.print();
    ['quochoi', 'tinh', 'phuong'].forEach(l => {
        if (appData[l].started) document.getElementById(`results-panel-${l}`).style.display = 'block';
    });
}

// ============ TOAST ============
function showToast(message, type = 'info') {
    const container = document.getElementById('toast-container');
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.textContent = message;
    container.appendChild(toast);
    setTimeout(() => toast.remove(), 3000);
}

// ============ DEBOUNCED LOCAL STORAGE ============
function debouncedSave() {
    if (_saveTimeout) clearTimeout(_saveTimeout);
    _saveTimeout = setTimeout(() => saveToStorage(), 500);
}

function saveToStorage() {
    try {
        const saveData = {};
        ['quochoi', 'tinh', 'phuong'].forEach(level => {
            const d = appData[level];
            saveData[level] = {
                label: d.label, totalCandidates: d.totalCandidates,
                electCount: d.electCount, crossCount: d.crossCount,
                candidates: d.candidates, stacks: d.stacks,
                currentStack: d.currentStack, started: d.started
            };
        });
        const json = JSON.stringify(saveData);
        if (json.length > 4 * 1024 * 1024) {
            const sizeKB = (json.length / 1024).toFixed(1);
            console.warn(`localStorage data: ${sizeKB}KB — approaching 5MB limit!`);
            showToast(`⚠ Dữ liệu đã ${sizeKB}KB, gần giới hạn. Hãy xuất CSV backup!`, 'warning');
        }
        localStorage.setItem('quochoivote_data', json);
    } catch (e) {
        console.error('Cannot save to localStorage:', e);
        showToast('❌ Không thể lưu dữ liệu! Hãy xuất CSV backup ngay.', 'error');
    }
}

function loadFromStorage() {
    try {
        const saved = localStorage.getItem('quochoivote_data');
        if (saved) {
            const parsed = JSON.parse(saved);
            ['quochoi', 'tinh', 'phuong'].forEach(level => {
                if (parsed[level]) {
                    const fields = ['label', 'totalCandidates', 'electCount', 'crossCount',
                                    'stacks', 'currentStack', 'started'];
                    fields.forEach(f => {
                        if (parsed[level][f] !== undefined) appData[level][f] = parsed[level][f];
                    });
                    // Only load saved candidates if the defaults are generic
                    const defaultsAreGeneric = appData[level].candidates.length === 0
                        || appData[level].candidates.every((c, i) => c === `Ứng cử viên ${i + 1}`);
                    if (parsed[level].candidates && defaultsAreGeneric) {
                        appData[level].candidates = parsed[level].candidates;
                    }
                }
            });
        }
    } catch (e) {
        console.warn('Cannot load from localStorage:', e);
    }
}
