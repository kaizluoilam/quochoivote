/* =====================================================
   Supabase Integration Module for Kiểm Phiếu Bầu Cử
   ===================================================== */

const supabaseSync = (() => {
    let _supabaseUrl = '';
    let _supabaseKey = '';
    let _connected = false;
    let _electionIds = {}; // { quochoi: 'uuid', tinh: 'uuid', phuong: 'uuid' }

    // ============ HELPERS ============
    async function supabaseFetch(path, options = {}) {
        const url = `${_supabaseUrl}/rest/v1/${path}`;
        const headers = {
            'apikey': _supabaseKey,
            'Authorization': `Bearer ${_supabaseKey}`,
            'Content-Type': 'application/json',
            'Prefer': options.prefer || 'return=representation',
            ...options.headers
        };
        const res = await fetch(url, { ...options, headers });
        if (!res.ok) {
            const text = await res.text();
            throw new Error(`Supabase error: ${res.status} — ${text}`);
        }
        const contentType = res.headers.get('content-type');
        if (contentType && contentType.includes('json')) {
            return res.json();
        }
        return null;
    }

    // ============ CONNECTION ============
    function isConnected() {
        return _connected;
    }

    async function connect(url, key) {
        _supabaseUrl = url.replace(/\/$/, '');
        _supabaseKey = key;

        // Test connection by fetching elections
        try {
            const data = await supabaseFetch('elections?select=id,level&limit=10');
            _connected = true;

            // Map existing elections
            if (data && data.length > 0) {
                data.forEach(e => {
                    _electionIds[e.level] = e.id;
                });
            }

            // Save credentials
            localStorage.setItem('supabase_url', _supabaseUrl);
            localStorage.setItem('supabase_key', _supabaseKey);

            return { success: true, elections: data };
        } catch (e) {
            _connected = false;
            throw e;
        }
    }

    function disconnect() {
        _connected = false;
        _electionIds = {};
        localStorage.removeItem('supabase_url');
        localStorage.removeItem('supabase_key');
    }

    function loadCredentials() {
        const url = localStorage.getItem('supabase_url');
        const key = localStorage.getItem('supabase_key');
        return { url, key };
    }

    // ============ SYNC ALL DATA ============
    async function syncAll() {
        if (!_connected) throw new Error('Not connected');

        const levels = ['quochoi', 'tinh', 'phuong'];

        for (const level of levels) {
            const data = appData[level];
            if (!data.started) continue;

            // 1. Create/update election
            let electionId = _electionIds[level];
            if (!electionId) {
                const elections = await supabaseFetch('elections', {
                    method: 'POST',
                    body: JSON.stringify({
                        name: `Bầu cử ${data.label} 2026`,
                        level: level,
                        total_candidates: data.totalCandidates,
                        elect_count: data.electCount,
                        cross_count: data.crossCount
                    })
                });
                electionId = elections[0].id;
                _electionIds[level] = electionId;
            }

            // 2. Sync candidates (upsert)
            const candidateRows = data.candidates.map((name, idx) => ({
                election_id: electionId,
                stt: idx + 1,
                name: name
            }));

            // Delete existing candidates and re-insert
            await supabaseFetch(`candidates?election_id=eq.${electionId}`, {
                method: 'DELETE',
                headers: { 'Prefer': 'return=minimal' }
            });
            if (candidateRows.length > 0) {
                await supabaseFetch('candidates', {
                    method: 'POST',
                    body: JSON.stringify(candidateRows)
                });
            }

            // 3. Sync stacks
            await supabaseFetch(`stacks?election_id=eq.${electionId}`, {
                method: 'DELETE',
                headers: { 'Prefer': 'return=minimal' }
            });

            const stackMap = {}; // old id -> new uuid
            for (let i = 0; i < data.stacks.length; i++) {
                const stack = data.stacks[i];
                const stackRows = await supabaseFetch('stacks', {
                    method: 'POST',
                    body: JSON.stringify({
                        election_id: electionId,
                        name: stack.name,
                        sort_order: i
                    })
                });
                stackMap[stack.id] = stackRows[0].id;
            }

            // 4. Sync ballots (batch insert in chunks of 500)
            await supabaseFetch(`ballots?election_id=eq.${electionId}`, {
                method: 'DELETE',
                headers: { 'Prefer': 'return=minimal' }
            });

            const allBallots = [];
            data.stacks.forEach(stack => {
                stack.ballots.forEach(b => {
                    allBallots.push({
                        election_id: electionId,
                        stack_id: stackMap[b.stackId],
                        ballot_number: b.number,
                        crossed_out: `{${b.crossedOut.join(',')}}`,
                        status: b.status
                    });
                });
            });

            // Batch insert in chunks of 500
            const BATCH_SIZE = 500;
            for (let i = 0; i < allBallots.length; i += BATCH_SIZE) {
                const chunk = allBallots.slice(i, i + BATCH_SIZE);
                await supabaseFetch('ballots', {
                    method: 'POST',
                    body: JSON.stringify(chunk),
                    headers: { 'Prefer': 'return=minimal' }
                });
            }
        }

        return { success: true };
    }

    // ============ INDIVIDUAL OPERATIONS ============
    async function insertBallot(level, ballot, stackIdx) {
        if (!_connected || !_electionIds[level]) return;

        try {
            const electionId = _electionIds[level];
            // We need the stack UUID — for simplicity, fetch from DB
            const stacks = await supabaseFetch(`stacks?election_id=eq.${electionId}&order=sort_order&limit=100`);
            const stackUuid = stacks[stackIdx]?.id;
            if (!stackUuid) return;

            await supabaseFetch('ballots', {
                method: 'POST',
                body: JSON.stringify({
                    election_id: electionId,
                    stack_id: stackUuid,
                    ballot_number: ballot.number,
                    crossed_out: `{${ballot.crossedOut.join(',')}}`,
                    status: ballot.status
                }),
                headers: { 'Prefer': 'return=minimal' }
            });
        } catch (e) {
            console.warn('Supabase insert ballot failed:', e);
        }
    }

    async function deleteBallot(ballotId) {
        // Note: in the current app, ballot IDs are timestamps, not UUIDs.
        // For now, individual deletes rely on full sync.
        // This is a placeholder for future improvement.
    }

    async function clearBallots(level) {
        if (!_connected || !_electionIds[level]) return;
        try {
            await supabaseFetch(`ballots?election_id=eq.${_electionIds[level]}`, {
                method: 'DELETE',
                headers: { 'Prefer': 'return=minimal' }
            });
        } catch (e) {
            console.warn('Supabase clear ballots failed:', e);
        }
    }

    // ============ PUBLIC API ============
    return {
        isConnected,
        connect,
        disconnect,
        loadCredentials,
        syncAll,
        insertBallot,
        deleteBallot,
        clearBallots
    };
})();

// ============ UI Functions for Supabase Panel ============
function toggleSupabasePanel() {
    const body = document.getElementById('supabase-body');
    body.style.display = body.style.display === 'none' ? 'block' : 'none';
}

async function connectSupabase() {
    const url = document.getElementById('supabase-url').value.trim();
    const key = document.getElementById('supabase-key').value.trim();

    if (!url || !key) {
        showToast('Vui lòng nhập Supabase URL và Anon Key!', 'warning');
        return;
    }

    try {
        showToast('Đang kết nối Supabase...', 'info');
        await supabaseSync.connect(url, key);

        const statusEl = document.getElementById('supabase-status');
        statusEl.textContent = '✓ Đã kết nối';
        statusEl.className = 'connection-status connected';
        document.getElementById('btn-sync').style.display = 'inline-flex';

        showToast('✓ Kết nối Supabase thành công!', 'success');
    } catch (e) {
        showToast(`❌ Kết nối thất bại: ${e.message}`, 'error');
    }
}

function disconnectSupabase() {
    supabaseSync.disconnect();

    const statusEl = document.getElementById('supabase-status');
    statusEl.textContent = 'Chưa kết nối';
    statusEl.className = 'connection-status';
    document.getElementById('btn-sync').style.display = 'none';

    showToast('Đã ngắt kết nối Supabase', 'info');
}

async function syncToSupabase() {
    try {
        showToast('🔄 Đang đồng bộ dữ liệu...', 'info');
        await supabaseSync.syncAll();
        showToast('✓ Đồng bộ Supabase thành công!', 'success');
    } catch (e) {
        showToast(`❌ Đồng bộ thất bại: ${e.message}`, 'error');
    }
}

// Auto-restore Supabase connection on page load
document.addEventListener('DOMContentLoaded', () => {
    const creds = supabaseSync.loadCredentials();
    if (creds.url && creds.key) {
        document.getElementById('supabase-url').value = creds.url;
        document.getElementById('supabase-key').value = creds.key;
        // Auto-reconnect
        supabaseSync.connect(creds.url, creds.key).then(() => {
            const statusEl = document.getElementById('supabase-status');
            statusEl.textContent = '✓ Đã kết nối';
            statusEl.className = 'connection-status connected';
            document.getElementById('btn-sync').style.display = 'inline-flex';
        }).catch(() => {
            // Silent fail on auto-reconnect
        });
    }
});
