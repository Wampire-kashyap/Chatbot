// ============================================================
// Admin Dashboard — No localStorage, fetches from JSON files
// ============================================================

document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('admin-login-btn').addEventListener('click', handleAdminLogin);
    document.getElementById('admin-pwd').addEventListener('keypress', (e) => {
        if (e.key === 'Enter') handleAdminLogin();
    });
});

let faqsData = [];
let usersData = [];
let resourcesData = [];
let logsData = [];
let departmentsData = [];
let deptChartInstance = null;
let resolutionChartInstance = null;

// ============================================================
// Admin Login
// ============================================================
function handleAdminLogin() {
    const pwd = document.getElementById('admin-pwd').value;
    if (pwd === 'admin123') { // Simple demo auth
        document.getElementById('admin-login-modal').style.display = 'none';
        document.getElementById('admin-dashboard').style.display = 'flex';
        loadAdminData();
    } else {
        document.getElementById('admin-error').style.display = 'block';
    }
}

// ============================================================
// Load Data (from JSON files, no localStorage)
// ============================================================
async function loadAdminData() {
    // Fetch resources from JSON
    try {
        const rRes = await fetch('resources.json');
        resourcesData = await rRes.json();
    } catch (e) {
        console.error("Failed to load resources.json", e);
        resourcesData = [];
    }

    // Fetch departments from JSON
    try {
        const dRes = await fetch('departments.json');
        departmentsData = await dRes.json();
    } catch (e) {
        console.error("Failed to load departments.json", e);
        departmentsData = [];
    }

    // Fetch FAQs from JSON
    try {
        const fRes = await fetch('faq.json');
        faqsData = await fRes.json();
    } catch (e) {
        console.error("Failed to load faq.json", e);
        faqsData = [];
    }

    // Fetch users from JSON (fallback for localhost)
    try {
        const uRes = await fetch('users.json');
        usersData = await uRes.json();
    } catch (e) {
        console.error("Failed to load users.json", e);
        usersData = [];
    }

    // Fetch interaction logs from SharePoint (or session fallback)
    try {
        logsData = await SharePointAPI.fetchLogs();
    } catch (e) {
        console.error("Failed to fetch logs", e);
        logsData = [];
    }

    // Calculate live stats from fetched logs
    const total = logsData.length;
    const resolved = logsData.filter(l => l.Status === 'Answered').length;
    const escalated = logsData.filter(l => l.Status === 'Escalated').length;

    document.getElementById('stat-total').innerText = total;
    document.getElementById('stat-resolved').innerText = resolved;
    document.getElementById('stat-escalated').innerText = escalated;

    renderLogs();
    renderFAQs();
    renderUsers();
    renderResources();
    renderDepartments();
    renderCharts(resolved, escalated);

    // Show mode indicator
    const mode = SharePointAPI.getMode();
    if (mode === 'localhost') {
        const header = document.querySelector('.admin-header h2');
        if (header) header.innerHTML = '<i class="fa-solid fa-chart-pie"></i> Admin Dashboard <span style="font-size:0.7rem; background:#f59e0b; color:white; padding:0.2rem 0.5rem; border-radius:0.25rem; vertical-align:middle;">LOCALHOST</span>';
    }
}

// ============================================================
// Charts
// ============================================================
function renderCharts(resolved, escalated) {
    if (deptChartInstance) deptChartInstance.destroy();
    if (resolutionChartInstance) resolutionChartInstance.destroy();

    // Build department counts from actual logs
    const deptCounts = {};
    logsData.forEach(l => {
        // Extract department from AnswerProvided text if escalated
        if (l.AnswerProvided && l.AnswerProvided.includes('Escalated to ')) {
            const dept = l.AnswerProvided.replace('Escalated to ', '').replace(' via WhatsApp', '');
            deptCounts[dept] = (deptCounts[dept] || 0) + 1;
        } else if (l.Status === 'Answered') {
            deptCounts['Resolved'] = (deptCounts['Resolved'] || 0) + 1;
        }
    });

    const dCtx = document.getElementById('deptChart').getContext('2d');
    deptChartInstance = new Chart(dCtx, {
        type: 'bar',
        data: {
            labels: Object.keys(deptCounts).length ? Object.keys(deptCounts) : ['No Data Yet'],
            datasets: [{
                label: 'Queries',
                data: Object.keys(deptCounts).length ? Object.values(deptCounts) : [0],
                backgroundColor: '#4F46E5',
                borderRadius: 4
            }]
        },
        options: { responsive: true, plugins: { legend: { display: false } } }
    });

    const other = logsData.length - resolved - escalated;
    const rCtx = document.getElementById('resolutionChart').getContext('2d');
    resolutionChartInstance = new Chart(rCtx, {
        type: 'pie',
        data: {
            labels: ['Resolved', 'Escalated', 'Other'],
            datasets: [{
                data: [resolved || 0, escalated || 0, other > 0 ? other : 0],
                backgroundColor: ['#10b981', '#ef4444', '#9ca3af'],
                borderWidth: 0
            }]
        },
        options: { responsive: true, plugins: { legend: { position: 'bottom' } } }
    });
}

// ============================================================
// Tabs
// ============================================================
window.switchTab = function(tabId, event) {
    document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(tc => tc.classList.remove('active'));

    event.target.classList.add('active');
    document.getElementById(`${tabId}-tab`).classList.add('active');
};

// ============================================================
// Interaction Logs — Now in SharePoint Excel
// ============================================================
function renderLogs() {
    const tbody = document.querySelector('#logs-table tbody');

    if (logsData.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="7" style="text-align: center; padding: 2rem; color: #6b7280;">
                    <i class="fa-solid fa-cloud" style="font-size: 2rem; color: #4F46E5; display: block; margin-bottom: 0.75rem;"></i>
                    <strong>No interaction logs yet.</strong><br>
                    <span style="font-size: 0.85rem;">Logs will appear here as users interact with the chatbot.<br>Data is stored in the <strong>USER_QUERIES_LOG</strong> SharePoint List.</span>
                </td>
            </tr>
        `;
        return;
    }

    tbody.innerHTML = logsData.map(log => `
        <tr>
            <td>${log.Created ? new Date(log.Created).toLocaleString() : (log.timestamp ? new Date(log.timestamp).toLocaleString() : '—')}</td>
            <td>${log.TMID || '—'}</td>
            <td>${log.UserName || '—'}</td>
            <td>${log.Title || '—'}</td>
            <td>${log.AnswerProvided ? log.AnswerProvided.substring(0, 80) + (log.AnswerProvided.length > 80 ? '...' : '') : '—'}</td>
            <td>${log.Status === 'Escalated' ? '<span style="color:#ef4444;">Escalated</span>' : '<span style="color:#10b981;">Answered</span>'}</td>
            <td>${log.Feedback || '—'}</td>
        </tr>
    `).join('');
}

// ============================================================
// FAQ Management
// ============================================================
function renderFAQs() {
    const tbody = document.querySelector('#faqs-table tbody');
    tbody.innerHTML = faqsData.map(faq => `
        <tr>
            <td>${faq.id}</td>
            <td>${faq.question}</td>
            <td>${faq.keywords.join(', ')}</td>
            <td>${faq.department}</td>
            <td>
                <button class="action-btn" onclick="editFAQ(${faq.id})" title="Edit FAQ"><i class="fa-solid fa-pen"></i></button>
                <button class="action-btn delete" onclick="deleteFAQ(${faq.id})" title="Delete FAQ"><i class="fa-solid fa-trash"></i></button>
            </td>
        </tr>
    `).join('');
}

window.deleteFAQ = function(id) {
    if(confirm("Are you sure you want to delete this FAQ? Remember to export and update faq.json on SharePoint.")) {
        faqsData = faqsData.filter(f => f.id !== id);
        renderFAQs();
    }
}

window.editFAQ = function(id) {
    openFAQModal(id);
}

window.openFAQModal = function(id = null) {
    document.getElementById('faq-modal').style.display = 'flex';
    if (id) {
        document.getElementById('faq-modal-title').innerText = 'Edit FAQ';
        const faq = faqsData.find(f => f.id === id);
        document.getElementById('faq-id').value = faq.id;
        document.getElementById('faq-question').value = faq.question;
        document.getElementById('faq-answer').value = faq.answer;
        document.getElementById('faq-keywords').value = faq.keywords.join(', ');
        document.getElementById('faq-dept').value = faq.department;
    } else {
        document.getElementById('faq-modal-title').innerText = 'Add New FAQ';
        document.getElementById('faq-id').value = '';
        document.getElementById('faq-question').value = '';
        document.getElementById('faq-answer').value = '';
        document.getElementById('faq-keywords').value = '';
        document.getElementById('faq-dept').value = 'HR';
    }
}

window.closeFAQModal = function() {
    document.getElementById('faq-modal').style.display = 'none';
}

window.saveFAQ = function() {
    const id = document.getElementById('faq-id').value;
    const q = document.getElementById('faq-question').value.trim();
    const a = document.getElementById('faq-answer').value.trim();
    const kStr = document.getElementById('faq-keywords').value.trim();
    const d = document.getElementById('faq-dept').value;

    if (!q || !a) {
        alert("Question and Answer are required.");
        return;
    }

    const keywordsArray = kStr.split(',').map(k => k.trim()).filter(k => k);

    if (id) {
        const idx = faqsData.findIndex(f => f.id == id);
        if (idx > -1) {
            faqsData[idx] = { id: parseInt(id), question: q, answer: a, keywords: keywordsArray, department: d };
        }
    } else {
        const newId = faqsData.length > 0 ? Math.max(...faqsData.map(f => f.id)) + 1 : 1;
        faqsData.push({ id: newId, question: q, answer: a, keywords: keywordsArray, department: d });
    }

    closeFAQModal();
    renderFAQs();
    alert("FAQ saved in session. Use 'Export FAQs' to download and update faq.json on SharePoint.");
}

// ============================================================
// CSV Export Utilities
// ============================================================
function convertToCSV(objArray) {
    const array = typeof objArray != 'object' ? JSON.parse(JSON.stringify(objArray)) : objArray;
    let str = '';
    
    if(array.length > 0) {
        let headers = Object.keys(array[0]).join(',');
        str += headers + '\r\n';
    }

    for (let i = 0; i < array.length; i++) {
        let line = '';
        for (let index in array[i]) {
            if (line != '') line += ','
            let val = array[i][index];
            if(Array.isArray(val)) {
                val = val.join(';');
            }
            if (typeof val === 'string') {
                val = val.replace(/"/g, '""');
                val = '"' + val + '"';
            }
            line += val;
        }
        str += line + '\r\n';
    }
    return str;
}

function exportToCsv(filename, data) {
    if (data.length === 0) {
        alert("No data available to export.");
        return;
    }
    const csvStr = convertToCSV(data);
    const blob = new Blob([csvStr], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement("a");
    if (link.download !== undefined) { 
        const url = URL.createObjectURL(blob);
        link.setAttribute("href", url);
        link.setAttribute("download", filename);
        link.style.visibility = 'hidden';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }
}

window.exportLogs = function() {
    alert("Interaction logs are now stored in SharePoint Excel.\nNavigate to /ChatbotData/ in your SharePoint document library.");
}

window.exportFAQs = function() {
    exportToCsv("faqs.csv", faqsData);
}

window.handleBulkFAQUpload = function(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        const text = e.target.result;
        processFAQCSVText(text);
        event.target.value = '';
    };
    reader.readAsText(file);
}

function processFAQCSVText(text) {
    const lines = text.split(/\r?\n/);
    let addedCount = 0;
    
    const validLines = lines.filter(l => l.trim().length > 0);
    let startIndex = 0;
    if (validLines.length > 0) {
        const firstLineLower = validLines[0].toLowerCase();
        if (firstLineLower.includes('question') || firstLineLower.includes('keywords') || firstLineLower.includes('department')) {
            startIndex = 1;
        }
    }

    for (let i = startIndex; i < validLines.length; i++) {
        const line = validLines[i].trim();
        
        let cols = [];
        let inQuotes = false;
        let currentString = '';
        
        for (let char of line) {
            if (char === '"' && inQuotes) {
                inQuotes = false;
            } else if (char === '"' && !inQuotes) {
                inQuotes = true;
            } else if (char === ',' && !inQuotes) {
                cols.push(currentString);
                currentString = '';
            } else {
                currentString += char;
            }
        }
        cols.push(currentString);

        if (cols.length >= 2) {
            let q = cols[0] ? cols[0].trim().replace(/^"|"$/g, '') : '';
            let a = cols[1] ? cols[1].trim().replace(/^"|"$/g, '') : '';
            let kStr = cols[2] ? cols[2].trim().replace(/^"|"$/g, '') : '';
            let d = cols[3] ? cols[3].trim().replace(/^"|"$/g, '') : 'HR';
            
            if (q && a) {
                const keywordsArray = kStr.split(',').map(k => k.trim()).filter(k => k);
                const newId = faqsData.length > 0 ? Math.max(...faqsData.map(f => f.id)) + 1 : 1;
                faqsData.push({ id: newId, question: q, answer: a, keywords: keywordsArray, department: d });
                addedCount++;
            }
        }
    }
    
    renderFAQs();
    alert(`Bulk FAQ upload complete! Processed ${addedCount} lines.\nUse 'Export FAQs' to save and update faq.json on SharePoint.`);
}

// ============================================================
// User Management
// ============================================================
function renderUsers() {
    const tbody = document.querySelector('#users-table tbody');
    tbody.innerHTML = usersData.map(u => `
        <tr>
            <td>${u.id}</td>
            <td>${u.name}</td>
            <td>
                <button class="action-btn" onclick="editUser('${u.id}')" title="Edit User"><i class="fa-solid fa-pen"></i></button>
                <button class="action-btn delete" onclick="deleteUser('${u.id}')" title="Delete User"><i class="fa-solid fa-trash"></i></button>
            </td>
        </tr>
    `).join('');
}

window.deleteUser = function(id) {
    if(confirm("Are you sure you want to delete this user? Remember to export and update users.json on SharePoint.")) {
        usersData = usersData.filter(u => u.id !== id);
        renderUsers();
    }
}

window.editUser = function(id) {
    openUserModal(id);
}

window.openUserModal = function(id = null) {
    document.getElementById('user-modal').style.display = 'flex';
    if (id) {
        document.getElementById('user-modal-title').innerText = 'Edit User';
        const user = usersData.find(u => u.id === id);
        document.getElementById('user-original-id').value = user.id;
        document.getElementById('user-id').value = user.id;
        document.getElementById('user-name').value = user.name;
    } else {
        document.getElementById('user-modal-title').innerText = 'Add New User';
        document.getElementById('user-original-id').value = '';
        document.getElementById('user-id').value = '';
        document.getElementById('user-name').value = '';
    }
}

window.closeUserModal = function() {
    document.getElementById('user-modal').style.display = 'none';
}

window.saveUser = function() {
    const originalId = document.getElementById('user-original-id').value;
    const newId = document.getElementById('user-id').value.trim();
    const newName = document.getElementById('user-name').value.trim();

    if (!newId || !newName) {
        alert("TM ID and Name are required.");
        return;
    }

    if (originalId && originalId.toLowerCase() !== newId.toLowerCase()) {
        if (usersData.some(u => u.id.toLowerCase() === newId.toLowerCase())) {
            alert("A user with this TM ID already exists.");
            return;
        }
    } else if (!originalId) {
        if (usersData.some(u => u.id.toLowerCase() === newId.toLowerCase())) {
            alert("A user with this TM ID already exists.");
            return;
        }
    }

    if (originalId) {
        const idx = usersData.findIndex(u => u.id === originalId);
        if (idx > -1) {
            usersData[idx] = { id: newId, name: newName };
        }
    } else {
        usersData.push({ id: newId, name: newName });
    }

    closeUserModal();
    renderUsers();
    alert("User saved in session. Use 'Export CSV' to download and update users.json on SharePoint.");
}

window.exportUsers = function() {
    exportToCsv("users.csv", usersData);
}

window.handleBulkUpload = function(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        const text = e.target.result;
        processCSVText(text);
        event.target.value = '';
    };
    reader.readAsText(file);
}

function processCSVText(text) {
    const lines = text.split(/\r?\n/);
    let addedCount = 0;
    
    const validLines = lines.filter(l => l.trim().length > 0);
    let startIndex = 0;
    if (validLines.length > 0) {
        const firstLineLower = validLines[0].toLowerCase();
        if (firstLineLower.includes('tmid') || firstLineLower.includes('id') || firstLineLower.includes('name')) {
            startIndex = 1;
        }
    }

    for (let i = startIndex; i < validLines.length; i++) {
        const line = validLines[i];
        let cols = line.split(',');
        if (cols.length >= 2) {
            let id = cols[0].trim().replace(/^"|"$/g, '');
            let name = cols.slice(1).join(',').trim().replace(/^"|"$/g, '');
            
            if (id && name) {
                const existingIdx = usersData.findIndex(u => u.id.toLowerCase() === id.toLowerCase());
                if (existingIdx > -1) {
                    usersData[existingIdx].name = name;
                } else {
                    usersData.push({ id, name });
                    addedCount++;
                }
            }
        }
    }
    
    renderUsers();
    alert(`Bulk upload complete! Processed ${addedCount} lines.\nUse 'Export CSV' to save and update users.json on SharePoint.`);
}

// ============================================================
// Resource Management
// ============================================================
function renderResources() {
    const tbody = document.querySelector('#resources-table tbody');
    tbody.innerHTML = resourcesData.map(r => `
        <tr>
            <td>${r.id}</td>
            <td>${r.title}</td>
            <td><span style="background: ${r.type === 'Document' ? '#eef2ff' : '#f0fdf4'}; color: ${r.type === 'Document' ? '#4f46e5' : '#16a34a'}; padding: 0.2rem 0.5rem; border-radius: 0.25rem; font-size: 0.8rem; font-weight:600;"><i class="fa-solid ${r.type === 'Document' ? 'fa-file-pdf' : 'fa-link'}"></i> ${r.type}</span></td>
            <td><a href="${r.url}" target="_blank" style="color:var(--primary); text-decoration:none;">Link <i class="fa-solid fa-up-right-from-square" style="font-size:0.75rem;"></i></a></td>
            <td>${r.department}</td>
            <td>
                <button class="action-btn" onclick="editResource(${r.id})" title="Edit Resource"><i class="fa-solid fa-pen"></i></button>
                <button class="action-btn delete" onclick="deleteResource(${r.id})" title="Delete Resource"><i class="fa-solid fa-trash"></i></button>
            </td>
        </tr>
    `).join('');
}

window.deleteResource = function(id) {
    if(confirm("Are you sure you want to delete this resource? Remember to export and update resources.json on SharePoint.")) {
        resourcesData = resourcesData.filter(r => r.id !== id);
        renderResources();
    }
}

window.editResource = function(id) {
    openResourceModal(id);
}

window.openResourceModal = function(id = null) {
    document.getElementById('resource-modal').style.display = 'flex';
    if (id) {
        document.getElementById('resource-modal-title').innerText = 'Edit Resource';
        const resource = resourcesData.find(r => r.id === id);
        document.getElementById('resource-id').value = resource.id;
        document.getElementById('resource-title').value = resource.title;
        document.getElementById('resource-type').value = resource.type;
        document.getElementById('resource-url').value = resource.url || '';
        document.getElementById('resource-keywords').value = resource.keywords.join(', ');
        document.getElementById('resource-dept').value = resource.department;
        
        document.getElementById('resource-file').value = '';
        if (resource.type === 'Document') document.getElementById('document-preview').style.display = 'block';
        else document.getElementById('document-preview').style.display = 'none';
    } else {
        document.getElementById('resource-modal-title').innerText = 'Add New Resource';
        document.getElementById('resource-id').value = '';
        document.getElementById('resource-title').value = '';
        document.getElementById('resource-type').value = 'Document';
        document.getElementById('resource-url').value = '';
        document.getElementById('resource-file').value = '';
        document.getElementById('document-preview').style.display = 'none';
        document.getElementById('resource-keywords').value = '';
        document.getElementById('resource-dept').value = 'HR';
    }
    toggleResourceInput();
}

window.toggleResourceInput = function() {
    const type = document.getElementById('resource-type').value;
    const urlInput = document.getElementById('resource-url');
    const fileInput = document.getElementById('resource-file');
    const label = document.getElementById('resource-url-label');
    
    if (type === 'Document') {
        label.innerText = 'Upload Document';
        urlInput.style.display = 'none';
        fileInput.style.display = 'block';
    } else {
        label.innerText = 'Portal URL';
        urlInput.style.display = 'block';
        fileInput.style.display = 'none';
        document.getElementById('document-preview').style.display = 'none';
    }
}

window.closeResourceModal = function() {
    document.getElementById('resource-modal').style.display = 'none';
}

window.saveResource = function() {
    const id = document.getElementById('resource-id').value;
    const t = document.getElementById('resource-title').value.trim();
    const type = document.getElementById('resource-type').value;
    let url = document.getElementById('resource-url').value.trim();
    const kStr = document.getElementById('resource-keywords').value.trim();
    const d = document.getElementById('resource-dept').value;
    const fileInput = document.getElementById('resource-file');

    if (!t) {
        alert("Title is required."); return;
    }

    if (type === 'Document' && fileInput.files.length > 0) {
        const file = fileInput.files[0];
        if (file.size > 2 * 1024 * 1024) {
            alert("Warning: PDF is larger than 2MB. Consider hosting the file directly on SharePoint and linking to it.");
        }
        const reader = new FileReader();
        reader.onload = function(e) {
            finishSaveResource(id, t, type, e.target.result, kStr, d);
        };
        reader.readAsDataURL(file);
    } else {
        if (type === 'Document' && id) {
            const existing = resourcesData.find(r => r.id == id);
            url = existing ? existing.url : '';
        } else if (type === 'Document' && !id && !fileInput.files.length) {
            alert("Please upload a PDF document."); return;
        } else if (type === 'Link' && !url) {
            alert("Portal URL is required."); return;
        }
        finishSaveResource(id, t, type, url, kStr, d);
    }
}

function finishSaveResource(id, t, type, url, kStr, d) {
    const keywordsArray = kStr.split(',').map(k => k.trim()).filter(k => k);
    if (id) {
        const idx = resourcesData.findIndex(r => r.id == id);
        if (idx > -1) {
            resourcesData[idx] = { id: parseInt(id), title: t, type: type, url: url, keywords: keywordsArray, department: d };
        }
    } else {
        const newId = resourcesData.length > 0 ? Math.max(...resourcesData.map(r => r.id)) + 1 : 1;
        resourcesData.push({ id: newId, title: t, type: type, url: url, keywords: keywordsArray, department: d });
    }
    closeResourceModal();
    renderResources();
    alert("Resource saved in session. Use 'Export CSV' to download and update resources.json on SharePoint.");
}

window.exportResources = function() {
    exportToCsv("resources.csv", resourcesData);
}

window.handleBulkResourceUpload = function(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        const text = e.target.result;
        processResourceCSVText(text);
        event.target.value = '';
    };
    reader.readAsText(file);
}

function processResourceCSVText(text) {
    const lines = text.split(/\r?\n/);
    let addedCount = 0;
    
    const validLines = lines.filter(l => l.trim().length > 0);
    let startIndex = 0;
    if (validLines.length > 0) {
        const firstLineLower = validLines[0].toLowerCase();
        if (firstLineLower.includes('title') || firstLineLower.includes('url') || firstLineLower.includes('type')) {
            startIndex = 1;
        }
    }

    for (let i = startIndex; i < validLines.length; i++) {
        const line = validLines[i].trim();
        
        let cols = [];
        let inQuotes = false;
        let currentString = '';
        
        for (let char of line) {
            if (char === '"' && inQuotes) {
                inQuotes = false;
            } else if (char === '"' && !inQuotes) {
                inQuotes = true;
            } else if (char === ',' && !inQuotes) {
                cols.push(currentString);
                currentString = '';
            } else {
                currentString += char;
            }
        }
        cols.push(currentString);

        if (cols.length >= 3) {
            let t = cols[0] ? cols[0].trim().replace(/^"|"$/g, '') : '';
            let type = cols[1] ? cols[1].trim().replace(/^"|"$/g, '') : 'Document';
            if(type.toLowerCase().includes('link')) type = 'Link';
            else type = 'Document';
            let url = cols[2] ? cols[2].trim().replace(/^"|"$/g, '') : '';
            let kStr = cols[3] ? cols[3].trim().replace(/^"|"$/g, '') : '';
            let d = cols[4] ? cols[4].trim().replace(/^"|"$/g, '') : 'HR';
            
            if (t && url) {
                const keywordsArray = kStr.split(',').map(k => k.trim()).filter(k => k);
                const newId = resourcesData.length > 0 ? Math.max(...resourcesData.map(r => r.id)) + 1 : 1;
                resourcesData.push({ id: newId, title: t, type: type, url: url, keywords: keywordsArray, department: d });
                addedCount++;
            }
        }
    }
    
    renderResources();
    alert(`Bulk resource upload complete! Processed ${addedCount} lines.\nUse 'Export CSV' to save and update resources.json on SharePoint.`);
}

// ============================================================
// Departments & Agents Management
// ============================================================
function renderDepartments() {
    const tbody = document.querySelector('#departments-table tbody');
    if (!tbody) return;
    
    if (departmentsData.length === 0) {
        tbody.innerHTML = '<tr><td colspan="5" style="text-align:center; padding: 2rem; color: #6b7280;">No support agents defined. Add one or bulk upload.</td></tr>';
        return;
    }
    
    tbody.innerHTML = departmentsData.map(agent => `
        <tr>
            <td style="font-weight: 500;">${agent.department}</td>
            <td>${agent.name}</td>
            <td>${agent.email}</td>
            <td><span style="background: #e0e7ff; color: #4338ca; padding: 0.2rem 0.5rem; border-radius: 0.25rem; font-size: 0.85rem;">${agent.shift}</span></td>
            <td>
                <button class="action-btn edit-btn" onclick="editDeptAgent(${agent.id})"><i class="fa-solid fa-pen"></i></button>
                <button class="action-btn delete-btn" onclick="deleteDeptAgent(${agent.id})"><i class="fa-solid fa-trash"></i></button>
            </td>
        </tr>
    `).join('');
}

window.openDeptModal = function() {
    document.getElementById('edit-dept-id').value = '';
    document.getElementById('dept-name-input').value = '';
    document.getElementById('agent-name-input').value = '';
    document.getElementById('agent-email-input').value = '';
    document.getElementById('agent-shift-input').value = '';
    document.getElementById('dept-error-msg').style.display = 'none';
    document.getElementById('dept-modal-title').innerText = 'Add Support Agent';
    document.getElementById('dept-modal').style.display = 'flex';
}

window.closeDeptModal = function() {
    document.getElementById('dept-modal').style.display = 'none';
}

window.editDeptAgent = function(id) {
    const agent = departmentsData.find(d => d.id == id);
    if (!agent) return;
    
    document.getElementById('edit-dept-id').value = agent.id;
    document.getElementById('dept-name-input').value = agent.department;
    document.getElementById('agent-name-input').value = agent.name;
    document.getElementById('agent-email-input').value = agent.email;
    document.getElementById('agent-shift-input').value = agent.shift || '';
    
    document.getElementById('dept-error-msg').style.display = 'none';
    document.getElementById('dept-modal-title').innerText = 'Edit Support Agent';
    document.getElementById('dept-modal').style.display = 'flex';
}

window.saveDeptAgent = function() {
    const dId = document.getElementById('edit-dept-id').value;
    const dept = document.getElementById('dept-name-input').value.trim();
    const name = document.getElementById('agent-name-input').value.trim();
    const email = document.getElementById('agent-email-input').value.trim();
    const shift = document.getElementById('agent-shift-input').value.trim();
    
    if (!dept || !name || !email || !shift) {
        document.getElementById('dept-error-msg').style.display = 'block';
        return;
    }
    
    if (dId) {
        const idx = departmentsData.findIndex(d => d.id == dId);
        if (idx > -1) {
            departmentsData[idx] = { id: parseInt(dId), department: dept, name: name, email: email, shift: shift };
        }
    } else {
        const newId = Date.now();
        departmentsData.push({ id: newId, department: dept, name: name, email: email, shift: shift });
    }
    
    closeDeptModal();
    renderDepartments();
    alert("Agent saved in session. Use 'Export CSV' to download and update departments.json on SharePoint.");
}

window.deleteDeptAgent = function(id) {
    if (confirm("Are you sure you want to remove this agent?")) {
        departmentsData = departmentsData.filter(d => d.id != id);
        renderDepartments();
    }
}

window.exportDepartments = function() {
    const headers = "id,department,name,email,shift\n";
    const csvContent = departmentsData.map(d => `"${d.id}","${(d.department||'').replace(/"/g, '""')}","${(d.name||'').replace(/"/g, '""')}","${(d.email||'').replace(/"/g, '""')}","${(d.shift||'').replace(/"/g, '""')}"`).join("\n");
    const blob = new Blob([headers + csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement("a");
    const url = URL.createObjectURL(blob);
    link.setAttribute("href", url);
    link.setAttribute("download", "departments.csv");
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

window.handleBulkDeptUpload = function(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        const text = e.target.result;
        const lines = text.split(/\r?\n/);
        let addedCount = 0;
        
        const validLines = lines.filter(l => l.trim().length > 0);
        let startIndex = 0;
        if (validLines.length > 0 && validLines[0].toLowerCase().includes('department')) {
            startIndex = 1;
        }

        for (let i = startIndex; i < validLines.length; i++) {
            const line = validLines[i].trim();
            const cols = line.split(',').map(c => c.replace(/^"|"$/g, '').trim());
            
            // Expected cols from template: Department, TM Name, Email, Shift
            if (cols.length >= 4) {
                const dept = cols[0];
                const name = cols[1];
                const email = cols[2];
                const shift = cols[3];
                
                if (dept && name && email) {
                    departmentsData.push({ id: Date.now() + i, department: dept, name: name, email: email, shift: shift || '' });
                    addedCount++;
                }
            }
        }
        
        renderDepartments();
        alert(`Bulk upload complete! Processed ${addedCount} agents.\\nUse 'Export CSV' to save and update departments.json on SharePoint.`);
        event.target.value = '';
    };
    reader.readAsText(file);
}
