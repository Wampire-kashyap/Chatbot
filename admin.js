document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('admin-login-btn').addEventListener('click', handleAdminLogin);
    document.getElementById('admin-pwd').addEventListener('keypress', (e) => {
        if (e.key === 'Enter') handleAdminLogin();
    });
});

let logsData = [];
let faqsData = [];
let deptChartInstance = null;
let resolutionChartInstance = null;

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

function loadAdminData() {
    logsData = JSON.parse(localStorage.getItem('chat_logs') || '[]');
    faqsData = JSON.parse(localStorage.getItem('chat_faqs') || '[]');

    // Calculate stats
    document.getElementById('stat-total').innerText = logsData.length;
    let resolved = logsData.filter(l => l.matched && !l.escalated).length;
    let escalated = logsData.filter(l => l.escalated).length;

    document.getElementById('stat-resolved').innerText = resolved;
    document.getElementById('stat-escalated').innerText = escalated;

    renderLogs();
    renderFAQs();
    renderCharts(resolved, escalated);
}

function renderCharts(resolved, escalated) {
    if (deptChartInstance) deptChartInstance.destroy();
    if (resolutionChartInstance) resolutionChartInstance.destroy();

    const deptCounts = {};
    logsData.forEach(l => {
        if(l.department) {
            deptCounts[l.department] = (deptCounts[l.department] || 0) + 1;
        }
    });

    const dCtx = document.getElementById('deptChart').getContext('2d');
    deptChartInstance = new Chart(dCtx, {
        type: 'bar',
        data: {
            labels: Object.keys(deptCounts).length ? Object.keys(deptCounts) : ['No Data'],
            datasets: [{
                label: 'Queries by Department',
                data: Object.keys(deptCounts).length ? Object.values(deptCounts) : [0],
                backgroundColor: '#4F46E5',
                borderRadius: 4
            }]
        },
        options: { responsive: true, plugins: { legend: { display: false } } }
    });

    const rCtx = document.getElementById('resolutionChart').getContext('2d');
    resolutionChartInstance = new Chart(rCtx, {
        type: 'pie',
        data: {
            labels: ['Resolved', 'Escalated', 'Other'],
            datasets: [{
                data: [resolved, escalated, logsData.length - resolved - escalated],
                backgroundColor: ['#10b981', '#ef4444', '#9ca3af'],
                borderWidth: 0
            }]
        },
        options: { responsive: true, plugins: { legend: { position: 'bottom' } } }
    });
}

window.switchTab = function(tabId, event) {
    document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(tc => tc.classList.remove('active'));

    event.target.classList.add('active');
    document.getElementById(`${tabId}-tab`).classList.add('active');
};

function renderLogs() {
    const tbody = document.querySelector('#logs-table tbody');
    tbody.innerHTML = logsData.map(log => `
        <tr>
            <td>${new Date(log.date).toLocaleString()}</td>
            <td>${log.tmId}</td>
            <td>${log.name}</td>
            <td>${log.question}</td>
            <td>${log.matchedFaq || '-'}</td>
            <td>${log.escalated ? '<span style="color:#ef4444;">Yes</span>' : '<span style="color:#10b981;">No</span>'}</td>
            <td>${log.department || '-'}</td>
        </tr>
    `).join('');
}

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
    if(confirm("Are you sure you want to delete this FAQ?")) {
        faqsData = faqsData.filter(f => f.id !== id);
        localStorage.setItem('chat_faqs', JSON.stringify(faqsData));
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

    localStorage.setItem('chat_faqs', JSON.stringify(faqsData));
    closeFAQModal();
    renderFAQs();
}

function convertToCSV(objArray) {
    const array = typeof objArray != 'object' ? JSON.parse(JSON.stringify(objArray)) : objArray;
    let str = '';
    
    if(array.length > 0) {
        let headers = Object.keys(array[0]).join(',');
        str += headers + '\\r\\n';
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
        str += line + '\\r\\n';
    }
    return str;
}

window.exportLogs = function() {
    exportToCsv("interaction_logs.csv", logsData);
}

window.exportFAQs = function() {
    exportToCsv("faqs.csv", faqsData);
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
