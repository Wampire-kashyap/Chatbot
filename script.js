// ============================================================
// Internal Support Chatbot — Main Script
// Uses SharePointAPI module for authentication & logging
// ============================================================

let faqs = [];
let resources = [];
let departments = [];
let currentUser = null;       // { tmid, name, email }
let chatHistory = [];          // in-memory only
let lastLogItemId = null;      // tracks the last logged query for feedback update

// ============================================================
// Initialization
// ============================================================
document.addEventListener('DOMContentLoaded', async () => {
    await initData();
    setupEventListeners();
    checkSession();
    showEnvironmentBadge();
});

async function initData() {
    // Fetch FAQs from JSON (hosted alongside HTML on SharePoint)
    try {
        const fRes = await fetch('faq.json');
        faqs = await fRes.json();
    } catch (e) {
        console.error("Failed to load faq.json", e);
        faqs = [];
    }

    // Fetch resources from JSON
    try {
        const rRes = await fetch('resources.json');
        resources = await rRes.json();
    } catch (e) {
        console.error("Failed to load resources.json", e);
        resources = [];
    }

    // Fetch departments for dynamic escalation
    try {
        const dRes = await fetch('departments.json');
        departments = await dRes.json();
    } catch (e) {
        console.error("Failed to load departments.json", e);
        departments = [];
    }
}

// Show a small badge in header indicating which mode is active
function showEnvironmentBadge() {
    const mode = SharePointAPI.getMode();
    const statusEl = document.querySelector('.status');
    if (statusEl && mode === 'localhost') {
        statusEl.innerHTML = `<span class="status-dot" style="background: #f59e0b;"></span> Localhost Mode`;
    }
}

// ============================================================
// Session Management
// ============================================================
function checkSession() {
    const savedUser = sessionStorage.getItem('current_user');
    if (savedUser) {
        currentUser = JSON.parse(savedUser);
        document.getElementById('login-modal').style.display = 'none';
        document.getElementById('chat-container').style.display = 'flex';
        loadChatHistory();
    }
}

// ============================================================
// Event Listeners
// ============================================================
function setupEventListeners() {
    const loginBtn = document.getElementById('login-btn');
    const tmInput = document.getElementById('tm-id-input');
    const sendBtn = document.getElementById('send-btn');
    const msgInput = document.getElementById('message-input');

    loginBtn.addEventListener('click', handleLogin);
    tmInput.addEventListener('keypress', (e) => { if (e.key === 'Enter') handleLogin(); });

    sendBtn.addEventListener('click', () => handleSend());
    msgInput.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') handleSend();
    });

    msgInput.addEventListener('input', (e) => {
        showSuggestions(e.target.value);
    });

    document.addEventListener('click', (e) => {
        if (!e.target.closest('.chat-input-area') && !e.target.closest('.suggestions-container')) {
            document.getElementById('suggestions-container').style.display = 'none';
        }
    });
}

// ============================================================
// Login — TMID Validation via SharePoint API
// ============================================================
async function handleLogin() {
    const tmId = document.getElementById('tm-id-input').value.trim();
    const errorText = document.getElementById('login-error');
    const loginBtn = document.getElementById('login-btn');

    if (!tmId) {
        errorText.textContent = 'Please enter your TM ID.';
        errorText.style.display = 'block';
        return;
    }

    // Show loading state
    loginBtn.disabled = true;
    loginBtn.textContent = 'Validating...';
    errorText.style.display = 'none';

    try {
        // Validate TMID against EMPLOYEE_MASTER (or fallback to users.json)
        const user = await SharePointAPI.validateTMID(tmId);

        if (user) {
            currentUser = user;
            sessionStorage.setItem('current_user', JSON.stringify(user));
            document.getElementById('login-modal').style.display = 'none';
            document.getElementById('chat-container').style.display = 'flex';
            errorText.style.display = 'none';

            setTimeout(() => {
                addMessage(`Hi ${user.name}! 👋 How can I help you today?`, 'bot');
            }, 500);
        } else {
            errorText.textContent = 'Invalid TMID. Access denied.';
            errorText.style.display = 'block';
        }
    } catch (error) {
        console.error('Login validation error:', error);
        errorText.textContent = 'Validation failed. Please try again.';
        errorText.style.display = 'block';
    } finally {
        // Reset button
        loginBtn.disabled = false;
        loginBtn.textContent = 'Start Chat';
    }
}

// ============================================================
// Send Message
// ============================================================
window.handleSend = function(textOverride = null) {
    const msgInput = document.getElementById('message-input');
    const text = textOverride || msgInput.value.trim();
    if (!text) return;

    document.getElementById('suggestions-container').style.display = 'none';
    msgInput.value = '';

    addMessage(text, 'user');
    showTyping();

    setTimeout(() => {
        removeTyping();
        processQuestion(text);
    }, 800 + Math.random() * 500);
}

// ============================================================
// Message Display
// ============================================================
function addMessage(text, sender, htmlContent = null, save = true) {
    const container = document.getElementById('chat-messages');

    const wrapper = document.createElement('div');
    wrapper.className = `message-wrapper ${sender}`;

    const msgDiv = document.createElement('div');
    msgDiv.className = `message ${sender}`;

    if (htmlContent) {
        msgDiv.innerHTML = htmlContent;
    } else {
        msgDiv.textContent = text;
    }

    const time = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    const timeSpan = document.createElement('span');
    timeSpan.className = 'timestamp';
    timeSpan.textContent = time;

    wrapper.appendChild(msgDiv);
    wrapper.appendChild(timeSpan);

    container.appendChild(wrapper);
    container.scrollTop = container.scrollHeight;

    if (save) {
        chatHistory.push({ text: htmlContent || text, sender, time, isHtml: !!htmlContent });
    }
}

function loadChatHistory() {
    if (chatHistory.length === 0) {
        setTimeout(() => {
            addMessage(`Welcome back, ${currentUser.name}! How can I help you today?`, 'bot');
        }, 500);
        return;
    }
    chatHistory.forEach(m => {
        addMessageHistory(m);
    });
}

function addMessageHistory(m) {
    const container = document.getElementById('chat-messages');
    const wrapper = document.createElement('div');
    wrapper.className = `message-wrapper ${m.sender}`;
    const msgDiv = document.createElement('div');
    msgDiv.className = `message ${m.sender}`;
    if (m.isHtml) msgDiv.innerHTML = m.text;
    else msgDiv.textContent = m.text;

    const timeSpan = document.createElement('span');
    timeSpan.className = 'timestamp';
    timeSpan.textContent = m.time;

    wrapper.appendChild(msgDiv);
    wrapper.appendChild(timeSpan);
    container.appendChild(wrapper);
    container.scrollTop = container.scrollHeight;
}

// ============================================================
// Typing Indicator
// ============================================================
function showTyping() {
    const container = document.getElementById('chat-messages');
    const wrapper = document.createElement('div');
    wrapper.className = `message-wrapper bot typing-msg`;
    wrapper.id = 'typing-indicator';

    const msgDiv = document.createElement('div');
    msgDiv.className = `message bot typing-indicator`;
    msgDiv.innerHTML = `<span></span><span></span><span></span>`;

    wrapper.appendChild(msgDiv);
    container.appendChild(wrapper);
    container.scrollTop = container.scrollHeight;
}

function removeTyping() {
    const el = document.getElementById('typing-indicator');
    if (el) el.remove();
}

// ============================================================
// Question Processing & Matching
// ============================================================
function processQuestion(text) {
    let lowerText = text.toLowerCase();

    let bestMatch = null;
    let maxMatches = 0;
    let matchType = null;

    faqs.forEach(faq => {
        let matches = 0;
        faq.keywords.forEach(kw => {
            if (lowerText.includes(kw.toLowerCase())) matches++;
        });

        if (faq.question.toLowerCase().includes(lowerText) && lowerText.length > 5) {
            matches += 3;
        }

        if (matches > maxMatches) {
            maxMatches = matches;
            bestMatch = faq;
            matchType = 'faq';
        }
    });

    resources.forEach(res => {
        let matches = 0;
        res.keywords.forEach(kw => {
            if (lowerText.includes(kw.toLowerCase())) matches++;
        });

        if (res.title.toLowerCase().includes(lowerText) && lowerText.length > 5) {
            matches += 3;
        }

        if (matches > maxMatches || (matches === maxMatches && matches > 0 && matchType !== 'resource')) {
            maxMatches = matches;
            bestMatch = res;
            matchType = 'resource';
        }
    });

    if (bestMatch && maxMatches > 0) {
        const safeText = text.replace(/'/g, "\\'").replace(/"/g, '&quot;');
        let responseHtml = '';
        let answerForLog = '';

        if (matchType === 'faq') {
            answerForLog = bestMatch.answer;
            responseHtml = `
                <div>${bestMatch.answer}</div>
                <div class="feedback-container">
                    <div class="feedback-text">Was this helpful?</div>
                    <div class="feedback-btns">
                        <button class="feedback-btn" onclick="handleFeedback(true, '${safeText}', '${bestMatch.answer.replace(/'/g, "\\'")}', '${bestMatch.department}', this)">👍 Yes</button>
                        <button class="feedback-btn" onclick="handleFeedback(false, '${safeText}', '${bestMatch.answer.replace(/'/g, "\\'")}', '${bestMatch.department}', this)">👎 No</button>
                    </div>
                </div>
            `;
        } else {
            const iconClass = bestMatch.type === 'Document' ? 'fa-file-pdf' : 'fa-link';
            const actionText = bestMatch.type === 'Document' ? 'View Document' : 'Open Portal';
            answerForLog = `Resource: ${bestMatch.title} (${actionText})`;
            responseHtml = `
                <div class="resource-card">
                    <div class="resource-icon"><i class="fa-solid ${iconClass}"></i></div>
                    <div class="resource-info">
                        <h4>${bestMatch.title}</h4>
                        <a href="${bestMatch.url}" target="_blank" class="resource-link">${actionText}</a>
                    </div>
                </div>
                <div class="feedback-container">
                    <div class="feedback-text">Was this helpful?</div>
                    <div class="feedback-btns">
                        <button class="feedback-btn" onclick="handleFeedback(true, '${safeText}', '${answerForLog.replace(/'/g, "\\'")}', '${bestMatch.department}', this)">👍 Yes</button>
                        <button class="feedback-btn" onclick="handleFeedback(false, '${safeText}', '${answerForLog.replace(/'/g, "\\'")}', '${bestMatch.department}', this)">👎 No</button>
                    </div>
                </div>
            `;
        }

        addMessage('', 'bot', responseHtml);

        // Log to SharePoint USER_QUERIES_LOG
        SharePointAPI.logQuery(
            currentUser.tmid, currentUser.name, currentUser.email,
            text, answerForLog, 'Answered', ''
        ).then(result => {
            lastLogItemId = result.itemId;
        });

    } else {
        const safeText = text.replace(/'/g, "\\'").replace(/"/g, '&quot;');
        const responseHtml = `
            <div>I couldn't find an exact answer to your question. Please select a department to connect with human support:</div>
            ${getEscalationHtml(safeText)}
        `;
        addMessage('', 'bot', responseHtml);

        // Log escalation to SharePoint
        SharePointAPI.logQuery(
            currentUser.tmid, currentUser.name, currentUser.email,
            text, 'No match found — Escalated', 'Escalated', ''
        ).then(result => {
            lastLogItemId = result.itemId;
        });
    }
}

// ============================================================
// Feedback Handler
// ============================================================
window.handleFeedback = function(isHelpful, questionText, answerText, department, btnElement) {
    if (btnElement) {
        btnElement.closest('.feedback-btns').style.display = 'none';
        btnElement.closest('.feedback-container').querySelector('.feedback-text').innerText = isHelpful ? 'Thanks for your feedback!' : '';
    }

    const feedbackValue = isHelpful ? 'Helpful' : 'Not Helpful';

    // Update feedback on the existing log entry
    if (lastLogItemId) {
        SharePointAPI.updateFeedback(lastLogItemId, feedbackValue);
    }

    if (isHelpful) {
        const danceHtml = `
            <div class="celebration-bot">
                <img src="https://media1.giphy.com/media/v1.Y2lkPTc5MGI3NjExcGpmMWVlaHJ5YXkwcnZxeXozZHF4ZnIybTh5cjBxZHh3NTVqcmN1OCZlcD12MV9pbnRlcm5hbF9naWZfYnlfaWQmY3Q9Zw/kC8xMvHjix3RntS9bK/giphy.gif" alt="Dancing Robot" style="width: 120px; border-radius: 12px; margin-bottom: 0.5rem; display: block; margin-left: auto; margin-right: auto;" />
                <div style="text-align: center; font-weight: 500;">Glad I could help! 😊</div>
            </div>
        `;
        addMessage('', 'bot', danceHtml);

        setTimeout(() => {
            const containers = document.querySelectorAll('.celebration-bot');
            if (containers.length > 0) {
                const latestImg = containers[containers.length - 1].querySelector('img');
                if (latestImg) latestImg.style.display = 'none';
            }
        }, 4000);
    } else {
        const safeText = questionText.replace(/'/g, "\\'").replace(/"/g, '&quot;');
        const responseHtml = `
            <div>Sorry about that. Please select a department to connect with support:</div>
            ${getEscalationHtml(safeText)}
        `;
        addMessage('', 'bot', responseHtml);
    }
};

// ============================================================
// MS Teams Escalation Logic
// ============================================================
function getEscalationHtml(safeText) {
    const uniqueDepts = [...new Set(departments.map(d => d.department))].sort();
    let optionsHtml = '<option value="" disabled selected>Select Department</option>';
    
    if (uniqueDepts.length === 0) {
        optionsHtml += '<option value="" disabled>No departments configured</option>';
    } else {
        uniqueDepts.forEach(d => {
            optionsHtml += `<option value="${d}">${d}</option>`;
        });
    }
    
    return `
        <div class="escalate-form">
            <select class="dept-select" style="margin-bottom:0.5rem;" onchange="window.handleDeptSelect(this, '${safeText}')">
                ${optionsHtml}
            </select>
            <div class="agent-select-container" style="display:none; margin-bottom:0.5rem;"></div>
            <div class="teams-link-container"></div>
        </div>
    `;
}

window.handleDeptSelect = function(selectEl, safeText) {
    const dept = selectEl.value;
    const container = selectEl.parentElement;
    const agentContainer = container.querySelector('.agent-select-container');
    const linkContainer = container.querySelector('.teams-link-container');
    
    linkContainer.innerHTML = ''; // clear link
    
    const availableAgents = departments.filter(d => d.department === dept);
    if (availableAgents.length === 0) {
        agentContainer.innerHTML = '<div style="color:#ef4444; font-size:0.85rem; padding: 0.5rem;">No agents currently assigned.</div>';
        agentContainer.style.display = 'block';
        return;
    }
    
    let agentOptions = availableAgents.map(a => `<option value="${a.email}" data-name="${a.name}">${a.name} — ${a.shift || 'Available'}</option>`).join('');
    
    agentContainer.innerHTML = `
        <select class="dept-select" onchange="window.handleAgentSelect(this, '${safeText}', '${dept}')">
            <option value="" disabled selected>Select Support Agent</option>
            ${agentOptions}
        </select>
    `;
    agentContainer.style.display = 'block';
}

window.handleAgentSelect = function(selectEl, questionText, dept) {
    const email = selectEl.value;
    const agentName = selectEl.options[selectEl.selectedIndex].getAttribute('data-name');
    const container = selectEl.parentElement.parentElement.querySelector('.teams-link-container');
    
    const msg = `Hi ${agentName}, my name is ${currentUser.name} (ID: ${currentUser.tmid}). I need help with: "${questionText}"`;
    const encodedMsg = encodeURIComponent(msg);
    
    // MS Teams Deep Link
    const teamsLink = `https://teams.microsoft.com/l/chat/0/0?users=${email}&message=${encodedMsg}`;
    
    container.innerHTML = `<a href="${teamsLink}" target="_blank" class="whatsapp-btn" style="background-color: #5b5fc7; margin-top: 0.25rem;"><i class="fa-solid fa-users"></i> Chat on Teams with ${agentName}</a>`;
    
    // Log escalation routing to SharePoint
    SharePointAPI.logQuery(
        currentUser.tmid, currentUser.name, currentUser.email,
        questionText, `Escalated to ${agentName} (${dept}) via Teams`, 'Escalated', ''
    );
};

// ============================================================
// Search Suggestions
// ============================================================
window.showSuggestions = function(text) {
    const container = document.getElementById('suggestions-container');
    if (text.length < 2) {
        container.style.display = 'none';
        return;
    }

    const lower = text.toLowerCase();

    const faqMatches = faqs.filter(f => f.question.toLowerCase().includes(lower) || f.keywords.some(k => k.toLowerCase().includes(lower)))
        .map(f => ({ text: f.question, isRes: false, isDoc: false }));

    const resMatches = resources.filter(r => r.title.toLowerCase().includes(lower) || r.keywords.some(k => k.toLowerCase().includes(lower)))
        .map(r => ({ text: r.title, isRes: true, isDoc: r.type === 'Document' }));

    const combinedMatches = [...resMatches, ...faqMatches].slice(0, 4);

    if (combinedMatches.length > 0) {
        container.innerHTML = combinedMatches.map(m => `
            <div class="suggestion-item" onclick="handleSend('${m.text.replace(/'/g, "\\'").replace(/"/g, '&quot;')}')">
                <i class="fa-solid ${m.isRes ? (m.isDoc ? 'fa-file-pdf' : 'fa-link') : 'fa-magnifying-glass'}"></i> ${m.text}
            </div>
        `).join('');
        container.style.display = 'block';
    } else {
        container.style.display = 'none';
    }
}
