document.addEventListener('DOMContentLoaded', async () => {
    await initData();
    setupEventListeners();
    checkSession();
});

let faqs = [];
let users = [];
let currentUser = null;

async function initData() {
    // Load default users if uninitialized
    if (!localStorage.getItem('chat_users')) {
        try {
            const uRes = await fetch('users.json');
            const uData = await uRes.json();
            localStorage.setItem('chat_users', JSON.stringify(uData));
        } catch (e) {
            console.error("Failed to load users", e);
        }
    }
    
    // Load default faqs if uninitialized
    if (!localStorage.getItem('chat_faqs')) {
        try {
            const fRes = await fetch('faq.json');
            const fData = await fRes.json();
            localStorage.setItem('chat_faqs', JSON.stringify(fData));
        } catch (e) {
            console.error("Failed to load faqs", e);
        }
    }

    if (!localStorage.getItem('chat_logs')) {
        localStorage.setItem('chat_logs', JSON.stringify([]));
    }

    users = JSON.parse(localStorage.getItem('chat_users') || "[]");
    faqs = JSON.parse(localStorage.getItem('chat_faqs') || "[]");
}

function checkSession() {
    const savedUser = sessionStorage.getItem('current_user');
    if (savedUser) {
        currentUser = JSON.parse(savedUser);
        document.getElementById('login-modal').style.display = 'none';
        document.getElementById('chat-container').style.display = 'flex';
        loadChatHistory();
    }
}

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

function handleLogin() {
    const tmId = document.getElementById('tm-id-input').value.trim();
    const errorText = document.getElementById('login-error');
    
    const user = users.find(u => u.id.toLowerCase() === tmId.toLowerCase());
    
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
        errorText.style.display = 'block';
    }
}

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
        const currentMsgs = JSON.parse(sessionStorage.getItem('chat_history') || '[]');
        currentMsgs.push({ text: htmlContent || text, sender, time, isHtml: !!htmlContent });
        sessionStorage.setItem('chat_history', JSON.stringify(currentMsgs));
    }
}

function loadChatHistory() {
    const msgs = JSON.parse(sessionStorage.getItem('chat_history') || '[]');
    if (msgs.length === 0) {
        setTimeout(() => {
            addMessage(`Welcome back, ${currentUser.name}! How can I help you today?`, 'bot');
        }, 500);
        return;
    }
    msgs.forEach(m => {
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

function processQuestion(text) {
    let lowerText = text.toLowerCase();
    
    let bestMatch = null;
    let maxMatches = 0;
    
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
        }
    });

    if (bestMatch && maxMatches > 0) {
        const safeText = text.replace(/'/g, "\\'").replace(/"/g, '&quot;');
        const responseHtml = `
            <div>${bestMatch.answer}</div>
            <div class="feedback-container">
                <div class="feedback-text">Was this helpful?</div>
                <div class="feedback-btns">
                    <button class="feedback-btn" onclick="handleFeedback(true, '${safeText}', ${bestMatch.id}, '${bestMatch.department}', this)">👍 Yes</button>
                    <button class="feedback-btn" onclick="handleFeedback(false, '${safeText}', ${bestMatch.id}, '${bestMatch.department}', this)">👎 No</button>
                </div>
            </div>
        `;
        addMessage('', 'bot', responseHtml);
        logInteraction(text, true, false, bestMatch.department, bestMatch.question);
    } else {
        const safeText = text.replace(/'/g, "\\'").replace(/"/g, '&quot;');
        const responseHtml = `
            <div>I couldn't find an exact answer to your question. Please select a department to connect with human support:</div>
            <div class="escalate-form">
                <select class="dept-select" onchange="generateWhatsappLink('${safeText}', this.value, this)">
                    <option value="" disabled selected>Select Department</option>
                    <option value="HR">HR</option>
                    <option value="IT">IT Support</option>
                    <option value="Finance">Finance</option>
                </select>
                <div class="wa-link-container" style="margin-top:0.75rem;"></div>
            </div>
        `;
        addMessage('', 'bot', responseHtml);
        logInteraction(text, false, true, null, null);
    }
}

window.handleFeedback = function(isHelpful, questionText, faqId, department, btnElement) {
    if (btnElement) {
        btnElement.closest('.feedback-btns').style.display = 'none';
        btnElement.closest('.feedback-container').querySelector('.feedback-text').innerText = isHelpful ? 'Thanks for your feedback!' : '';
    }

    if (isHelpful) {
        addMessage("Glad I could help! 😊", 'bot');
    } else {
        const safeText = questionText.replace(/'/g, "\\'").replace(/"/g, '&quot;');
        const responseHtml = `
            <div>Sorry about that. Please select a department to connect with support:</div>
            <div class="escalate-form">
                <select class="dept-select" onchange="generateWhatsappLink('${safeText}', this.value, this)">
                    <option value="" disabled selected>Select Department</option>
                    <option value="HR">HR</option>
                    <option value="IT">IT Support</option>
                    <option value="Finance">Finance</option>
                </select>
                <div style="margin-top:0.75rem;" class="wa-link-container"></div>
            </div>
        `;
        addMessage('', 'bot', responseHtml);
        
        const logs = JSON.parse(localStorage.getItem('chat_logs') || '[]');
        if (logs.length > 0) {
            logs[logs.length - 1].escalated = true;
            localStorage.setItem('chat_logs', JSON.stringify(logs));
        }
    }
};

window.generateWhatsappLink = function(questionText, dept, selectElement) {
    const msg = `Hi ${dept} team, my name is ${currentUser.name} (ID: ${currentUser.id}). I need help with: "${questionText}"`;
    const encodedMsg = encodeURIComponent(msg);
    const numbers = {
        'HR': '1234567890',
        'IT': '0987654321',
        'Finance': '1122334455'
    };
    const waLink = `https://wa.me/${numbers[dept] || '000000000'}?text=${encodedMsg}`;
    
    const linkHtml = `<a href="${waLink}" target="_blank" class="whatsapp-btn"><i class="fa-brands fa-whatsapp"></i> Chat on WhatsApp with ${dept}</a>`;
    
    const container = selectElement.parentElement.querySelector('.wa-link-container');
    if (container) {
        container.innerHTML = linkHtml;
    }

    const logs = JSON.parse(localStorage.getItem('chat_logs') || '[]');
    if (logs.length > 0) {
        logs[logs.length - 1].department = dept;
        localStorage.setItem('chat_logs', JSON.stringify(logs));
    }
};

function logInteraction(question, matched, escalated, department, matchedFaq) {
    const logs = JSON.parse(localStorage.getItem('chat_logs') || '[]');
    logs.push({
        date: new Date().toISOString(),
        tmId: currentUser.id,
        name: currentUser.name,
        question: question,
        matched: matched,
        escalated: escalated,
        department: department,
        matchedFaq: matchedFaq
    });
    localStorage.setItem('chat_logs', JSON.stringify(logs));
}

window.showSuggestions = function(text) {
    const container = document.getElementById('suggestions-container');
    if (text.length < 2) {
        container.style.display = 'none';
        return;
    }
    
    const lower = text.toLowerCase();
    const matches = faqs.filter(f => f.question.toLowerCase().includes(lower) || f.keywords.some(k => k.toLowerCase().includes(lower))).slice(0, 3);
    
    if (matches.length > 0) {
        container.innerHTML = matches.map(m => `
            <div class="suggestion-item" onclick="handleSend('${m.question.replace(/'/g, "\\'").replace(/"/g, '&quot;')}')">
                <i class="fa-solid fa-magnifying-glass"></i> ${m.question}
            </div>
        `).join('');
        container.style.display = 'block';
    } else {
        container.style.display = 'none';
    }
}
