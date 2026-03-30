// ============================================================
// SharePoint API Module — Simplified for Static File Deployment
// ============================================================
// MODE: Uses JSON files (faq.json, users.json, resources.json)
// No SharePoint Lists required. Just upload files and go.
// ============================================================

const SharePointAPI = (function() {

    // ============================================================
    // TMID Validation — from users.json
    // ============================================================
    async function validateTMID(tmid) {
        try {
            const response = await fetch('users.json');
            const users = await response.json();
            const user = users.find(u => u.id.toLowerCase() === tmid.toLowerCase());
            if (user) {
                return {
                    tmid: user.id,
                    name: user.name,
                    email: user.email || `${user.id.toLowerCase()}@company.com`
                };
            }
            return null;
        } catch (e) {
            console.error('[Chatbot] Failed to validate TMID:', e);
            return null;
        }
    }

    // ============================================================
    // Query Logging — stored in sessionStorage for admin dashboard
    // ============================================================
    function logQuery(tmid, userName, userEmail, question, answer, status, feedback) {
        const logEntry = {
            Id: Date.now(),
            timestamp: new Date().toISOString(),
            Created: new Date().toISOString(),
            Title: question,
            TMID: tmid,
            UserName: userName,
            UserEmail: userEmail,
            AnswerProvided: answer,
            Status: status,
            Feedback: feedback || ''
        };

        // Store in sessionStorage so admin dashboard can display logs
        const logs = JSON.parse(sessionStorage.getItem('chatbot_logs') || '[]');
        logs.push(logEntry);
        sessionStorage.setItem('chatbot_logs', JSON.stringify(logs));

        console.log('[Chatbot] Interaction logged:', logEntry.Title, '→', logEntry.Status);
        return { success: true, itemId: logEntry.Id };
    }

    // ============================================================
    // Update Feedback on existing log
    // ============================================================
    function updateFeedback(itemId, feedback) {
        if (!itemId) return;
        const logs = JSON.parse(sessionStorage.getItem('chatbot_logs') || '[]');
        const entry = logs.find(l => l.Id === itemId);
        if (entry) {
            entry.Feedback = feedback;
            sessionStorage.setItem('chatbot_logs', JSON.stringify(logs));
            console.log('[Chatbot] Feedback updated:', feedback);
        }
    }

    // ============================================================
    // Fetch logs (for admin dashboard)
    // ============================================================
    function fetchLogs() {
        return JSON.parse(sessionStorage.getItem('chatbot_logs') || '[]');
    }

    // ============================================================
    // Mode indicator
    // ============================================================
    function getMode() {
        if (window.location.hostname.includes('sharepoint.com')) return 'sharepoint';
        return 'localhost';
    }

    return {
        validateTMID,
        logQuery,
        updateFeedback,
        fetchLogs,
        getMode
    };

})();
