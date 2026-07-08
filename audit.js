// Shared audit trail helpers — load after Firebase SDK, config.js, and common.js
(function (global) {
    var AUDIT_ACTION_LABELS = {
        role_update: 'Kemas kini peranan',
        user_deleted: 'Padam pengguna',
        user_updated: 'Kemas kini pengguna',
        password_set_by_admin: 'Tetapkan kata laluan (admin)',
        divisions_updated: 'Kemas kini senarai bahagian',
        report_created: 'Hantar laporan baharu',
        report_updated: 'Kemas kini laporan',
        reports_cleared: 'Padam semua laporan',
        user_registered: 'Pendaftaran baharu',
        user_reregistered: 'Pendaftaran semula',
        profile_username_updated: 'Kemas kini username',
        password_changed: 'Tukar kata laluan sendiri',
        login_success: 'Log masuk'
    };

    function writeAuditLog(action, details) {
        details = details || {};
        var auth = typeof getAuth === 'function' ? getAuth() : null;
        var db = typeof getDb === 'function' ? getDb() : null;
        if (!auth || !auth.currentUser || !db || !action) {
            return Promise.resolve();
        }
        var by = (auth.currentUser.email || '').toLowerCase();
        var entry = {
            action: String(action),
            by: by,
            at: firebase.firestore.FieldValue.serverTimestamp()
        };
        if (details.targetEmail) entry.targetEmail = String(details.targetEmail).toLowerCase();
        if (details.division) entry.division = String(details.division);
        if (details.quarter) entry.quarter = String(details.quarter);
        if (details.reportId) entry.reportId = String(details.reportId);
        if (details.details) entry.details = String(details.details);
        if (details.newRole) entry.newRole = String(details.newRole);
        if (details.oldRole) entry.oldRole = String(details.oldRole);
        if (typeof details.count === 'number') entry.count = details.count;
        return db.collection('auditLog').add(entry).catch(function (err) {
            console.warn('Audit log failed:', err);
        });
    }

    function getAuditActionLabel(action) {
        return AUDIT_ACTION_LABELS[action] || action || '—';
    }

    global.AUDIT_ACTION_LABELS = AUDIT_ACTION_LABELS;
    global.writeAuditLog = writeAuditLog;
    global.getAuditActionLabel = getAuditActionLabel;
})(typeof window !== 'undefined' ? window : this);
