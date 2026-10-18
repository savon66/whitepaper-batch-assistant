/**
 * 日志模块
 * 提供统一的日志记录和错误管理
 */
var Logger = (function() {
    var logs = [];
    var errors = [];
    
    function getTimeStr() {
        var d = new Date();
        var h = d.getHours().toString().padStart(2, '0');
        var m = d.getMinutes().toString().padStart(2, '0');
        var s = d.getSeconds().toString().padStart(2, '0');
        return h + ":" + m + ":" + s;
    }
    
    return {
        info: function(msg) {
            var time = getTimeStr();
            logs.push({ type: 'info', time: time, message: msg });
            try { console.log("[INFO] " + msg); } catch(e) {}
        },
        
        warn: function(msg) {
            var time = getTimeStr();
            logs.push({ type: 'warn', time: time, message: msg });
            try { console.warn("[WARN] " + msg); } catch(e) {}
        },
        
        error: function(msg, detail) {
            var time = getTimeStr();
            var errMsg = detail ? msg + " - " + detail : msg;
            errors.push({ time: time, message: errMsg });
            logs.push({ type: 'error', time: time, message: errMsg });
            try { console.error("[ERROR] " + errMsg); } catch(e) {}
        },
        
        getErrors: function() {
            return errors.slice();
        },
        
        getAll: function() {
            return logs.slice();
        },
        
        clear: function() {
            logs = [];
            errors = [];
        }
    };
})();
