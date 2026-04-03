/**
 * Cross-tab sync + periodic pull of data/bookings.json (GitHub Pages / static hosting).
 */
(function () {
    var SYNC_KEY = 'db_sync_trigger';
    var REMOTE_PATH = 'data/bookings.json';
    var lastRemoteSnap = null;

    function remoteUrl() {
        var base = window.BOOKING_DATA_BASE_URL;
        if (typeof base === 'string' && base.trim()) {
            var b = base.trim();
            if (!b.endsWith('/')) b += '/';
            return new URL(REMOTE_PATH, b).href;
        }
        var path = window.location.pathname || '/';
        var i = path.lastIndexOf('/');
        var dir = i >= 0 ? path.slice(0, i + 1) : '/';
        return new URL(REMOTE_PATH, window.location.origin + dir).href;
    }

    window.BookingSync = {
        /** ดึง data/bookings.json ถี่ขึ้น ≈ อัปเดตเร็วแบบไม่ใช้ฐานข้อมูลคลาวด์ */
        pollIntervalMs: 4000,

        trigger: function () {
            try {
                localStorage.setItem(SYNC_KEY, String(Date.now()));
            } catch (e) { /* ignore */ }
        },

        onSync: function (fn) {
            window.addEventListener('storage', function (e) {
                if (e.key === SYNC_KEY) fn();
            });
        },

        remoteDataUrl: remoteUrl,

        /**
         * Upsert JSON array from repo into Dexie. Triggers cross-tab sync if changed.
         * @param {Dexie} db
         * @returns {Promise<boolean>}
         */
        pullRemote: async function (db) {
            if (!db || !db.bookings) return false;
            var url = remoteUrl() + '?_=' + Date.now();
            var res;
            try {
                res = await fetch(url, { cache: 'no-store' });
            } catch (e) {
                return false;
            }
            if (!res.ok) return false;
            var text = await res.text();
            if (text === lastRemoteSnap) return false;
            lastRemoteSnap = text;
            var list;
            try {
                list = JSON.parse(text);
            } catch (e) {
                return false;
            }
            if (!Array.isArray(list)) return false;

            var changed = false;
            await db.transaction('rw', db.bookings, async function () {
                for (var i = 0; i < list.length; i++) {
                    var b = list[i];
                    if (!b || b.id == null) continue;
                    var prev = await db.bookings.get(String(b.id));
                    var normalized = Object.assign({}, b, { id: String(b.id) });
                    if (!prev || JSON.stringify(prev) !== JSON.stringify(normalized)) {
                        await db.bookings.put(normalized);
                        changed = true;
                    }
                }
            });
            if (changed) this.trigger();
            return changed;
        }
    };
})();
