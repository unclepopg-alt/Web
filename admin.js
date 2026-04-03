document.addEventListener('DOMContentLoaded', async function () {
    const calendarEl = document.getElementById('calendar');
    const tableBody = document.getElementById('admin-table-body');
    const searchInput = document.getElementById('table-search');

    const db = new Dexie('TrainingRoomDB');
    db.version(1).stores({
        bookings: 'id, title, start, end, agency, user, room'
    });

    let bookings = [];
    let currentUser = JSON.parse(localStorage.getItem('current_user')) || null;

    let monthlyChart, agencyChart, timeSlotChart;
    let calendar;

    if (!currentUser || currentUser.role !== 'admin') {
        window.location.href = 'admin-login.html';
        return;
    }

    const displayNameEl = document.querySelector('.user-name');
    if (displayNameEl && currentUser.name) {
        displayNameEl.textContent = currentUser.name;
    }

    const rooms = {
        'comp-room': { name: 'ห้องอบรมคอมพิวเตอร์ กทส.กห.', color: '#6366f1' }
    };

    const syncNote = document.getElementById('data-sync-note');
    if (syncNote && window.BookingSync && typeof BookingSync.remoteDataUrl === 'function') {
        syncNote.textContent =
            'ข้อมูลเชื่อมกับหน้าหลัก re-main ชุดเดียวกัน · แหล่งซิงค์ไฟล์: ' + BookingSync.remoteDataUrl();
    }

    async function migrateLegacyLocalStorage() {
        const raw = localStorage.getItem('room_bookings');
        if (!raw) return;
        try {
            const arr = JSON.parse(raw);
            if (!Array.isArray(arr)) return;
            await db.transaction('rw', db.bookings, async () => {
                for (const b of arr) {
                    if (b && b.id != null) await db.bookings.put(b);
                }
            });
            localStorage.removeItem('room_bookings');
            BookingSync.trigger();
        } catch (e) {
            console.warn('Legacy migration skipped:', e);
        }
    }

    async function loadBookings() {
        bookings = await db.bookings.toArray();
    }

    await migrateLegacyLocalStorage();
    await loadBookings();

    async function refreshAll() {
        await loadBookings();
        calendar.removeAllEvents();
        calendar.addEventSource(bookings);
        renderAdminTable(searchInput.value);
        updateStats();
        updateCharts();
    }

    BookingSync.onSync(refreshAll);

    calendar = new FullCalendar.Calendar(calendarEl, {
        initialView: 'dayGridMonth',
        headerToolbar: {
            left: 'prev,next today',
            center: 'title',
            right: 'dayGridMonth,timeGridWeek,timeGridDay'
        },
        locale: 'th',
        events: bookings,
        selectable: true,
        editable: true,
        height: 'auto',
        select: function (info) {
            openBookingModal(info.startStr, info.endStr);
        },
        eventClick: function (info) {
            handleEditBooking(info.event.id);
        }
    });

    calendar.render();
    initCharts();
    updateStats();
    renderAdminTable();

    await BookingSync.pullRemote(db);
    await refreshAll();

    setInterval(async () => {
        if (await BookingSync.pullRemote(db)) await refreshAll();
    }, BookingSync.pollIntervalMs);

    document.addEventListener('visibilitychange', async () => {
        if (document.visibilityState === 'visible') {
            if (await BookingSync.pullRemote(db)) await refreshAll();
        }
    });

    document.querySelectorAll('.nav-link[data-target]').forEach(link => {
        link.addEventListener('click', function () {
            const targetId = this.dataset.target;

            document.querySelectorAll('.nav-link').forEach(l => l.classList.remove('active'));
            this.classList.add('active');

            document.querySelectorAll('.view-tab').forEach(tab => tab.classList.remove('active'));
            document.getElementById(targetId).classList.add('active');

            let title = 'แผงควบคุมผู้ดูแลระบบ';
            if (targetId === 'calendar-view') title = 'แผงควบคุมผู้ดูแลระบบ (Calendar)';
            if (targetId === 'manage-view') title = 'จัดการข้อมูลการจอง (Management)';
            if (targetId === 'stats-view') title = 'สรุปสถิติการใช้งาน (Statistics)';

            document.getElementById('view-title').textContent = title;

            if (targetId === 'calendar-view') {
                calendar.updateSize();
            }
        });
    });

    function renderAdminTable(filterText = '') {
        const filtered = bookings.filter(b =>
            b.title.toLowerCase().includes(filterText.toLowerCase()) ||
            b.extendedProps.user.toLowerCase().includes(filterText.toLowerCase()) ||
            b.extendedProps.agency.toLowerCase().includes(filterText.toLowerCase())
        );

        tableBody.innerHTML = filtered.length === 0
            ? '<tr><td colspan="5" style="text-align:center; padding: 2rem; color: var(--text-muted);">ไม่พบข้อมูลที่ค้นหา</td></tr>'
            : filtered.map((b, index) => `
            <tr>
                <td>${index + 1}</td>
                <td style="font-weight:600;">${b.title}</td>
                <td>
                    <div style="font-weight:500;">${b.extendedProps.user}</div>
                    <div style="font-size:0.75rem; color: var(--text-muted);">${b.extendedProps.agency}</div>
                </td>
                <td>
                    <div class="badge badge-room">${b.extendedProps.roomName}</div>
                    <div style="font-size:0.75rem; margin-top:4px;">
                        ${new Date(b.start).toLocaleString('th-TH', { dateStyle: 'short', timeStyle: 'short' })}
                    </div>
                </td>
                <td>
                    <div class="action-btns">
                        <button class="edit-btn" onclick="handleEditBooking('${b.id}')" title="แก้ไข">✏️</button>
                        <button class="del-btn" onclick="handleDeleteBooking('${b.id}')" title="ลบ">🗑️</button>
                    </div>
                </td>
            </tr>
        `).join('');
    }

    searchInput.addEventListener('input', (e) => {
        renderAdminTable(e.target.value);
    });

    window.handleEditBooking = function (id) {
        const booking = bookings.find(b => b.id === id);
        if (!booking) return;

        const startTime = new Date(booking.start).toTimeString().slice(0, 5);
        const endTime = new Date(booking.end).toTimeString().slice(0, 5);
        const dateStr = booking.start.split('T')[0];

        Swal.fire({
            title: 'แก้ไขข้อมูลการจอง',
            html: `
                <div style="text-align:left; margin-bottom:10px; font-weight:600; color:var(--primary-color);">แก้ไข ID: ${id}</div>
                <div style="display:flex; flex-direction:column; gap:12px; text-align:left; font-family:'Sarabun', sans-serif;">
                    <div>
                        <label style="font-size:0.85rem; color:#495057; font-weight:600;">หัวข้อการอบรม</label>
                        <input id="edit-title" class="swal2-input" placeholder="ระบุหัวข้อการอบรม" value="${booking.title}" style="margin:4px 0 0 0; width:100%; box-sizing:border-box;">
                    </div>
                    
                    <div style="display:flex; gap:10px;">
                        <div style="flex:1;">
                            <label style="font-size:0.85rem; color:#495057; font-weight:600;">หน่วยงาน</label>
                            <select id="edit-agency" class="swal2-input" style="margin:4px 0 0 0; width:100%; box-sizing:border-box;">
                                <option value="" disabled>เลือกหน่วยงาน</option>
                                <optgroup label="--- หน่วยขึ้นตรง ---">
                                    ${ [
                                        "สำนักปลัดกระทรวงกลาโหม", "สำนักพัฒนาระบบราชการกลาโหม", "สำนักงานเลขานุการ สป.",
                                        "สำนักนโยบายและแผนกลาโหม", "กรมแสนยุทธนา", "สำนักงานประมาณกลาโหม",
                                        "กรมพระธรรมนูญ", "กรมการเงินกลาโหม", "ศูนย์การอุตสาหกรรมป้องกันประเทศและพลังงานทหาร",
                                        "กรมเทคโนโลยีสารสนเทศและอวกาศกลาโหม", "กรมวิทยาศาสตร์และเทคโนโลยีกลาโหม",
                                        "กรมการสรรพกำลังกลาโหม", "สำนักงานสนับสนุน สป.", "สำนักงานตรวจสอบภายในกลาโหม",
                                        "องค์การสงเคราะห์ทหารผ่านศึก"
                                    ].map(opt => `<option value="${opt}" ${opt === booking.extendedProps.agency ? 'selected' : ''}>${opt}</option>`).join('') }
                                </optgroup>
                                <optgroup label="--- กองทัพ ---">
                                    ${ ["กองทัพบก", "กองทัพเรือ", "กองทัพอากาศ"]
                                        .map(opt => `<option value="${opt}" ${opt === booking.extendedProps.agency ? 'selected' : ''}>${opt}</option>`).join('') }
                                </optgroup>
                            </select>
                        </div>
                        <div style="flex:1;">
                            <label style="font-size:0.85rem; color:#495057; font-weight:600;">ยศ (ถ้ามี)</label>
                            <select id="edit-rank" class="swal2-input" style="margin:4px 0 0 0; width:100%; box-sizing:border-box;">
                                <option value="" ${!booking.extendedProps.rank ? 'selected' : ''}>ไม่ระบุ</option>
                                ${ [
                                    "พลเอก (General)", "พลเรือเอก (Admiral)", "พลอากาศเอก (Air Chief Marshal)",
                                    "พลโท (Lieutenant General)", "พลเรือโท (Vice Admiral)", "พลอากาศโท (Air Marshal)",
                                    "พลตรี (Major General)", "พลเรือตรี (Rear Admiral)", "พลอากาศตรี (Air Vice Marshal)",
                                    "พันเอก (Colonel)", "นาวาเอก (Captain)", "นาวาอากาศเอก (Group Captain)",
                                    "พันโท (Lieutenant Colonel)", "นาวาโท (Commander)", "นาวาอากาศโท (Wing Commander)",
                                    "พันตรี (Major)", "นาวาตรี (Lieutenant Commander)", "นาวาอากาศตรี (Squadron Leader)",
                                    "ร้อยเอก (Captain)", "เรือเอก (Lieutenant)", "เรืออากาศเอก (Flight Lieutenant)",
                                    "ร้อยโท (Lieutenant)", "เรือโท (Lieutenant Junior Grade)", "เรืออากาศโท (Flying Officer)",
                                    "ร้อยตรี (Sub Lieutenant)", "เรือตรี (Sub Lieutenant)", "เรืออากาศตรี (Pilot Officer)"
                                ].map(opt => `<option value="${opt}" ${opt === booking.extendedProps.rank ? 'selected' : ''}>${opt}</option>`).join('') }
                            </select>
                        </div>
                    </div>

                    <div>
                        <label style="font-size:0.85rem; color:#495057; font-weight:600;">ชื่อผู้จอง</label>
                        <input id="edit-user" class="swal2-input" placeholder="ระบุชื่อผู้จอง" value="${booking.extendedProps.user}" style="margin:4px 0 0 0; width:100%; box-sizing:border-box;">
                    </div>
                    
                    <div style="display:flex; gap:10px;">
                        <div style="flex:1;">
                            <label style="font-size:0.85rem; color:#495057; font-weight:600;">เวลาเริ่ม</label>
                            <input id="edit-start-time" type="time" class="swal2-input" value="${startTime}" style="margin:4px 0 0 0; width:100%; box-sizing:border-box;">
                        </div>
                        <div style="flex:1;">
                            <label style="font-size:0.85rem; color:#495057; font-weight:600;">เวลาสิ้นสุด</label>
                            <input id="edit-end-time" type="time" class="swal2-input" value="${endTime}" style="margin:4px 0 0 0; width:100%; box-sizing:border-box;">
                        </div>
                    </div>
                </div>
            `,
            showCancelButton: true,
            confirmButtonText: 'บันทึกการแก้ไข',
            cancelButtonText: 'ยกเลิก',
            preConfirm: () => {
                const title = document.getElementById('edit-title').value;
                const agency = document.getElementById('edit-agency').value;
                const rank = document.getElementById('edit-rank').value;
                const user = document.getElementById('edit-user').value;
                const startT = document.getElementById('edit-start-time').value;
                const endT = document.getElementById('edit-end-time').value;

                if (!title || !user || !agency) {
                    Swal.showValidationMessage('กรุณากรอกข้อมูลให้ครบถ้วน');
                    return false;
                }
                return { title, agency, rank, user, startT, endT };
            }
        }).then(async (result) => {
            if (result.isConfirmed) {
                const idx = bookings.findIndex(b => b.id === id);
                bookings[idx].title = result.value.title;
                bookings[idx].start = `${dateStr}T${result.value.startT}`;
                bookings[idx].end = `${dateStr}T${result.value.endT}`;
                bookings[idx].extendedProps.user = result.value.user;
                bookings[idx].extendedProps.agency = result.value.agency;
                bookings[idx].extendedProps.rank = result.value.rank || '';

                try {
                    await db.bookings.put(bookings[idx]);
                    await saveAndRefresh();
                    Swal.fire('สำเร็จ!', 'อัปเดตข้อมูลการจองเรียบร้อยแล้ว', 'success');
                } catch (err) {
                    console.error(err);
                    Swal.fire('บันทึกไม่สำเร็จ', err.message || String(err), 'error');
                }
            }
        });
    };

    window.handleDeleteBooking = function (id) {
        Swal.fire({
            title: 'ยืนยันการลบ?',
            text: 'ข้อมูลการจองนี้จะถูกลบออกจากระบบอย่างถาวร',
            icon: 'warning',
            showCancelButton: true,
            confirmButtonColor: '#ef4444',
            confirmButtonText: 'ใช่, ลบเลย',
            cancelButtonText: 'ยกเลิก'
        }).then(async (result) => {
            if (result.isConfirmed) {
                try {
                    await db.bookings.delete(id);
                    bookings = bookings.filter(b => b.id !== id);
                    await saveAndRefresh();
                    Swal.fire('ลบแล้ว!', 'ข้อมูลการจองถูกลบออกแล้ว', 'success');
                } catch (err) {
                    console.error(err);
                    Swal.fire('ลบไม่สำเร็จ', err.message || String(err), 'error');
                }
            }
        });
    };

    async function saveAndRefresh() {
        BookingSync.trigger();
        await refreshAll();
    }

    function openBookingModal(start, end) {
        const dateStr = start.split('T')[0];
        Swal.fire({
            title: 'เพิ่มการจอง (โหมด Admin)',
            html: `
                <div style="display:flex; flex-direction:column; gap:12px; text-align:left; font-family:'Sarabun', sans-serif;">
                    <div>
                        <label style="font-size:0.85rem; color:#495057; font-weight:600;">หัวข้อการอบรม</label>
                        <input id="swal-title" class="swal2-input" placeholder="ระบุหัวข้อการอบรม" style="margin:4px 0 0 0; width:100%; box-sizing:border-box;">
                    </div>
                    
                    <div style="display:flex; gap:10px;">
                        <div style="flex:1;">
                            <label style="font-size:0.85rem; color:#495057; font-weight:600;">หน่วยงาน</label>
                            <select id="swal-agency" class="swal2-input" style="margin:4px 0 0 0; width:100%; box-sizing:border-box;">
                                <option value="" disabled selected>เลือกหน่วยงาน</option>
                                <optgroup label="--- หน่วยขึ้นตรง ---">
                                    ${ [
                                        "สำนักปลัดกระทรวงกลาโหม", "สำนักพัฒนาระบบราชการกลาโหม", "สำนักงานเลขานุการ สป.",
                                        "สำนักนโยบายและแผนกลาโหม", "กรมแสนยุทธนา", "สำนักงานประมาณกลาโหม",
                                        "กรมพระธรรมนูญ", "กรมการเงินกลาโหม", "ศูนย์การอุตสาหกรรมป้องกันประเทศและพลังงานทหาร",
                                        "กรมเทคโนโลยีสารสนเทศและอวกาศกลาโหม", "กรมวิทยาศาสตร์และเทคโนโลยีกลาโหม",
                                        "กรมการสรรพกำลังกลาโหม", "สำนักงานสนับสนุน สป.", "สำนักงานตรวจสอบภายในกลาโหม",
                                        "องค์การสงเคราะห์ทหารผ่านศึก"
                                    ].map(opt => `<option value="${opt}">${opt}</option>`).join('') }
                                </optgroup>
                                <optgroup label="--- กองทัพ ---">
                                    ${ ["กองทัพบก", "กองทัพเรือ", "กองทัพอากาศ"]
                                        .map(opt => `<option value="${opt}">${opt}</option>`).join('') }
                                </optgroup>
                            </select>
                        </div>
                        <div style="flex:1;">
                            <label style="font-size:0.85rem; color:#495057; font-weight:600;">ยศ (ถ้ามี)</label>
                            <select id="swal-rank" class="swal2-input" style="margin:4px 0 0 0; width:100%; box-sizing:border-box;">
                                <option value="" selected>ไม่ระบุ</option>
                                ${ [
                                    "พลเอก (General)", "พลเรือเอก (Admiral)", "พลอากาศเอก (Air Chief Marshal)",
                                    "พลโท (Lieutenant General)", "พลเรือโท (Vice Admiral)", "พลอากาศโท (Air Marshal)",
                                    "พลตรี (Major General)", "พลเรือตรี (Rear Admiral)", "พลอากาศตรี (Air Vice Marshal)",
                                    "พันเอก (Colonel)", "นาวาเอก (Captain)", "นาวาอากาศเอก (Group Captain)",
                                    "พันโท (Lieutenant Colonel)", "นาวาโท (Commander)", "นาวาอากาศโท (Wing Commander)",
                                    "พันตรี (Major)", "นาวาตรี (Lieutenant Commander)", "นาวาอากาศตรี (Squadron Leader)",
                                    "ร้อยเอก (Captain)", "เรือเอก (Lieutenant)", "เรืออากาศเอก (Flight Lieutenant)",
                                    "ร้อยโท (Lieutenant)", "เรือโท (Lieutenant Junior Grade)", "เรืออากาศโท (Flying Officer)",
                                    "ร้อยตรี (Sub Lieutenant)", "เรือตรี (Sub Lieutenant)", "เรืออากาศตรี (Pilot Officer)"
                                ].map(opt => `<option value="${opt}">${opt}</option>`).join('') }
                            </select>
                        </div>
                    </div>

                    <div>
                        <label style="font-size:0.85rem; color:#495057; font-weight:600;">ชื่อผู้จอง</label>
                        <input id="swal-user" class="swal2-input" placeholder="ระบุชื่อผู้จอง" style="margin:4px 0 0 0; width:100%; box-sizing:border-box;">
                    </div>

                    <div style="display:flex; gap:10px;">
                        <div style="flex:1;">
                            <label style="font-size:0.85rem; color:#495057; font-weight:600;">เวลาเริ่ม</label>
                            <input id="swal-start-time" type="time" class="swal2-input" value="08:00" style="margin:4px 0 0 0; width:100%; box-sizing:border-box;">
                        </div>
                        <div style="flex:1;">
                            <label style="font-size:0.85rem; color:#495057; font-weight:600;">เวลาสิ้นสุด</label>
                            <input id="swal-end-time" type="time" class="swal2-input" value="16:00" style="margin:4px 0 0 0; width:100%; box-sizing:border-box;">
                        </div>
                    </div>
                </div>
            `,
            showCancelButton: true,
            confirmButtonText: 'บันทึก',
            cancelButtonText: 'ยกเลิก',
            preConfirm: () => {
                const title = document.getElementById('swal-title').value;
                const agency = document.getElementById('swal-agency').value;
                const rank = document.getElementById('swal-rank') ? document.getElementById('swal-rank').value : '';
                const user = document.getElementById('swal-user').value;
                const startTime = document.getElementById('swal-start-time').value;
                const endTime = document.getElementById('swal-end-time').value;

                if (!title || !user || !agency) {
                    Swal.showValidationMessage('กรุณากรอกข้อมูลให้ครบถ้วน');
                    return false;
                }
                return { title, user, agency, rank, startTime, endTime };
            }
        }).then(async (result) => {
            if (result.isConfirmed) {
                const newB = {
                    id: Date.now().toString(),
                    title: result.value.title,
                    start: `${dateStr}T${result.value.startTime}`,
                    end: `${dateStr}T${result.value.endTime}`,
                    backgroundColor: rooms['comp-room'].color,
                    extendedProps: {
                        user: result.value.user,
                        agency: result.value.agency,
                        rank: result.value.rank || '',
                        room: 'comp-room',
                        roomName: rooms['comp-room'].name,
                        owner: 'admin'
                    }
                };
                try {
                    await db.bookings.put(newB);
                    await saveAndRefresh();
                } catch (err) {
                    console.error(err);
                    Swal.fire('บันทึกไม่สำเร็จ', err.message || String(err), 'error');
                }
            }
        });
    }

    document.getElementById('add-booking-btn').addEventListener('click', () => {
        const now = new Date().toISOString().split('T')[0];
        openBookingModal(now, now);
    });

    function updateStats() {
        document.getElementById('total-bookings').textContent = bookings.length;
        const now = new Date();
        const monthCount = bookings.filter(b => {
            const d = new Date(b.start);
            return d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear();
        }).length;
        document.getElementById('month-bookings').textContent = monthCount;
    }

    function initCharts() {
        const ctxMonthly = document.getElementById('monthlyTrendChart').getContext('2d');
        const ctxAgency = document.getElementById('agencyChart').getContext('2d');
        const ctxTime = document.getElementById('timeSlotChart').getContext('2d');

        const commonOptions = {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'bottom', labels: { font: { family: 'Sarabun' } } }
            }
        };

        monthlyChart = new Chart(ctxMonthly, {
            type: 'line',
            data: { labels: [], datasets: [{ label: 'จำนวนการจอง', data: [], borderColor: '#6366f1', backgroundColor: 'rgba(99, 102, 241, 0.1)', fill: true, tension: 0.4 }] },
            options: commonOptions
        });

        agencyChart = new Chart(ctxAgency, {
            type: 'bar',
            data: { labels: [], datasets: [{ label: 'จำนวนการจอง', data: [], backgroundColor: '#b8860b' }] },
            options: { ...commonOptions, indexAxis: 'y' }
        });

        timeSlotChart = new Chart(ctxTime, {
            type: 'bar',
            data: { labels: [], datasets: [{ label: 'จำนวนการจอง', data: [], backgroundColor: '#003366' }] },
            options: commonOptions
        });

        updateCharts();
    }

    function updateCharts() {
        if (!monthlyChart) return;

        const monthlyData = {};
        const months = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
        const currentYear = new Date().getFullYear();

        months.forEach(m => monthlyData[m] = 0);
        bookings.forEach(b => {
            const d = new Date(b.start);
            if (d.getFullYear() === currentYear) {
                monthlyData[months[d.getMonth()]]++;
            }
        });

        monthlyChart.data.labels = months;
        monthlyChart.data.datasets[0].data = months.map(m => monthlyData[m]);
        monthlyChart.update();

        const agencyData = {};
        bookings.forEach(b => {
            const agency = b.extendedProps.agency || 'ไม่ระบุ';
            agencyData[agency] = (agencyData[agency] || 0) + 1;
        });

        const sortedAgencies = Object.entries(agencyData).sort((a, b) => b[1] - a[1]).slice(0, 10);
        agencyChart.data.labels = sortedAgencies.map(a => a[0]);
        agencyChart.data.datasets[0].data = sortedAgencies.map(a => a[1]);
        agencyChart.update();

        const timeData = Array(24).fill(0);
        bookings.forEach(b => {
            const hour = new Date(b.start).getHours();
            timeData[hour]++;
        });

        const timeLabels = Array.from({ length: 24 }, (_, i) => `${i.toString().padStart(2, '0')}:00`);
        timeSlotChart.data.labels = timeLabels.slice(7, 20);
        timeSlotChart.data.datasets[0].data = timeData.slice(7, 20);
        timeSlotChart.update();
    }

    document.getElementById('logout-btn').addEventListener('click', () => {
        localStorage.removeItem('current_user');
        window.location.href = 'admin-login.html';
    });

    document.getElementById('export-pdf-btn').addEventListener('click', () => {
        if (bookings.length === 0) {
            Swal.fire('ไม่มีข้อมูล', 'ไม่พบการจองเพื่อส่งออก', 'info');
            return;
        }

        const printWindow = window.open('', '_blank');
        const sorted = [...bookings].sort((a, b) => new Date(a.start) - new Date(b.start));
        const now = new Date();
        const monthCount = bookings.filter(b => {
            const d = new Date(b.start);
            return d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear();
        }).length;

        let html = `
            <html>
            <head>
                <title>รายงานสรุปการจองห้องอบรม (ส่วนงานผู้ดูแลระบบ)</title>
                <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@400;700&display=swap" rel="stylesheet">
                <style>
                    body { font-family: 'Sarabun', sans-serif; padding: 40px; color: #333; line-height: 1.6; }
                    h1 { color: #4338ca; font-size: 24px; border-bottom: 3px solid #6366f1; padding-bottom: 12px; margin-bottom: 8px; }
                    .info { margin-bottom: 25px; color: #666; font-size: 14px; }
                    .summary-box { display: flex; gap: 20px; margin: 25px 0; }
                    .summary-card { flex: 1; background: #f5f3ff; border: 1px solid #ddd6fe; border-radius: 12px; padding: 20px; text-align: center; }
                    .summary-card .label { font-size: 13px; color: #6b7280; text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 5px; }
                    .summary-card .value { font-size: 32px; font-weight: 800; color: #4338ca; }
                    table { width: 100%; border-collapse: collapse; margin-top: 20px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); }
                    th, td { border: 1px solid #e5e7eb; padding: 12px 15px; text-align: left; font-size: 14px; }
                    th { background-color: #4338ca; color: white; font-weight: 700; text-transform: uppercase; font-size: 12px; }
                    tr:nth-child(even) { background-color: #f9fafb; }
                    .badge { display: inline-block; padding: 2px 8px; border-radius: 4px; font-size: 11px; font-weight: 700; background: #eef2ff; color: #4338ca; }
                    @media print { .no-print { display: none; } }
                </style>
            </head>
            <body>
                <h1>แผนบันทึกและรายงานสรุปการใช้ห้องอบรมคอมพิวเตอร์ (Admin)</h1>
                <div class="info">ออกรายงาน ณ วันที่: ${now.toLocaleDateString('th-TH', { year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit' })} น.</div>

                <div class="summary-box">
                    <div class="summary-card">
                        <div class="label">รายการจองสะสมทั้งหมด</div>
                        <div class="value">${bookings.length}</div>
                    </div>
                    <div class="summary-card">
                        <div class="label">รายการจองประจำเดือนนี้</div>
                        <div class="value">${monthCount}</div>
                    </div>
                </div>

                <table>
                    <thead>
                        <tr>
                            <th style="width: 40px; text-align:center;">ลำดับ</th>
                            <th>หัวข้อการอบรม / กิจกรรม</th>
                            <th>หน่วยงาน (ผู้ดำเนินการ/ผู้จอง)</th>
                            <th>วัน - เวลา ที่ใช้งาน</th>
                            <th style="width: 100px;">สถานะห้อง</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${sorted.map((b, i) => `
                            <tr>
                                <td style="text-align:center;">${i + 1}</td>
                                <td><strong>${b.title}</strong></td>
                                <td>${b.extendedProps.agency} <br><small style="color:#666;">(${b.extendedProps.user})</small></td>
                                <td>
                                    ${new Date(b.start).toLocaleDateString('th-TH', { day: '2-digit', month: '2-digit', year: '2-digit' })} <br>
                                    <small>${new Date(b.start).toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit' })} - ${new Date(b.end).toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit' })} น.</small>
                                </td>
                                <td style="text-align:center;"><span class="badge">ใช้ห้องปกติ</span></td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
                <div style="margin-top: 40px; text-align: right; font-size: 12px; color: #999;">
                    * รายงานนี้ถูกสร้างโดยระบบอัตโนมัติ กทส.กห.
                </div>
                <script>
                    window.onload = function() {
                        window.print();
                        window.onafterprint = function() { window.close(); };
                    };
                </script>
            </body>
            </html>
        `;

        printWindow.document.write(html);
        printWindow.document.close();
    });
});
