document.addEventListener('DOMContentLoaded', async function () {
    const calendarEl = document.getElementById('calendar');
    if (!calendarEl) return;

    // Initialize Dexie Database
    const db = new Dexie('TrainingRoomDB');
    db.version(1).stores({
        bookings: 'id, title, start, end, agency, user, room'
    });

    let bookings = [];

    async function loadBookings() {
        bookings = await db.bookings.toArray();
    }
    await loadBookings();

    const rooms = {
        'comp-room': { name: 'ห้องอบรม กรมเทคโนโลยีสารสนเทศและอวกาศกลาโหม', color: '#003366' }
    };

    // Mobile Menu
    const mobileMenuBtn = document.getElementById('mobile-menu-btn');
    const sidebar = document.querySelector('.sidebar');
    const overlay = document.getElementById('sidebar-overlay');
    if (mobileMenuBtn && sidebar && overlay) {
        mobileMenuBtn.addEventListener('click', () => {
            sidebar.classList.toggle('active');
            overlay.classList.toggle('active');
        });
        overlay.addEventListener('click', () => {
            sidebar.classList.remove('active');
            overlay.classList.remove('active');
        });
    }

    let calendar;

    async function refreshFromDb() {
        await loadBookings();
        calendar.refetchEvents();
        updateStats();
        renderRecentList();
    }

    BookingSync.onSync(refreshFromDb);

    function triggerSync() {
        BookingSync.trigger();
    }

    // Calendar
    calendar = new FullCalendar.Calendar(calendarEl, {
        initialView: 'dayGridMonth',
        headerToolbar: {
            left: 'prev,next today',
            center: 'title',
            right: 'dayGridMonth,timeGridWeek'
        },
        locale: 'th',
        events: (info, success) => success(bookings),
        selectable: true,
        height: 'auto',
        select: (info) => openBookingModal(info.startStr, info.endStr),
        eventClick: (info) => {
            const b = info.event;
            Swal.fire({
                title: 'รายละเอียดการจอง',
                html: `<div style="text-align:left; font-family:'Sarabun', sans-serif;">
                    <p><strong>หัวข้อ:</strong> ${b.title}</p>
                    <p><strong>หน่วยงาน:</strong> ${b.extendedProps.agency}</p>
                    ${b.extendedProps.rank ? `<p><strong>ยศ:</strong> ${b.extendedProps.rank}</p>` : ''}
                    <p><strong>ผู้จอง:</strong> ${b.extendedProps.user}</p>
                    <p><strong>เริ่ม:</strong> ${b.start.toLocaleString('th-TH')}</p>
                </div>`,
                showCancelButton: true,
                confirmButtonText: 'ลบการจอง',
                confirmButtonColor: '#7a1f2d',
                cancelButtonText: 'ปิด'
            }).then(r => r.isConfirmed && deleteBooking(b.id));
        }
    });

    calendar.render();
    updateStats();
    renderRecentList();

    await BookingSync.pullRemote(db);
    await refreshFromDb();

    setInterval(async () => {
        if (await BookingSync.pullRemote(db)) await refreshFromDb();
    }, BookingSync.pollIntervalMs);

    document.addEventListener('visibilitychange', async () => {
        if (document.visibilityState === 'visible') {
            if (await BookingSync.pullRemote(db)) await refreshFromDb();
        }
    });

    // Modals
    function openBookingModal(start, end) {
        const initialDate = start.split('T')[0];
        const todayStr = new Date(new Date().getTime() - new Date().getTimezoneOffset() * 60000).toISOString().split('T')[0];
        Swal.fire({
            title: 'จองห้องอบรม',
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
                                    <option value="สำนักปลัดกระทรวงกลาโหม">สำนักปลัดกระทรวงกลาโหม</option>
                                    <option value="สำนักพัฒนาระบบราชการกลาโหม">สำนักพัฒนาระบบราชการกลาโหม</option>
                                    <option value="สำนักงานเลขานุการ สป.">สำนักงานเลขานุการ สป.</option>
                                    <option value="สำนักนโยบายและแผนกลาโหม">สำนักนโยบายและแผนกลาโหม</option>
                                    <option value="กรมแสนยุทธนา">กรมแสนยุทธนา</option>
                                    <option value="สำนักงานประมาณกลาโหม">สำนักงานประมาณกลาโหม</option>
                                    <option value="กรมพระธรรมนูญ">กรมพระธรรมนูญ</option>
                                    <option value="กรมการเงินกลาโหม">กรมการเงินกลาโหม</option>
                                    <option value="ศูนย์การอุตสาหกรรมป้องกันประเทศและพลังงานทหาร">ศูนย์การอุตสาหกรรมป้องกันประเทศและพลังงานทหาร</option>
                                    <option value="กรมเทคโนโลยีสารสนเทศและอวกาศกลาโหม">กรมเทคโนโลยีสารสนเทศและอวกาศกลาโหม</option>
                                    <option value="กรมวิทยาศาสตร์และเทคโนโลยีกลาโหม">กรมวิทยาศาสตร์และเทคโนโลยีกลาโหม</option>
                                    <option value="กรมการสรรพกำลังกลาโหม">กรมการสรรพกำลังกลาโหม</option>
                                    <option value="สำนักงานสนับสนุน สป.">สำนักงานสนับสนุน สป.</option>
                                    <option value="สำนักงานตรวจสอบภายในกลาโหม">สำนักงานตรวจสอบภายในกลาโหม</option>
                                    <option value="องค์การสงเคราะห์ทหารผ่านศึก">องค์การสงเคราะห์ทหารผ่านศึก</option>
                                </optgroup>
                                <optgroup label="--- กองทัพ ---">
                                    <option value="กองทัพบก">กองทัพบก</option>
                                    <option value="กองทัพเรือ">กองทัพเรือ</option>
                                    <option value="กองทัพอากาศ">กองทัพอากาศ</option>
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
                    
                    <div>
                        <label style="font-size:0.85rem; color:#495057; font-weight:600;">วันที่จอง</label>
                        <input id="swal-date" type="date" class="swal2-input" value="${initialDate}" style="margin:4px 0 0 0; width:100%; box-sizing:border-box;">
                    </div>

                    <div style="display:flex; gap:10px;">
                        <div style="flex:1;">
                            <label style="font-size:0.85rem; color:#495057; font-weight:600;">เวลาเริ่ม</label>
                            <input id="swal-start" type="time" class="swal2-input" value="08:00" style="margin:4px 0 0 0; width:100%; box-sizing:border-box;">
                        </div>
                        <div style="flex:1;">
                            <label style="font-size:0.85rem; color:#495057; font-weight:600;">เวลาสิ้นสุด</label>
                            <input id="swal-end" type="time" class="swal2-input" value="16:00" style="margin:4px 0 0 0; width:100%; box-sizing:border-box;">
                        </div>
                    </div>
                </div>
            `,
            showCancelButton: true,
            confirmButtonText: 'บันทึก',
            cancelButtonText: 'ยกเลิก',
            preConfirm: () => {
                const title = document.getElementById('swal-title').value;
                const user = document.getElementById('swal-user').value;
                const agency = document.getElementById('swal-agency').value;
                const rank = document.getElementById('swal-rank').value;
                const date = document.getElementById('swal-date').value;
                if (!title || !user || !agency || !date) return Swal.showValidationMessage('กรุณากรอกข้อมูลให้ครบ');

                return {
                    title,
                    user,
                    agency,
                    rank,
                    date,
                    start: document.getElementById('swal-start').value,
                    end: document.getElementById('swal-end').value
                };
            }
        }).then(async r => {
            if (r.isConfirmed) {
                const b = {
                    id: Date.now().toString(),
                    title: r.value.title,
                    start: `${r.value.date}T${r.value.start}`,
                    end: `${r.value.date}T${r.value.end}`,
                    backgroundColor: rooms['comp-room'].color,
                    extendedProps: { user: r.value.user, agency: r.value.agency, rank: r.value.rank, room: 'comp-room', roomName: rooms['comp-room'].name }
                };
                try {
                    await db.bookings.put(b);
                    await loadBookings();
                    calendar.refetchEvents();
                    updateStats();
                    renderRecentList();
                    triggerSync();
                    Swal.fire('สำเร็จ!', 'บันทึกการจองเรียบร้อยแล้ว', 'success');
                } catch (err) {
                    console.error(err);
                    Swal.fire('บันทึกไม่สำเร็จ', err.message || String(err), 'error');
                }
            }
        });
    }

    async function deleteBooking(id) {
        try {
            await db.bookings.delete(id);
            await loadBookings();
            calendar.refetchEvents();
            updateStats();
            renderRecentList();
            triggerSync();
            Swal.fire('ลบแล้ว!', 'ลบการจองเรียบร้อยแล้ว', 'success');
        } catch (err) {
            console.error(err);
            Swal.fire('ลบไม่สำเร็จ', err.message || String(err), 'error');
        }
    }

    function updateStats() {
        if (document.getElementById('total-bookings')) document.getElementById('total-bookings').textContent = bookings.length;
        if (document.getElementById('month-bookings')) {
            const now = new Date();
            document.getElementById('month-bookings').textContent = bookings.filter(b => {
                const d = new Date(b.start);
                return d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear();
            }).length;
        }
        if (document.getElementById('count-comp-room')) document.getElementById('count-comp-room').textContent = bookings.length;
    }

    function renderRecentList() {
        const el = document.getElementById('recent-bookings');
        if (!el) return;
        if (!bookings.length) return el.innerHTML = '<p style="text-align:center;color:#6c757d;font-size:0.8rem;margin-top:1rem;">ไม่มีการจองล่าสุด</p>';
        const sorted = [...bookings].sort((a, b) => b.id - a.id).slice(0, 5);
        el.innerHTML = sorted.map(b => `
            <div class="recent-item">
                <div class="recent-title">${b.title}</div>
                <div class="recent-meta">${b.extendedProps.agency} • ${new Date(b.start).toLocaleDateString('th-TH')}</div>
            </div>
        `).join('');
    }

    document.getElementById('add-booking-btn')?.addEventListener('click', () => openBookingModal(new Date().toISOString().split('T')[0]));

    // PDF EXPORT: Ultimate Balance with Background Watermark
    document.getElementById('export-pdf-btn')?.addEventListener('click', async () => {
        const btn = document.getElementById('export-pdf-btn');
        btn.disabled = true;
        btn.textContent = '⏱️ กำลังจัดลายน้ำสมดุล...';

        try {
            const now = new Date();
            const thaiYear = now.getFullYear() + 543;
            const thaiFullDate = `${now.getDate()} ${now.toLocaleDateString('th-TH', { month: 'long' })} พ.ศ. ${thaiYear}`;

            // Create report container for full A4 coverage with proportional padding
            const reportEl = document.createElement('div');
            reportEl.style.width = '210mm';
            reportEl.style.minHeight = '297mm';
            reportEl.style.padding = '25mm 20mm'; // Standard Official Margins
            reportEl.style.background = 'white';
            reportEl.style.color = '#000';
            reportEl.style.fontFamily = "'Sarabun', sans-serif";
            reportEl.style.boxSizing = 'border-box';
            reportEl.style.position = 'relative';
            reportEl.style.overflow = 'hidden';

            reportEl.innerHTML = `
                <!-- Background Watermark -->
                <div style="position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); opacity: 0.04; z-index: 0; pointer-events: none;">
                    <img src="1.png" style="width: 140mm; height: 140mm; object-fit: contain;">
                </div>

                <div style="position: relative; z-index: 1;">
                    <div style="text-align: center; margin-bottom: 40px;">
                        <img src="1.png" style="width: 22mm; height: 22mm; margin-bottom: 15px; display: block; margin-left: auto; margin-right: auto;">
                        <h1 style="font-size: 22pt; color: #003366; margin: 0 0 15px 0; font-weight: bold; line-height: 1.3;">รายงานสรุปการใช้ห้องอบรม<br>กรมเทคโนโลยีสารสนเทศและอวกาศกลาโหม</h1>
                        <div style="margin: 10px auto; border-bottom: 2px solid #c9a227; width: 60mm;"></div>
                    </div>
                    
                    <div style="display: flex; justify-content: space-between; margin-bottom: 25px; font-size: 13pt; border-bottom: 1px solid #eee; padding-bottom: 10px;">
                        <div style="font-weight: bold;">ประเภทรายงาน: <span style="font-weight: normal;">ข้อมูลแผนการใช้ห้องรายเดือน</span></div>
                        <div style="font-weight: bold;">วันที่พิมพ์: <span style="font-weight: normal;">${thaiFullDate}</span></div>
                    </div>

                    <table style="width: 100%; border-collapse: collapse; margin-bottom: 50px; font-size: 11pt; table-layout: fixed;">
                        <thead>
                            <tr style="background: #003366; color: white;">
                                <th style="border: 1px solid #003366; padding: 12px; text-align: center; width: 15mm;">ลำดับ</th>
                                <th style="border: 1px solid #003366; padding: 12px; text-align: left;">หัวข้อการอบรม / กิจกรรม</th>
                                <th style="border: 1px solid #003366; padding: 12px; text-align: left; width: 45mm;">หน่วยงาน</th>
                                <th style="border: 1px solid #003366; padding: 12px; text-align: left; width: 35mm;">ผู้รับผิดชอบ</th>
                                <th style="border: 1px solid #003366; padding: 12px; text-align: center; width: 25mm;">วันที่จอง</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${bookings.length ? bookings.sort((a, b) => new Date(a.start) - new Date(b.start)).map((b, i) => {
                const bd = new Date(b.start);
                const btd = `${bd.getDate()}/${bd.getMonth() + 1}/${bd.getFullYear() + 543}`;
                return `
                                <tr style="background: ${i % 2 === 0 ? 'rgba(255,255,255,0.8)' : 'rgba(252,252,252,0.8)'};">
                                    <td style="border: 1px solid #ccc; padding: 10px; text-align: center;">${i + 1}</td>
                                    <td style="border: 1px solid #ccc; padding: 10px; font-weight: bold;">${b.title}</td>
                                    <td style="border: 1px solid #ccc; padding: 10px;">${b.extendedProps.agency}</td>
                                    <td style="border: 1px solid #ccc; padding: 10px;">${b.extendedProps.user}</td>
                                    <td style="border: 1px solid #ccc; padding: 10px; text-align: center;">${btd}</td>
                                </tr>
                            `;
            }).join('') : '<tr><td colspan="5" style="border: 1px solid #ccc; text-align:center; padding:40px; color: #666;">--- ไม่พบข้อมูลการจองในระบบ ---</td></tr>'}
                        </tbody>
                    </table>
                    
                    <div style="margin-top: 80px; display: flex; justify-content: flex-end;">
                        <div style="text-align: center; width: 85mm; padding-right: 5mm;">
                            <p style="margin-bottom: 45px;">ลงชื่อ.......................................................................</p>
                            <p style="font-size: 13pt; font-weight: bold;">(.......................................................................)</p>
                            <p style="font-size: 11pt; margin-top: 10px; color: #333;">ตำแหน่ง...........................................................</p>
                            <p style="font-size: 11pt; margin-top: 5px; color: #333;">วันที่............/............/............</p>
                        </div>
                    </div>
                </div>

                </div>
            `;

            const opt = {
                margin: 0,
                filename: `Official_Balanced_Report_${thaiYear}.pdf`,
                image: { type: 'jpeg', quality: 1.0 },
                html2canvas: { scale: 2.5, useCORS: true, letterRendering: true },
                jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
            };

            await html2pdf().set(opt).from(reportEl).save();
            Swal.fire('สำเร็จ!', 'จัดทำเอกสาร PDF แบบสมดุลสูงสุดเรียบร้อยแล้ว', 'success');
        } catch (error) {
            console.error('PDF Error:', error);
            Swal.fire('เกิดข้อผิดพลาด', 'ไม่สามารถจัดสมดุลรายงานได้ กรุณาลองใหม่อีกครั้ง', 'error');
        } finally {
            btn.disabled = false;
            btn.textContent = '🖨️ พิมพ์รายงาน PDF';
        }
    });
});
