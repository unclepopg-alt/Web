// Configuration
Chart.defaults.color = '#94a3b8';
Chart.defaults.font.family = "'Kanit', sans-serif";
Chart.register(ChartDataLabels);

let categoryChartInstance = null;
let trendChartInstance = null;
let globalData = [];

const excelFileInput = document.getElementById('excelFileInput');
const webFilter = document.getElementById('webFilter');
const categoryFilter = document.getElementById('categoryFilter');
const timeFilter = document.getElementById('timeFilter');
const syncStatus = document.getElementById('syncStatus');
const syncText = document.getElementById('syncText');

document.addEventListener('DOMContentLoaded', () => {
    excelFileInput.addEventListener('change', handleFileUpload);
    webFilter.addEventListener('change', applyFilters);
    categoryFilter.addEventListener('change', applyFilters);
    timeFilter.addEventListener('change', applyFilters);
});

function handleFileUpload(e) {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    syncText.textContent = "กำลังอ่านไฟล์...";
    syncStatus.classList.remove('connected');
    globalData = []; // Reset ข้อมูลเดิม

    let filesProcessed = 0;

    Array.from(files).forEach(file => {
        const reader = new FileReader();

        reader.onload = (event) => {
            try {
                const data = new Uint8Array(event.target.result);
                const workbook = XLSX.read(data, { type: 'array', cellDates: true });

                workbook.SheetNames.forEach(sheetName => {
                    const worksheet = workbook.Sheets[sheetName];
                    const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                    // กำหนดตัวแปรเว็บไซต์ (กห, สป, ทสอ, สม.กห, ศดจ) จากชื่อไฟล์
                    let websiteName = "ไม่ระบุ";
                    let checkName = (file.name + " " + sheetName).toLowerCase();
                    if (checkName.includes("ศูนย์ดิจิทัล")) websiteName = "ศดจ";
                    else if (checkName.includes("กรมเทคโนโลยีสารสนเทศ") || checkName.includes("dist")) websiteName = "ทสอ";
                    else if (checkName.includes("สมาคมภริยา") || checkName.includes("wives")) websiteName = "สม.กห";
                    else if (checkName.includes("สำนักงานปลัด") || checkName.includes("opsd")) websiteName = "สป";
                    else if (checkName.includes("กระทรวงกลาโหม") || checkName.includes("mod")) websiteName = "กห";

                    let headerRowIdx = -1;
                    // หาว่าหัวตารางอยู่บรรทัดไหน
                    for (let i = 0; i < Math.min(10, rawData.length); i++) {
                        if (rawData[i] && rawData[i].some(cell => typeof cell === 'string' && (cell.includes('ชื่อเรื่อง') || cell.includes('ลำดับ')))) {
                            headerRowIdx = i;
                            break;
                        }
                    }

                    if (headerRowIdx !== -1) {
                        const headers = rawData[headerRowIdx];
                        // ลูปอ่านข้อมูลทีละบรรทัด
                        for (let i = headerRowIdx + 1; i < rawData.length; i++) {
                            const row = rawData[i];
                            if (!row || row.length === 0) continue;

                            let rowObj = { _website: websiteName };
                            headers.forEach((h, colIdx) => {
                                if (h) rowObj[h.toString().trim()] = row[colIdx];
                            });

                            const titleKey = Object.keys(rowObj).find(k => k.includes('ชื่อเรื่อง'));
                            if (titleKey && rowObj[titleKey]) {
                                globalData.push(rowObj);
                            }
                        }
                    }
                });
            } catch (error) {
                console.error("Error reading file:", file.name, error);
            }

            filesProcessed++;
            if (filesProcessed === files.length) {
                processDataReady();
            }
        };

        reader.readAsArrayBuffer(file);
    });
}

function processDataReady() {
    if (globalData.length > 0) {
        populateFilters(globalData);
        applyFilters();
        syncStatus.classList.add('connected');
        syncText.textContent = `เชื่อมต่อสำเร็จ (${globalData.length} รายการ)`;
    } else {
        alert('ไม่พบข้อมูลในไฟล์ โปรดตรวจสอบไฟล์อีกครั้ง');
        syncText.textContent = "อ่านไฟล์ผิดพลาด";
    }
}

function populateFilters(data) {
    const catSet = new Set();

    data.forEach(row => {
        // ดึงประเภทข่าว
        const catKey = Object.keys(row).find(k => k.includes('ประเภท') || k.includes('งานที่ลง'));
        const category = (catKey && row[catKey]) ? row[catKey].toString().trim() : 'ไม่ระบุประเภท';
        catSet.add(category);
    });

    // อัปเดต Dropdown ประเภทข่าว
    categoryFilter.innerHTML = '<option value="all">ทุกประเภทข่าว / งานที่ลง</option>';
    Array.from(catSet).sort().forEach(cat => {
        categoryFilter.innerHTML += `<option value="${cat}">${cat}</option>`;
    });

    // รีเซ็ตค่าตัวกรองเป็น Default
    webFilter.value = 'all';
    timeFilter.value = 'all';
}

function applyFilters() {
    const selectedWeb = webFilter.value;
    const selectedCat = categoryFilter.value;
    const selectedTime = timeFilter.value;

    // หาวันที่ "ล่าสุด" ที่มีในข้อมูล เพื่อใช้เป็นจุดอ้างอิงของ "วันนี้" (หลีกเลี่ยงปัญหาข้อมูลคนละปีกับปีปัจจุบัน)
    let maxDateVal = 0;
    globalData.forEach(row => {
        const dateInKey = Object.keys(row).find(k => k.includes('เข้า') || k.includes('รับ'));
        const dateDoneKey = Object.keys(row).find(k => k.includes('แล้วเสร็จ'));
        let dateToUse = dateInKey ? row[dateInKey] : (dateDoneKey ? row[dateDoneKey] : null);

        if (dateToUse) {
            let d = (dateToUse instanceof Date) ? dateToUse : new Date(dateToUse);
            if (!isNaN(d.getTime()) && d.getTime() > maxDateVal) {
                maxDateVal = d.getTime();
            }
        }
    });
    // ตั้งค่า maxDate เป็นวันที่ใหม่ที่สุดในระบบ
    let maxDate = maxDateVal > 0 ? new Date(maxDateVal) : new Date();

    let filteredData = globalData.filter(row => {
        // 1. กรองหน่วยงาน (กห สป ทสอ สม.กห ศดจ)
        if (selectedWeb !== 'all' && row._website !== selectedWeb) return false;

        // 2. กรองประเภทข่าว
        const catKey = Object.keys(row).find(k => k.includes('ประเภท') || k.includes('งานที่ลง'));
        const category = (catKey && row[catKey]) ? row[catKey].toString().trim() : 'ไม่ระบุประเภท';
        if (selectedCat !== 'all' && category !== selectedCat) return false;

        // 3. กรองช่วงเวลา (วัน, สัปดาห์, เดือน, 3 เดือน)
        if (selectedTime !== 'all') {
            const dateInKey = Object.keys(row).find(k => k.includes('เข้า') || k.includes('รับ'));
            const dateDoneKey = Object.keys(row).find(k => k.includes('แล้วเสร็จ'));
            let dateToUse = dateInKey ? row[dateInKey] : (dateDoneKey ? row[dateDoneKey] : null);

            if (!dateToUse) return false;

            let d = (dateToUse instanceof Date) ? dateToUse : new Date(dateToUse);
            if (isNaN(d.getTime())) return false;

            // เทียบความห่างของวัน
            const diffTime = Math.abs(maxDate - d);
            const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

            if (diffDays > parseInt(selectedTime)) return false;
        }

        return true;
    });

    updateDashboard(filteredData);
}

function updateDashboard(data) {
    let completedCount = 0;
    let pendingCount = 0;

    const categoryCounts = {};
    const sourceSet = new Set();
    const monthlyCounts = {};
    const recentItems = [];

    data.forEach(row => {
        const titleKey = Object.keys(row).find(k => k.includes('ชื่อเรื่อง'));
        const catKey = Object.keys(row).find(k => k.includes('ประเภท') || k.includes('งานที่ลง'));
        const srcKey = Object.keys(row).find(k => k.includes('ที่มา'));
        const dateDoneKey = Object.keys(row).find(k => k.includes('แล้วเสร็จ'));
        const dateInKey = Object.keys(row).find(k => k.includes('เข้า') || k.includes('รับ'));

        const title = titleKey ? row[titleKey] : '-';
        const category = (catKey && row[catKey]) ? row[catKey].toString().trim() : 'ไม่ระบุประเภท';
        const source = (srcKey && row[srcKey]) ? row[srcKey] : 'ไม่ระบุที่มา';
        const dateDone = dateDoneKey ? row[dateDoneKey] : null;

        let isCompleted = false;
        if (dateDone && dateDone.toString().trim() !== '') {
            isCompleted = true;
            completedCount++;
        } else {
            pendingCount++;
        }

        if (source !== 'ไม่ระบุที่มา') sourceSet.add(source);

        categoryCounts[category] = (categoryCounts[category] || 0) + 1;

        let dateToUse = dateInKey ? row[dateInKey] : dateDone;
        let dateObj = null;

        if (dateToUse) {
            dateObj = (dateToUse instanceof Date) ? dateToUse : new Date(dateToUse);
            if (!isNaN(dateObj.getTime())) {
                const monthYear = dateObj.toLocaleDateString('th-TH', { month: 'short', year: '2-digit' });
                monthlyCounts[monthYear] = (monthlyCounts[monthYear] || 0) + 1;
            }
        }

        recentItems.push({
            website: row._website,
            title: title,
            category: category,
            date: dateToUse,
            status: isCompleted ? 'แล้วเสร็จ' : 'กำลังดำเนินการ'
        });
    });

    // อัปเดตตัวเลข
    document.getElementById('totalItems').innerText = data.length.toLocaleString();
    document.getElementById('completedItems').innerText = completedCount.toLocaleString();
    document.getElementById('pendingItems').innerText = pendingCount.toLocaleString();
    document.getElementById('totalSources').innerText = sourceSet.size.toLocaleString();

    // อัปเดตกราฟและตาราง
    updateCategoryChart(categoryCounts);
    updateTrendChart(monthlyCounts);

    recentItems.sort((a, b) => {
        let da = new Date(a.date), db = new Date(b.date);
        return db - da;
    });
    updateTable(recentItems.slice(0, 15));
}

function updateCategoryChart(categoryData) {
    const ctx = document.getElementById('categoryChart').getContext('2d');
    if (categoryChartInstance) categoryChartInstance.destroy();

    const sortedCategories = Object.entries(categoryData).sort((a, b) => b[1] - a[1]);
    let labels = [];
    let data = [];
    let otherCount = 0;

    sortedCategories.forEach((item, index) => {
        if (index < 5) {
            labels.push(item[0]);
            data.push(item[1]);
        } else {
            otherCount += item[1];
        }
    });

    if (otherCount > 0) {
        labels.push('อื่นๆ');
        data.push(otherCount);
    }

    categoryChartInstance = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: labels,
            datasets: [{
                data: data,
                backgroundColor: ['#3b82f6', '#10b981', '#f59e0b', '#8b5cf6', '#ef4444', '#64748b'],
                borderWidth: 0,
                hoverOffset: 4
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            cutout: '65%',
            plugins: {
                legend: {
                    position: 'right',
                    labels: { color: '#94a3b8', usePointStyle: true, font: { family: 'Kanit' } }
                },
                datalabels: {
                    color: '#ffffff',
                    font: { family: 'Kanit', weight: 'bold', size: 12 },
                    formatter: (value, context) => {
                        let sum = context.chart.data.datasets[0].data.reduce((a, b) => a + b, 0);
                        let percentage = (value * 100 / sum).toFixed(1) + "%";
                        return value + ' \n(' + percentage + ')';
                    },
                    textAlign: 'center'
                }
            }
        }
    });
}

function updateTrendChart(monthlyData) {
    const ctx = document.getElementById('trendChart').getContext('2d');
    if (trendChartInstance) trendChartInstance.destroy();

    const labels = Object.keys(monthlyData);
    const data = Object.values(monthlyData);

    trendChartInstance = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: 'จำนวนเรื่องที่ลง',
                data: data,
                backgroundColor: 'rgba(59, 130, 246, 0.8)',
                borderRadius: 4
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            layout: { padding: { top: 20 } },
            plugins: {
                legend: { display: false },
                datalabels: {
                    anchor: 'end',
                    align: 'top',
                    color: '#94a3b8',
                    font: { family: 'Kanit', weight: 'bold', size: 14 },
                    formatter: (value) => value
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    grid: { color: 'rgba(255, 255, 255, 0.05)' },
                    ticks: { color: '#94a3b8', stepSize: 1, font: { family: 'Kanit' } }
                },
                x: {
                    grid: { display: false },
                    ticks: { color: '#94a3b8', font: { family: 'Kanit' } }
                }
            }
        }
    });
}

function formatThaiDate(dateValue) {
    if (!dateValue) return '-';
    let d = (dateValue instanceof Date) ? dateValue : new Date(dateValue);
    if (isNaN(d.getTime())) return String(dateValue);
    return d.toLocaleDateString('th-TH', { year: 'numeric', month: 'short', day: 'numeric' });
}

function updateTable(items) {
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = '';

    if (items.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6" style="text-align:center;">ไม่มีข้อมูลที่ตรงกับเงื่อนไขการกรอง</td></tr>';
        return;
    }

    items.forEach((item, index) => {
        const tr = document.createElement('tr');
        let statusClass = item.status === 'แล้วเสร็จ' ? 'status-completed' : 'status-pending';

        tr.innerHTML = `
            <td>${index + 1}</td>
            <td><span style="color: #3b82f6; font-weight: 500;">${item.website}</span></td>
            <td style="max-width: 250px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="${item.title}">${item.title}</td>
            <td>${item.category}</td>
            <td>${formatThaiDate(item.date)}</td>
            <td><span class="status-badge ${statusClass}">${item.status}</span></td>
        `;
        tbody.appendChild(tr);
    });
}