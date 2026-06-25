// Main JavaScript for OT Recording System - Direct Recording

let attendanceRecords = {}; // Store attendance status for each employee
let resetMode = false; // Track if we're in reset mode
let currentOTSubEmployee = null; // Store current employee for OT substitute modal
let allWorkplaces = []; // Store all available workplaces (for report filter)
let allGroups = [];     // Store all groups (for OT substitute modal)
let currentDisplayedEmployees = []; // Employees currently shown after filters
let currentLeaveEmployee = null; // Store current employee for leave option modal
const preferredSupervisors = [
    { match: 'ภาคภูมิพรมชา', fallback: 'ภาคภูมิ พรมชา', displayName: 'ภาคภูมิ พรมชา (SV)' },
    { match: 'ศุภมาศปลงจิตร', fallback: 'ศุภมาศ ปลงจิตร', displayName: 'ศุภมาศ ปลงจิตร (FM)' },
    { match: 'วัชราศรีบาลชื่น', fallback: 'วัชรา ศรีบาลชื่น', displayName: 'วัชรา ศรีบาลชื่น (FM)' },
    { match: 'ประถมสุขรัมย์', fallback: 'ประถม สุขรัมย์', displayName: 'ประถม สุขรัมย์ (FM)' }
];

// Initialize page
document.addEventListener('DOMContentLoaded', function() {
    // Set today's date
    const today = new Date().toISOString().split('T')[0];
    document.getElementById('workDate').value = today;

    // Load filters
    loadSupervisors();
    loadWorkplaces();
    loadGroups();

    // Set up event listeners
    document.getElementById('glFilter').addEventListener('change', filterEmployees);
    document.getElementById('processFilter')?.addEventListener('change', filterEmployees);
    document.getElementById('shiftFilter').addEventListener('change', function() {
        loadSupervisors();
        filterEmployees();
    });
    document.getElementById('workDate').addEventListener('change', filterEmployees);

    // Initial load
    filterEmployees();
});

// Load supervisors/GLs filtered by shift
function loadSupervisors() {
    const shift = document.getElementById('shiftFilter').value;
    fetch(`/api/supervisors?shift=${shift}`)
        .then(response => response.json())
        .then(data => {
            const select = document.getElementById('glFilter');
            // includeMissingPreferred เฉพาะตอน "ทั้งหมด" — ถ้าเลือก shift เฉพาะให้แสดงแค่ GL ของ shift นั้น
            const supervisors = sortSupervisorOptions(data, shift === 'All');
            const prevValue = select.value;
            select.innerHTML = '<option value="All">All</option>';
            supervisors.forEach(supervisor => {
                const option = document.createElement('option');
                option.value = supervisor;
                option.textContent = getSupervisorDisplayName(supervisor);
                select.appendChild(option);
            });
            // คงค่าเดิมถ้ายังอยู่ใน list — ถ้าไม่มีให้ reset เป็น All
            select.value = supervisors.includes(prevValue) ? prevValue : 'All';
        });
}

// Load workplaces/processes
function loadWorkplaces() {
    fetch('/api/workplaces')
        .then(response => response.json())
        .then(data => {
            allWorkplaces = data;

            const select = document.getElementById('processFilter');
            if (!select) return;
            select.innerHTML = '<option value="All">All</option>';
            data.forEach(workplace => {
                const option = document.createElement('option');
                option.value = workplace;
                option.textContent = workplace;
                select.appendChild(option);
            });
        });
}

// Load groups for OT substitute dropdown
function loadGroups() {
    fetch('/api/groups')
        .then(response => response.json())
        .then(data => { allGroups = data; });
}

// Update summary cards
function updateSummaryCards(workDate) {
    fetch(`/api/attendance/history?work_date=${workDate}`)
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                // Summary cards are recalculated against the currently filtered employees in displayEmployees().
            }
        })
        .catch(() => {});
}

// Filter employees and load attendance status
function filterEmployees() {
    const gl      = document.getElementById('glFilter').value;
    const process = document.getElementById('processFilter')?.value || 'All';
    const shift   = document.getElementById('shiftFilter').value;
    const workDate = document.getElementById('workDate').value;

    updateSummaryCards(workDate);

    // เมื่อเลือก GL เฉพาะ → ดึง incoming substitutes ด้วย
    const incomingUrl = (gl && gl !== 'All')
        ? `/api/attendance/incoming-substitutes?work_date=${workDate}&gl=${encodeURIComponent(gl)}&shift=${shift}`
        : null;

    Promise.all([
        fetch(`/api/employees/filter?gl=${gl}&process=${process}&shift=${shift}`).then(r => r.json()),
        fetch(`/api/attendance/check?work_date=${workDate}`).then(r => r.json()),
        incomingUrl ? fetch(incomingUrl).then(r => r.json()) : Promise.resolve([])
    ])
    .then(([employees, attendance, incomingSubs]) => {
        attendanceRecords = attendance;
        displayEmployees(employees, incomingSubs);
    })
    .catch(error => {
        console.error('Error:', error);
    });
}

// Display employees in table with attendance status
function displayEmployees(employees, incomingSubs = []) {
    const tbody = document.getElementById('employeeTable');
    const colspanCount = resetMode ? 7 : 7; // Table always has 7 visible data columns in the refreshed UI

    currentDisplayedEmployees = employees;
    updateLeaderBulkActions();
    updateFilteredSummaryCards(employees);
    updateShiftHoursSummary(employees);
    updateLeaderSummary(employees, incomingSubs);

    if (employees.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="${colspanCount}" class="empty-cell">
                    <i class="bi bi-inbox"></i> ไม่พบพนักงานตามเงื่อนไขที่เลือก
                </td>
            </tr>
        `;
        return;
    }

    tbody.innerHTML = '';
    employees.forEach(emp => {
        const row = document.createElement('tr');
        const empAttendance = attendanceRecords[emp.emp_id] || { types: [], workplace_substitute: null, leave_period: null, leave_time: null };
        const attendanceTypes = empAttendance.types || [];

        // Check which types are already recorded
        const hasWork = attendanceTypes.includes('work');
        const hasLeave = attendanceTypes.includes('leave');
        const hasOT = attendanceTypes.includes('ot');
        const workplaceSubstitute = empAttendance.workplace_substitute;
        const leaveLabel = empAttendance.leave_period && empAttendance.leave_time
            ? `${empAttendance.leave_period} (${empAttendance.leave_time})`
            : 'ลางาน';

        // Disable conflicting buttons
        const workDisabled = hasLeave;  // ถ้าลาแล้ว ปุ่มมาทำงานจะ disable
        const leaveDisabled = hasWork;  // ถ้ามาทำงานแล้ว ปุ่มลางานจะ disable
        const otDisabled = hasLeave || !hasWork;    // ต้องบันทึกมาทำงานก่อนจึงทำ OT ได้

        // Show workplace only when OT is recorded
        let otWorkplace = '-';
        if (hasOT) {
            if (workplaceSubstitute === 'OT ต่างแผนก') {
                otWorkplace = `<span class="badge" style="background:#6f42c1">🔀 ต่างแผนก</span>`;
            } else if (workplaceSubstitute) {
                otWorkplace = `<span class="badge bg-warning text-dark">${workplaceSubstitute} <small>(แทน)</small></span>`;
            } else {
                otWorkplace = `<span class="badge bg-info">${emp.workplace || '-'}</span>`;
            }
        }

        // Checkbox column - only show in reset mode
        const checkboxColumn = resetMode ? `
            <td class="text-center checkbox-column">
                <input type="checkbox" class="form-check-input employee-checkbox" value="${emp.emp_id}">
            </td>
        ` : '';

        // Escape HTML to prevent XSS and quote issues
        const safeEmpId = escapeHtml(emp.emp_id);
        const safeEmpName = escapeHtml(emp.name);
        const safeShift = escapeHtml(emp.shift);
        const safeGroupName = escapeHtml(emp.group_name || '');

        row.innerHTML = `
            ${checkboxColumn}
            <td>${safeEmpId}</td>
            <td>${safeEmpName}</td>
            <td><span class="badge bg-${emp.emp_type === 'MCP' ? 'success' : 'info'}">${emp.emp_type}</span></td>
            <td class="text-center"><span class="badge bg-secondary">${safeShift}</span></td>
            <td>${otWorkplace}</td>
            <td>
                <div class="btn-group btn-group-sm" role="group">
                    <button type="button" class="btn ${hasWork ? 'btn-success' : 'btn-outline-success'} ${workDisabled ? 'disabled' : ''}"
                            data-action="work" data-emp-id="${safeEmpId}" data-emp-name="${safeEmpName}" data-is-recorded="${hasWork}"
                            title="${workDisabled ? 'ไม่สามารถกดได้เนื่องจากบันทึกลางานแล้ว' : (hasWork ? 'คลิกเพื่อยกเลิก' : 'ทำงาน')}"
                            ${workDisabled ? 'disabled' : ''}>
                        <i class="bi ${hasWork ? 'bi-check-circle-fill' : 'bi-check-circle'}"></i>
                        ทำงาน
                    </button>
                    <button type="button" class="btn ${hasLeave ? 'btn-leave-active' : 'btn-leave'} ${leaveDisabled ? 'disabled' : ''}"
                            data-action="leave" data-emp-id="${safeEmpId}" data-emp-name="${safeEmpName}" data-is-recorded="${hasLeave}" data-shift="${safeShift}"
                            title="${leaveDisabled ? 'ไม่สามารถกดได้เนื่องจากบันทึกมาทำงานแล้ว' : (hasLeave ? leaveLabel + ' - คลิกเพื่อยกเลิก' : 'ลางาน')}"
                            ${leaveDisabled ? 'disabled' : ''}>
                        <i class="bi ${hasLeave ? 'bi-calendar-x-fill' : 'bi-calendar-x'}"></i>
                        ลางาน
                    </button>
                    <button type="button" class="btn ${hasOT ? 'btn-primary' : 'btn-outline-primary'} ${otDisabled ? 'disabled' : ''}"
                            data-action="ot" data-emp-id="${safeEmpId}" data-emp-name="${safeEmpName}" data-is-recorded="${hasOT}"
                            title="${hasLeave ? 'ไม่สามารถกดได้เนื่องจากบันทึกลางานแล้ว' : (!hasWork ? 'กรุณาบันทึกมาทำงานก่อน' : (hasOT ? 'คลิกเพื่อยกเลิก' : 'ทำ OT'))}"
                            ${otDisabled ? 'disabled' : ''}>
                        <i class="bi ${hasOT ? 'bi-clock-fill' : 'bi-clock'}"></i>
                        OT
                    </button>
                    <button type="button" class="btn ${hasOT && workplaceSubstitute ? 'btn-substitute-active' : 'btn-substitute'} ${otDisabled || hasOT ? 'disabled' : ''}"
                            data-action="ot-substitute" data-emp-id="${safeEmpId}" data-emp-name="${safeEmpName}" data-group="${safeGroupName}"
                            title="${hasLeave ? 'ไม่สามารถกดได้เนื่องจากบันทึกลางานแล้ว' : (!hasWork ? 'กรุณาบันทึกมาทำงานก่อน' : (hasOT ? 'ยกเลิก OT แทนก่อน' : 'ทำ OT แทน'))}"
                            ${otDisabled || hasOT ? 'disabled' : ''}>
                        <i class="bi bi-arrow-left-right"></i>
                        OT แทน
                    </button>
                </div>
            </td>
        `;
        tbody.appendChild(row);
    });

    // Reset "Select All" checkbox state
    document.getElementById('selectAllEmployees').checked = false;

    // Add event delegation for attendance buttons
    setupAttendanceButtonListeners();
}

// Event delegation for dynamically created buttons
function setupAttendanceButtonListeners() {
    const tbody = document.getElementById('employeeTable');

    // Remove existing listener if any to prevent duplicate listeners
    tbody.removeEventListener('click', handleAttendanceButtonClick);
    tbody.addEventListener('click', handleAttendanceButtonClick);
}

function handleAttendanceButtonClick(event) {
    const button = event.target.closest('button[data-action]');

    if (!button) return;
    if (button.disabled) return;

    const action = button.getAttribute('data-action');
    const empId = button.getAttribute('data-emp-id');
    const empName = button.getAttribute('data-emp-name');
    const isRecorded = button.getAttribute('data-is-recorded') === 'true';

    try {
        switch (action) {
            case 'work':
                toggleAttendance(empId, empName, 'work', isRecorded);
                break;
            case 'leave':
                const shift = button.getAttribute('data-shift');
                toggleAttendance(empId, empName, 'leave', isRecorded, shift);
                break;
            case 'ot':
                toggleAttendance(empId, empName, 'ot', isRecorded);
                break;
            case 'ot-substitute':
                const groupName = button.getAttribute('data-group');
                showOTSubstituteModal(empId, empName, groupName);
                break;
        }
    } catch (error) {
        console.error('Error handling button click:', error);
        alert('เกิดข้อผิดพลาด: ' + error.message);
    }
}

function updateLeaderBulkActions() {
    const selectedLeader = document.getElementById('glFilter').value;
    const showBulkActions = selectedLeader && selectedLeader !== 'All' && currentDisplayedEmployees.length > 0;

    document.getElementById('markAllWorkBtn')?.classList.toggle('d-none', !showBulkActions);

    const otBtn = document.getElementById('markAllOTBtn');
    if (!otBtn) return;
    otBtn.classList.toggle('d-none', !showBulkActions);

    if (showBulkActions) {
        // ตรวจสอบว่าพนักงานทุกคนได้รับการยืนยันสถานะ (work หรือ leave) แล้วหรือยัง
        const allConfirmed = currentDisplayedEmployees.every(emp => {
            const types = (attendanceRecords[emp.emp_id]?.types) || [];
            return types.includes('work') || types.includes('leave');
        });

        otBtn.disabled = !allConfirmed;
        otBtn.title = allConfirmed
            ? 'บันทึก OT ให้พนักงานทุกคนที่มาทำงาน'
            : 'กรุณายืนยันสถานะ (มาทำงาน/ลางาน) ให้ครบทุกคนก่อน';
    }
}

async function markAllForSelectedLeader(attendanceType) {
    const selectedLeader = document.getElementById('glFilter').value;
    const workDate = document.getElementById('workDate').value;

    if (!selectedLeader || selectedLeader === 'All') {
        alert('กรุณาเลือกหัวหน้างานก่อน');
        return;
    }

    if (!workDate) {
        alert('กรุณาเลือกวันที่ทำงาน');
        return;
    }

    if (currentDisplayedEmployees.length === 0) {
        alert('ไม่มีพนักงานในเงื่อนไขที่เลือก');
        return;
    }

    const actionText = attendanceType === 'work' ? 'มาทำงานครบ' : 'ทำ OT ทุกคน';
    if (!confirm(`ต้องการบันทึก "${actionText}" ให้พนักงาน ${currentDisplayedEmployees.length} คน ภายใต้หัวหน้า ${selectedLeader} ใช่หรือไม่?`)) {
        return;
    }

    setBulkButtonsLoading(true);

    let successCount = 0;
    let errorCount = 0;

    for (const emp of currentDisplayedEmployees) {
        const empAttendance = attendanceRecords[emp.emp_id] || { types: [], workplace_substitute: null };
        const attendanceTypes = empAttendance.types || [];

        try {
            if (attendanceTypes.includes('leave')) {
                await deleteAttendanceRecord(workDate, emp.emp_id, 'leave');
            }

            if (!attendanceTypes.includes(attendanceType)) {
                await recordAttendanceForBulk(workDate, emp.emp_id, attendanceType);
            }

            successCount++;
        } catch (error) {
            console.error(`Bulk ${attendanceType} error for ${emp.emp_id}:`, error);
            errorCount++;
        }
    }

    setBulkButtonsLoading(false);

    if (errorCount === 0) {
        showSuccessMessage(`บันทึก "${actionText}" สำเร็จ ${successCount} คน`);
    } else {
        alert(`บันทึกสำเร็จ ${successCount} คน, ล้มเหลว ${errorCount} คน`);
    }

    filterEmployees();
}

function setBulkButtonsLoading(isLoading) {
    const workBtn = document.getElementById('markAllWorkBtn');
    const otBtn = document.getElementById('markAllOTBtn');

    [workBtn, otBtn].forEach(button => {
        if (!button) return;
        button.disabled = isLoading;
    });

    if (workBtn) {
        workBtn.innerHTML = isLoading
            ? '<span class="spinner-border spinner-border-sm"></span> กำลังบันทึก'
            : '<i class="bi bi-check2-circle"></i> มาทำงานครบ';
    }

    if (otBtn) {
        otBtn.innerHTML = isLoading
            ? '<span class="spinner-border spinner-border-sm"></span> กำลังบันทึก'
            : '<i class="bi bi-stopwatch"></i> ทำ OT ทุกคน';
    }
}

function recordAttendanceForBulk(workDate, empId, attendanceType) {
    return fetch('/api/attendance/record', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            work_date: workDate,
            emp_id: empId,
            attendance_type: attendanceType
        })
    })
    .then(response => response.json())
    .then(result => {
        if (!result.success) {
            throw new Error(result.message);
        }
        return result;
    });
}

function updateFilteredSummaryCards(employees) {
    const total = employees.length;
    const summary = employees.reduce((acc, emp) => {
        const empAttendance = attendanceRecords[emp.emp_id] || { types: [], workplace_substitute: null };
        const attendanceTypes = empAttendance.types || [];
        if (attendanceTypes.includes('work')) acc.work += 1;
        if (attendanceTypes.includes('leave')) acc.leave += 1;
        if (attendanceTypes.includes('ot')) acc.ot += 1;
        return acc;
    }, { work: 0, leave: 0, ot: 0 });

    updateSummaryMetric('summaryWork', 'summaryWorkPercent', 'summaryWorkBar', summary.work, total);
    updateSummaryMetric('summaryLeave', 'summaryLeavePercent', 'summaryLeaveBar', summary.leave, total);
    updateSummaryMetric('summaryOT', 'summaryOTPercent', 'summaryOTBar', summary.ot, total);

    document.querySelectorAll('.summaryTotal').forEach(element => {
        element.textContent = total;
    });
}

// ชั่วโมงการทำงานต่อคน: มาทำงาน 8.16 ชม. + ทำ OT เพิ่มอีก 2.32 ชม.
const WORK_HOURS_PER_PERSON = 8.16;
const OT_HOURS_PER_PERSON = 2.32;
const SHIFT_ORDER = ['A', 'B', 'D'];

// คำนวณและแสดงชั่วโมงการทำงานแยกตาม Shift
function updateShiftHoursSummary(employees) {
    const grid = document.getElementById('shiftHoursGrid');
    if (!grid) return;

    if (!employees.length) {
        grid.innerHTML = '<div class="empty-cell compact"><i class="bi bi-inbox"></i> ไม่มีข้อมูลพนักงานตามเงื่อนไขที่เลือก</div>';
        return;
    }

    const shiftMap = new Map();
    employees.forEach(emp => {
        const shift = emp.shift || 'ไม่ระบุ';
        if (!shiftMap.has(shift)) {
            shiftMap.set(shift, { shift, work: 0, ot: 0 });
        }
        const summary = shiftMap.get(shift);
        const attendanceTypes = (attendanceRecords[emp.emp_id] || {}).types || [];
        if (attendanceTypes.includes('work')) summary.work += 1;
        if (attendanceTypes.includes('ot')) summary.ot += 1;
    });

    const summaries = [...shiftMap.values()].sort((a, b) => {
        const ia = SHIFT_ORDER.indexOf(a.shift);
        const ib = SHIFT_ORDER.indexOf(b.shift);
        return (ia === -1 ? 99 : ia) - (ib === -1 ? 99 : ib);
    });

    const totals = summaries.reduce((acc, s) => {
        acc.work += s.work;
        acc.ot += s.ot;
        return acc;
    }, { work: 0, ot: 0 });

    let html = summaries.map(s => renderShiftHoursCard(s.shift, s.work, s.ot)).join('');
    if (summaries.length > 1) {
        html += renderShiftHoursCard('รวมทุก Shift', totals.work, totals.ot, true);
    }
    grid.innerHTML = html;
}

function shiftHours(work, ot) {
    return (work * WORK_HOURS_PER_PERSON) + (ot * OT_HOURS_PER_PERSON);
}

function renderShiftHoursCard(label, work, ot, isTotal = false) {
    const hours = shiftHours(work, ot);
    return `
        <div class="shift-hours-card${isTotal ? ' shift-hours-total' : ''}">
            <div class="shift-hours-head">
                ${isTotal ? '' : '<span class="shift-hours-label">Shift</span>'}
                <span class="shift-hours-name">${escapeHtml(String(label))}</span>
            </div>
            <div class="shift-hours-value">
                <strong>${hours.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</strong>
                <span>ชม.</span>
            </div>
            <div class="shift-hours-detail">
                <span><i class="bi bi-check-circle-fill"></i> มาทำงาน ${work} คน</span>
                <span><i class="bi bi-stopwatch-fill"></i> OT ${ot} คน</span>
            </div>
        </div>
    `;
}

function updateSummaryMetric(countId, percentId, barId, count, total) {
    const percent = total > 0 ? Math.round((count / total) * 100) : 0;
    const countEl = document.getElementById(countId);
    const percentEl = document.getElementById(percentId);
    const barEl = document.getElementById(barId);

    if (countEl) countEl.textContent = count;
    if (percentEl) percentEl.textContent = `${percent}%`;
    if (barEl) barEl.style.width = `${percent}%`;
}

function updateLeaderSummary(employees, incomingSubs = []) {
    const panel = document.getElementById('leaderSummaryPanel');
    const grid = document.getElementById('leaderSummaryGrid');
    const caption = document.getElementById('leaderSummaryCaption');

    if (!panel || !grid) return;

    if (!employees.length) {
        grid.innerHTML = '<div class="empty-cell compact"><i class="bi bi-inbox"></i> ไม่มีข้อมูลพนักงานตามเงื่อนไขที่เลือก</div>';
        return;
    }

    const selectedLeader = document.getElementById('glFilter').value;
    const leaderMap = new Map();

    employees.forEach(emp => {
        const leader = emp.supervisor || 'ไม่ระบุหัวหน้างาน';
        if (!leaderMap.has(leader)) {
            leaderMap.set(leader, { leader, total: 0, work: 0, leave: 0, ot: 0, substitute: 0 });
        }

        const summary = leaderMap.get(leader);
        const empAttendance = attendanceRecords[emp.emp_id] || { types: [], workplace_substitute: null };
        const attendanceTypes = empAttendance.types || [];
        summary.total += 1;
        if (attendanceTypes.includes('work')) summary.work += 1;
        if (attendanceTypes.includes('leave')) summary.leave += 1;
        if (attendanceTypes.includes('ot')) summary.ot += 1;
        if (attendanceTypes.includes('ot') && empAttendance.workplace_substitute) summary.substitute += 1;
    });

    const summaries = [...leaderMap.values()].sort((a, b) => sortLeaderSummary(a, b, selectedLeader));

    caption.textContent = selectedLeader === 'All'
        ? 'แสดงทุกหัวหน้างานตามเงื่อนไขที่เลือก'
        : `แสดงเฉพาะหัวหน้า ${selectedLeader}`;

    const showSubstitute  = selectedLeader !== 'All';
    const singleLeader    = selectedLeader !== 'All';   // เมื่อเลือก GL เดียว → ยืดเต็ม
    let html = summaries.map(summary => renderLeaderSummaryCard(summary, showSubstitute, singleLeader)).join('');

    // แสดงพนักงานจากทีมอื่นที่มาทำ OT ให้ทีมนี้
    if (selectedLeader !== 'All' && incomingSubs.length > 0) {
        html += renderIncomingSubstitutes(incomingSubs);
    }

    grid.innerHTML = html;
}

function renderIncomingSubstitutes(subs) {
    const listHtml = subs.map(sub => `
        <div class="incoming-sub-item">
            <div class="incoming-sub-emp">
                <span class="fw-bold">${escapeHtml(sub.emp_id)} ${escapeHtml(sub.name)}</span>
                <span class="badge bg-${sub.emp_type === 'MCP' ? 'success' : 'info'} ms-1">${sub.emp_type}</span>
                <span class="badge bg-secondary ms-1">Shift ${sub.shift}</span>
            </div>
            <div class="incoming-sub-detail">
                <i class="bi bi-arrow-right-circle-fill incoming-arrow"></i>
                <span class="fw-semibold text-primary">${escapeHtml(sub.substitute_group)}</span>
                <span class="text-muted ms-2" style="font-size:0.75rem">
                    จาก: ${escapeHtml(sub.original_group)} · GL: ${escapeHtml(sub.original_gl)}
                </span>
            </div>
        </div>
    `).join('');

    return `
        <div class="incoming-sub-card">
            <div class="incoming-sub-header">
                <i class="bi bi-person-fill-up"></i>
                พนักงานจากทีมอื่นที่มาทำ OT ให้ทีมนี้
                <span class="badge bg-primary rounded-pill ms-1">${subs.length} คน</span>
            </div>
            <div class="incoming-sub-list">${listHtml}</div>
        </div>
    `;
}

function renderLeaderSummaryCard(summary, showSubstitute = false, fullWidth = false) {
    const total = Math.max(summary.total, 1);
    const workPercent = Math.round((summary.work / total) * 100);
    const leavePercent = Math.round((summary.leave / total) * 100);
    const otPercent = Math.round((summary.ot / total) * 100);
    const substitutePercent = Math.round((summary.substitute / total) * 100);

    return `
        <article class="leader-card${fullWidth ? ' leader-card-full' : ''}">
            <div class="leader-card-top">
                <div>
                    <span class="leader-label">หัวหน้างาน</span>
                    <h3>${escapeHtml(summary.leader)}</h3>
                </div>
                <span class="leader-total">${summary.total} คน</span>
            </div>
            <div class="leader-stats">
                ${renderLeaderStat('มาทำงาน', summary.work, summary.total, workPercent, 'work')}
                ${renderLeaderStat('ลางาน', summary.leave, summary.total, leavePercent, 'leave')}
                ${renderLeaderStat('ทำ OT', summary.ot, summary.total, otPercent, 'ot')}
                ${showSubstitute ? renderLeaderStat('OT แทน', summary.substitute, summary.total, substitutePercent, 'substitute') : ''}
            </div>
        </article>
    `;
}

function sortLeaderSummary(a, b, selectedLeader) {
    const aRank = getPreferredSupervisorRank(a.leader);
    const bRank = getPreferredSupervisorRank(b.leader);

    if (selectedLeader !== 'All') return a.leader.localeCompare(b.leader, 'th');
    if (aRank !== bRank) return aRank - bRank;
    if (aRank < preferredSupervisors.length) return a.leader.localeCompare(b.leader, 'th');
    return b.total - a.total || a.leader.localeCompare(b.leader, 'th');
}

function sortSupervisorOptions(supervisors, includeMissingPreferred = false) {
    const values = [...new Set((supervisors || []).map(value => (value || '').trim()).filter(Boolean))];

    if (includeMissingPreferred) {
        preferredSupervisors.forEach(meta => {
            if (!values.some(value => isPreferredSupervisor(value, meta))) {
                values.push(meta.fallback);
            }
        });
    }

    return values.sort((a, b) => {
        const aRank = getPreferredSupervisorRank(a);
        const bRank = getPreferredSupervisorRank(b);
        if (aRank !== bRank) return aRank - bRank;
        return a.localeCompare(b, 'th');
    });
}

function getPreferredSupervisorRank(supervisorName) {
    const index = preferredSupervisors.findIndex(meta => isPreferredSupervisor(supervisorName, meta));
    return index === -1 ? preferredSupervisors.length : index;
}

function isPreferredSupervisor(supervisorName, meta) {
    const normalized = (supervisorName || '').replace(/\s+/g, '');
    return normalized.includes(meta.match);
}

function getSupervisorDisplayName(supervisorName) {
    const meta = preferredSupervisors.find(item => isPreferredSupervisor(supervisorName, item));
    return meta ? meta.displayName : supervisorName;
}

function renderLeaderStat(label, count, total, percent, type) {
    return `
        <div class="leader-stat leader-stat-${type}">
            <div class="leader-stat-line">
                <span>${label}</span>
                <strong>${count}/${total}</strong>
            </div>
            <div class="leader-progress" aria-hidden="true">
                <span style="width: ${percent}%"></span>
            </div>
        </div>
    `;
}

function escapeHtml(value) {
    return (value || '').toString()
        .replaceAll('&', '&amp;')
        .replaceAll('<', '&lt;')
        .replaceAll('>', '&gt;')
        .replaceAll('"', '&quot;')
        .replaceAll("'", '&#039;');
}

// Toggle select all employees
function toggleSelectAll() {
    const selectAllCheckbox = document.getElementById('selectAllEmployees');
    const employeeCheckboxes = document.querySelectorAll('.employee-checkbox');

    employeeCheckboxes.forEach(checkbox => {
        checkbox.checked = selectAllCheckbox.checked;
    });
}

// Toggle attendance (add if not exists, delete if exists)
async function toggleAttendance(empId, empName, attendanceType, isRecorded, shift = '') {
    const workDate = document.getElementById('workDate').value;

    if (!workDate) {
        alert('กรุณาเลือกวันที่ทำงาน');
        return;
    }

    const empAttendance = attendanceRecords[empId] || { types: [], workplace_substitute: null };
    const attendanceTypes = empAttendance.types || [];

    if (attendanceType === 'leave' && !isRecorded) {
        showLeaveOptionModal(empId, empName, shift);
        return;
    }

    // เช็คว่าปุ่มนี้ควรถูก disable หรือไม่ (ถ้าไม่ได้กดไว้แล้ว)
    if (!isRecorded) {
        const hasWork = attendanceTypes.includes('work');
        const hasLeave = attendanceTypes.includes('leave');
        const hasOT = attendanceTypes.includes('ot');

        // ถ้าจะกด "ทำงาน" แต่มี "ลางาน" อยู่แล้ว (ปุ่มควร disable)
        if (attendanceType === 'work' && hasLeave) {
            alert('ไม่สามารถบันทึกมาทำงานได้ เนื่องจากบันทึกลางานไว้แล้ว\nกรุณายกเลิกการลางานก่อน');
            return;
        }

        // ถ้าจะกด "ลางาน" แต่มี "ทำงาน" อยู่แล้ว (ปุ่มควร disable)
        if (attendanceType === 'leave' && hasWork) {
            alert('ไม่สามารถบันทึกลางานได้ เนื่องจากบันทึกมาทำงานไว้แล้ว\nกรุณายกเลิกการมาทำงานก่อน');
            return;
        }

        // ถ้าจะกด "ลางาน" แต่มี "OT" อยู่แล้ว
        if (attendanceType === 'leave' && hasOT) {
            // ถามว่าจะลบ OT ด้วยไหม
            let confirmMessage = 'พบการบันทึก "OT"\nระบบจะลบการบันทึก OT และบันทึกเป็น "ลางาน" แทน\n\nต้องการดำเนินการต่อหรือไม่?';

            if (!confirm(confirmMessage)) {
                return;
            }

            // ลบ "OT"
            try {
                await deleteAttendanceRecord(workDate, empId, 'ot');
            } catch (error) {
                console.error('Error deleting OT record:', error);
                alert('เกิดข้อผิดพลาดในการลบการบันทึก OT');
                return;
            }
        }

        // ถ้าจะกด "OT" แต่มี "ลางาน" อยู่แล้ว (ปุ่มควร disable)
        if (attendanceType === 'ot' && hasLeave) {
            alert('ไม่สามารถบันทึก OT ได้ เนื่องจากบันทึกลางานไว้แล้ว\nกรุณายกเลิกการลางานก่อน');
            return;
        }

        if (attendanceType === 'ot' && !hasWork) {
            alert('กรุณาบันทึกมาทำงานก่อน จึงจะสามารถบันทึก OT ได้');
            return;
        }
    }

    const data = {
        work_date: workDate,
        emp_id: empId,
        attendance_type: attendanceType
    };

    // If already recorded, delete it. Otherwise, add it.
    const endpoint = isRecorded ? '/api/attendance/delete' : '/api/attendance/record';

    fetch(endpoint, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(data)
    })
    .then(response => response.json())
    .then(result => {
        if (result.success) {
            showSuccessMessage(result.message);
            filterEmployees();
        } else {
            alert('เกิดข้อผิดพลาด: ' + result.message);
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('เกิดข้อผิดพลาดในการบันทึกข้อมูล');
    });
}

// Helper function to delete attendance record
function deleteAttendanceRecord(workDate, empId, attendanceType) {
    return fetch('/api/attendance/delete', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            work_date: workDate,
            emp_id: empId,
            attendance_type: attendanceType
        })
    })
    .then(response => response.json())
    .then(result => {
        if (!result.success) {
            throw new Error(result.message);
        }
        return result;
    });
}

function showLeaveOptionModal(empId, empName, shift) {
    const workDate = document.getElementById('workDate').value;

    if (!workDate) {
        alert('กรุณาเลือกวันที่ทำงาน');
        return;
    }

    currentLeaveEmployee = { emp_id: empId, name: empName, shift };
    document.getElementById('leaveEmpInfo').textContent = `${empId} - ${empName}`;

    const options = getLeaveOptionsForShift(shift, workDate);

    const grid = document.getElementById('leaveOptionGrid');
    grid.innerHTML = options.map(option => `
        <button type="button" class="leave-option-card" onclick='confirmLeaveOption(${JSON.stringify(option)})'>
            <span>${option.label}</span>
            <strong>${option.time}</strong>
        </button>
    `).join('');

    const modal = new bootstrap.Modal(document.getElementById('leaveOptionModal'));
    modal.show();
}

// คำนวณกะเช้า/ดึก ตามวันที่ (A/B สลับทุกอาทิตย์, D เช้าตลอด)
function getLeaveOptionsForShift(shift, workDate) {
    const DAY_OPTIONS = [
        { label: 'เต็มวัน',      time: '07:40-17:00' },
        { label: 'ครึ่งวันเช้า', time: '07:40-12:50' },
        { label: 'ครึ่งวันบ่าย', time: '12:50-17:00' },
        { label: 'ลาย่อย 1',    time: '07:40-10:00' },
        { label: 'ลาย่อย 2',    time: '10:10-12:00' },
        { label: 'ลาย่อย 3',    time: '12:50-15:00' },
        { label: 'ลาย่อย 4',    time: '15:10-17:00' },
    ];
    const NIGHT_OPTIONS = [
        { label: 'เต็มวัน',      time: '19:40-05:00' },
        { label: 'ครึ่งวันแรก',  time: '19:40-00:50' },
        { label: 'ครึ่งวันหลัง', time: '00:50-05:00' },
        { label: 'ลาย่อย 1',    time: '19:40-22:00' },
        { label: 'ลาย่อย 2',    time: '22:10-00:00' },
        { label: 'ลาย่อย 3',    time: '00:50-03:00' },
        { label: 'ลาย่อย 4',    time: '03:10-05:00' },
    ];

    const s = (shift || '').toUpperCase();

    // Shift D เช้าตลอด
    if (s === 'D') return DAY_OPTIONS;

    // Shift A/B — คำนวณจากวันที่
    // จุดอ้างอิง: อาทิตย์ 18 พ.ค. 2026 (จันทร์) = Shift A เช้า
    const REFERENCE_MONDAY = new Date('2026-05-18T00:00:00');
    const date = new Date(workDate);

    // หาวันจันทร์ของอาทิตย์นั้น
    const dow = date.getDay(); // 0=อา, 1=จ, ...
    const daysToMonday = dow === 0 ? 6 : dow - 1;
    const monday = new Date(date);
    monday.setDate(date.getDate() - daysToMonday);
    monday.setHours(0, 0, 0, 0);

    // จำนวนอาทิตย์ห่างจาก reference
    const msPerWeek = 7 * 24 * 60 * 60 * 1000;
    const weeksDiff = Math.round((monday - REFERENCE_MONDAY) / msPerWeek);

    // คู่ = Shift A เช้า, คี่ = Shift A ดึก
    const shiftAIsDay = (weeksDiff % 2 === 0);

    if (s === 'A') return shiftAIsDay ? DAY_OPTIONS : NIGHT_OPTIONS;
    if (s === 'B') return shiftAIsDay ? NIGHT_OPTIONS : DAY_OPTIONS;

    return DAY_OPTIONS; // fallback
}

async function confirmLeaveOption(option) {
    const workDate = document.getElementById('workDate').value;

    if (!currentLeaveEmployee) {
        alert('ไม่พบข้อมูลพนักงาน');
        return;
    }

    const data = {
        work_date: workDate,
        emp_id: currentLeaveEmployee.emp_id,
        attendance_type: 'leave',
        leave_period: option.label,
        leave_time: option.time
    };

    fetch('/api/attendance/record', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(data)
    })
    .then(response => response.json())
    .then(result => {
        if (result.success) {
            const modal = bootstrap.Modal.getInstance(document.getElementById('leaveOptionModal'));
            modal.hide();
            showSuccessMessage(`บันทึกลางาน ${option.label} (${option.time}) สำหรับ ${currentLeaveEmployee.name} เรียบร้อยแล้ว`);
            currentLeaveEmployee = null;
            filterEmployees();
        } else {
            alert('เกิดข้อผิดพลาด: ' + result.message);
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('เกิดข้อผิดพลาดในการบันทึกข้อมูล');
    });
}

// Show OT Substitute Modal
function showOTSubstituteModal(empId, empName, originalGroup) {
    const workDate = document.getElementById('workDate').value;

    if (!workDate) {
        alert('กรุณาเลือกวันที่ทำงาน');
        return;
    }

    const empAttendance = attendanceRecords[empId] || { types: [], workplace_substitute: null };
    if (!(empAttendance.types || []).includes('work')) {
        alert('กรุณาบันทึกมาทำงานก่อน จึงจะสามารถบันทึก OT แทนได้');
        return;
    }

    // Store current employee info
    currentOTSubEmployee = {
        emp_id: empId,
        name: empName,
        original_group: originalGroup
    };

    // Populate modal
    document.getElementById('otSubEmpInfo').textContent = `${empId} - ${empName}`;
    document.getElementById('otSubOriginalProcess').textContent = originalGroup || '-';

    // Populate group select (exclude employee's own group)
    const select = document.getElementById('otSubProcessSelect');
    select.innerHTML = '<option value="">-- เลือก Group --</option>';

    allGroups.forEach(group => {
        if (group !== originalGroup) {
            const option = document.createElement('option');
            option.value = group;
            option.textContent = group;
            select.appendChild(option);
        }
    });

    // ── ตัวเลือกพิเศษ ───────────────────────────────
    const sep = document.createElement('option');
    sep.disabled = true;
    sep.textContent = '──────────────────────────';
    select.appendChild(sep);

    const otherDeptOpt = document.createElement('option');
    otherDeptOpt.value = 'OT ต่างแผนก';
    otherDeptOpt.textContent = '🔀 OT ต่างแผนก (ไม่ระบุ Group)';
    select.appendChild(otherDeptOpt);

    const otherOpt = document.createElement('option');
    otherOpt.value = 'อื่นๆ';
    otherOpt.textContent = '📝 อื่นๆ (ระบุเหตุผล)';
    select.appendChild(otherOpt);

    // Show modal
    const modal = new bootstrap.Modal(document.getElementById('otSubstituteModal'));
    modal.show();
}

// Toggle Other Reason Input
function toggleOtherReasonInput() {
    const select = document.getElementById('otSubProcessSelect');
    const container = document.getElementById('otOtherReasonContainer');
    const input = document.getElementById('otOtherReasonInput');

    if (select.value === 'อื่นๆ') {
        container.style.display = 'block';
        input.required = true;
    } else {
        container.style.display = 'none';
        input.required = false;
        input.value = '';
    }
}

// Confirm OT Substitute
async function confirmOTSubstitute() {
    const workDate = document.getElementById('workDate').value;
    const selectedWorkplace = document.getElementById('otSubProcessSelect').value;
    const otherReasonInput = document.getElementById('otOtherReasonInput');

    if (!selectedWorkplace) {
        alert('กรุณาเลือก Group ที่จะไปทำแทน');
        return;
    }

    // Check if "อื่นๆ" is selected and reason is provided
    if (selectedWorkplace === 'อื่นๆ') {
        if (!otherReasonInput.value.trim()) {
            alert('กรุณาระบุเหตุผลที่ต้องทำ OT แทน');
            otherReasonInput.focus();
            return;
        }
    }

    if (!currentOTSubEmployee) {
        alert('เกิดข้อผิดพลาด: ไม่พบข้อมูลพนักงาน');
        return;
    }

    const data = {
        work_date: workDate,
        emp_id: currentOTSubEmployee.emp_id,
        attendance_type: 'ot',
        workplace_substitute: selectedWorkplace === 'อื่นๆ'
            ? `อื่นๆ: ${otherReasonInput.value.trim()}`
            : selectedWorkplace
    };

    try {
        const response = await fetch('/api/attendance/record', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(data)
        });

        const result = await response.json();

        if (result.success) {
            // Close modal
            const modal = bootstrap.Modal.getInstance(document.getElementById('otSubstituteModal'));
            modal.hide();

            // Show success message
            showSuccessMessage(`บันทึก OT แทน Group ${selectedWorkplace} สำหรับ ${currentOTSubEmployee.name} เรียบร้อยแล้ว`);

            // Reload employees
            filterEmployees();

            // Reset current employee
            currentOTSubEmployee = null;
        } else {
            alert('เกิดข้อผิดพลาด: ' + result.message);
        }
    } catch (error) {
        console.error('Error:', error);
        alert('เกิดข้อผิดพลาดในการบันทึกข้อมูล');
    }
}

// Toggle reset mode (show/hide checkboxes)
function toggleResetMode() {
    resetMode = !resetMode;

    // Toggle checkbox column visibility
    const checkboxColumns = document.querySelectorAll('.checkbox-column');
    checkboxColumns.forEach(col => {
        if (resetMode) {
            col.classList.remove('d-none');
        } else {
            col.classList.add('d-none');
        }
    });

    // Toggle button visibility
    const resetBtn = document.getElementById('resetBtn');
    const confirmDeleteBtn = document.getElementById('confirmDeleteBtn');

    if (resetMode) {
        // Show "ลบที่เลือก" button and change Reset to "ยกเลิก"
        confirmDeleteBtn.classList.remove('d-none');
        resetBtn.innerHTML = '<i class="bi bi-x-lg"></i> ยกเลิก';
        resetBtn.classList.remove('btn-danger-subtle');
        resetBtn.classList.add('btn-secondary');
    } else {
        // Hide "ลบที่เลือก" button and change back to "Reset"
        confirmDeleteBtn.classList.add('d-none');
        resetBtn.innerHTML = '<i class="bi bi-x-circle"></i> Reset';
        resetBtn.classList.remove('btn-secondary');
        resetBtn.classList.add('btn-danger-subtle');

        // Uncheck all checkboxes
        document.getElementById('selectAllEmployees').checked = false;
        document.querySelectorAll('.employee-checkbox').forEach(cb => cb.checked = false);
    }

    // Reload employees to show/hide checkbox column
    filterEmployees();
}

// Confirm and delete selected employees
async function confirmDeleteSelected() {
    const workDate = document.getElementById('workDate').value;

    if (!workDate) {
        alert('กรุณาเลือกวันที่ทำงาน');
        return;
    }

    // Get all checked checkboxes
    const selectedCheckboxes = document.querySelectorAll('.employee-checkbox:checked');

    if (selectedCheckboxes.length === 0) {
        alert('กรุณาเลือกพนักงานที่ต้องการลบการบันทึก');
        return;
    }

    const confirmMessage = `ต้องการลบการบันทึกของพนักงาน ${selectedCheckboxes.length} คนที่เลือกใช่หรือไม่?\n\nการดำเนินการนี้ไม่สามารถยกเลิกได้`;

    if (!confirm(confirmMessage)) {
        return;
    }

    let successCount = 0;
    let errorCount = 0;

    try {
        for (const checkbox of selectedCheckboxes) {
            const empId = checkbox.value;
            const empAttendance = attendanceRecords[empId] || { types: [], workplace_substitute: null };
            const attendanceTypes = empAttendance.types || [];

            try {
                // Delete all attendance types for this employee
                for (const type of attendanceTypes) {
                    await deleteAttendanceRecord(workDate, empId, type);
                }
                successCount++;
            } catch (error) {
                console.error(`Error resetting ${empId}:`, error);
                errorCount++;
            }
        }

        // Show result message
        if (errorCount === 0) {
            showSuccessMessage(`ลบการบันทึกของพนักงาน ${successCount} คนเรียบร้อยแล้ว`);
        } else {
            alert(`ลบการบันทึกสำเร็จ ${successCount} คน, ล้มเหลว ${errorCount} คน`);
        }

        // Exit reset mode and reload employees
        toggleResetMode();
    } catch (error) {
        console.error('Error resetting attendance:', error);
        alert('เกิดข้อผิดพลาดในการลบการบันทึก');
    }
}

// Show reset modal with employee selection
function showResetModal() {
    const workDate = document.getElementById('workDate').value;

    if (!workDate) {
        alert('กรุณาเลือกวันที่ทำงาน');
        return;
    }

    // Format date for display
    const date = new Date(workDate);
    const formattedDate = `${String(date.getDate()).padStart(2, '0')}/${String(date.getMonth() + 1).padStart(2, '0')}/${date.getFullYear()}`;
    document.getElementById('resetDate').textContent = formattedDate;

    // Get employees with attendance records
    const employeesWithRecords = [];

    // Fetch current employee list to get names
    fetch(`/api/employees/filter?gl=All&process=All&shift=All`)
        .then(response => response.json())
        .then(employees => {
            // Filter employees who have attendance records
            employees.forEach(emp => {
                const empAttendance = attendanceRecords[emp.emp_id] || { types: [], workplace_substitute: null };
                const attendanceTypes = empAttendance.types || [];

                if (attendanceTypes.length > 0) {
                    let types = [];
                    if (attendanceTypes.includes('work')) types.push('มาทำงาน');
                    if (attendanceTypes.includes('leave')) types.push('ลางาน');
                    if (attendanceTypes.includes('ot')) {
                        if (empAttendance.workplace_substitute) {
                            types.push(`OT แทน (${empAttendance.workplace_substitute})`);
                        } else {
                            types.push('OT');
                        }
                    }

                    employeesWithRecords.push({
                        emp_id: emp.emp_id,
                        name: emp.name,
                        types: types
                    });
                }
            });

            // Populate the list
            const listDiv = document.getElementById('resetEmployeeList');
            const noRecordsDiv = document.getElementById('noRecordsMessage');

            if (employeesWithRecords.length === 0) {
                listDiv.style.display = 'none';
                noRecordsDiv.style.display = 'block';
                document.getElementById('selectAllReset').style.display = 'none';
            } else {
                listDiv.style.display = 'block';
                noRecordsDiv.style.display = 'none';
                document.getElementById('selectAllReset').style.display = 'block';

                listDiv.innerHTML = '';
                employeesWithRecords.forEach(emp => {
                    const checkboxDiv = document.createElement('div');
                    checkboxDiv.className = 'form-check mb-2 p-2 border rounded';
                    checkboxDiv.innerHTML = `
                        <input class="form-check-input reset-checkbox" type="checkbox" value="${emp.emp_id}" id="reset_${emp.emp_id}">
                        <label class="form-check-label" for="reset_${emp.emp_id}">
                            <strong>${emp.emp_id} - ${emp.name}</strong><br>
                            <small class="text-muted">การบันทึก: ${emp.types.join(', ')}</small>
                        </label>
                    `;
                    listDiv.appendChild(checkboxDiv);
                });

                // Setup select all checkbox
                document.getElementById('selectAllReset').checked = false;
                document.getElementById('selectAllReset').onclick = function() {
                    const checkboxes = document.querySelectorAll('.reset-checkbox');
                    checkboxes.forEach(cb => cb.checked = this.checked);
                };
            }

            // Show modal
            const modal = new bootstrap.Modal(document.getElementById('resetModal'));
            modal.show();
        })
        .catch(error => {
            console.error('Error:', error);
            alert('เกิดข้อผิดพลาดในการโหลดข้อมูล');
        });
}

// Confirm and reset selected employees
async function confirmReset() {
    const workDate = document.getElementById('workDate').value;
    const selectedCheckboxes = document.querySelectorAll('.reset-checkbox:checked');

    if (selectedCheckboxes.length === 0) {
        alert('กรุณาเลือกพนักงานที่ต้องการลบการบันทึก');
        return;
    }

    const confirmMessage = `ต้องการลบการบันทึกของพนักงาน ${selectedCheckboxes.length} คนที่เลือกใช่หรือไม่?\n\nการดำเนินการนี้ไม่สามารถยกเลิกได้`;

    if (!confirm(confirmMessage)) {
        return;
    }

    let successCount = 0;
    let errorCount = 0;

    try {
        for (const checkbox of selectedCheckboxes) {
            const empId = checkbox.value;
            const empAttendance = attendanceRecords[empId] || { types: [], workplace_substitute: null };
            const attendanceTypes = empAttendance.types || [];

            try {
                // Delete all attendance types for this employee
                for (const type of attendanceTypes) {
                    await deleteAttendanceRecord(workDate, empId, type);
                }
                successCount++;
            } catch (error) {
                console.error(`Error resetting ${empId}:`, error);
                errorCount++;
            }
        }

        // Close modal
        const modal = bootstrap.Modal.getInstance(document.getElementById('resetModal'));
        modal.hide();

        // Show result message
        if (errorCount === 0) {
            showSuccessMessage(`ลบการบันทึกของพนักงาน ${successCount} คนเรียบร้อยแล้ว`);
        } else {
            alert(`ลบการบันทึกสำเร็จ ${successCount} คน, ล้มเหลว ${errorCount} คน`);
        }

        // Reload employees to update button states
        filterEmployees();
    } catch (error) {
        console.error('Error resetting attendance:', error);
        alert('เกิดข้อผิดพลาดในการลบการบันทึก');
    }
}

// ====== Reset ALL Data ======
function openResetAllModal() {
    // Clear previous state
    const pinInput = document.getElementById('resetAllPin');
    const feedback = document.getElementById('resetAllPinFeedback');
    const confirmBtn = document.getElementById('confirmResetAllBtn');
    const eyeIcon = document.getElementById('resetAllPinEye');

    pinInput.value = '';
    pinInput.type = 'password';
    eyeIcon.className = 'bi bi-eye';
    feedback.textContent = '';
    feedback.className = 'form-text mt-1';
    confirmBtn.disabled = true;

    new bootstrap.Modal(document.getElementById('resetAllModal')).show();
}

function checkResetAllPin() {
    const pin = document.getElementById('resetAllPin').value;
    const feedback = document.getElementById('resetAllPinFeedback');
    const confirmBtn = document.getElementById('confirmResetAllBtn');

    if (pin.length === 0) {
        feedback.textContent = '';
        feedback.className = 'form-text mt-1';
        confirmBtn.disabled = true;
    } else if (pin === '47117257') {
        feedback.textContent = '✓ รหัสถูกต้อง';
        feedback.className = 'form-text mt-1 text-success fw-bold';
        confirmBtn.disabled = false;
    } else if (pin.length >= 8) {
        feedback.textContent = '✗ รหัสไม่ถูกต้อง';
        feedback.className = 'form-text mt-1 text-danger';
        confirmBtn.disabled = true;
    } else {
        feedback.textContent = '';
        feedback.className = 'form-text mt-1';
        confirmBtn.disabled = true;
    }
}

function toggleResetAllPin() {
    const input = document.getElementById('resetAllPin');
    const eye = document.getElementById('resetAllPinEye');
    if (input.type === 'password') {
        input.type = 'text';
        eye.className = 'bi bi-eye-slash';
    } else {
        input.type = 'password';
        eye.className = 'bi bi-eye';
    }
}

async function confirmResetAll() {
    const pin = document.getElementById('resetAllPin').value;
    const confirmBtn = document.getElementById('confirmResetAllBtn');

    confirmBtn.disabled = true;
    confirmBtn.innerHTML = '<span class="spinner-border spinner-border-sm me-1"></span>กำลังลบข้อมูล...';

    try {
        const res = await fetch('/api/attendance/reset-all', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ pin })
        });
        const data = await res.json();

        if (data.success) {
            // Close modal
            const modal = bootstrap.Modal.getInstance(document.getElementById('resetAllModal'));
            modal.hide();
            // Refresh employee list
            filterEmployees();
            showSuccessMessage(data.message);
        } else {
            const feedback = document.getElementById('resetAllPinFeedback');
            feedback.textContent = data.message || 'เกิดข้อผิดพลาด';
            feedback.className = 'form-text mt-1 text-danger';
            confirmBtn.disabled = false;
            confirmBtn.innerHTML = '<i class="bi bi-trash3"></i> ยืนยัน ล้างข้อมูลทั้งหมด';
        }
    } catch (err) {
        console.error('reset-all error:', err);
        alert('เกิดข้อผิดพลาดในการเชื่อมต่อ');
        confirmBtn.disabled = false;
        confirmBtn.innerHTML = '<i class="bi bi-trash3"></i> ยืนยัน ล้างข้อมูลทั้งหมด';
    }
}

// Show success message with auto-dismiss
function showSuccessMessage(message) {
    // Create toast container if it doesn't exist
    let toastContainer = document.getElementById('toastContainer');
    if (!toastContainer) {
        toastContainer = document.createElement('div');
        toastContainer.id = 'toastContainer';
        toastContainer.className = 'position-fixed top-0 end-0 p-3';
        toastContainer.style.zIndex = '11';
        document.body.appendChild(toastContainer);
    }

    // Create toast element
    const toastId = 'toast_' + Date.now();
    const toastHTML = `
        <div id="${toastId}" class="toast align-items-center text-white bg-success border-0" role="alert" aria-live="assertive" aria-atomic="true">
            <div class="d-flex">
                <div class="toast-body">
                    <i class="bi bi-check-circle-fill"></i> ${message}
                </div>
                <button type="button" class="btn-close btn-close-white me-2 m-auto" data-bs-dismiss="toast" aria-label="Close"></button>
            </div>
        </div>
    `;

    toastContainer.insertAdjacentHTML('beforeend', toastHTML);

    // Show toast
    const toastElement = document.getElementById(toastId);
    const toast = new bootstrap.Toast(toastElement, { delay: 3000 });
    toast.show();

    // Remove toast element after it's hidden
    toastElement.addEventListener('hidden.bs.toast', function () {
        toastElement.remove();
    });
}

// Show attendance history
function showAttendanceHistory() {
    const workDate = document.getElementById('workDate').value;

    if (!workDate) {
        alert('กรุณาเลือกวันที่');
        return;
    }

    // Format date for display
    const date = new Date(workDate);
    const formattedDate = `${String(date.getDate()).padStart(2, '0')}/${String(date.getMonth() + 1).padStart(2, '0')}/${date.getFullYear()}`;
    document.getElementById('historyDate').textContent = formattedDate;

    // Fetch history data
    fetch(`/api/attendance/history?work_date=${workDate}`)
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                displayAttendanceHistory(data.records, data.summary);

                // Show modal
                const modal = new bootstrap.Modal(document.getElementById('historyModal'));
                modal.show();
            } else {
                alert('เกิดข้อผิดพลาด: ' + data.message);
            }
        })
        .catch(error => {
            console.error('Error:', error);
            alert('เกิดข้อผิดพลาดในการโหลดข้อมูล');
        });
}

// Display attendance history
function displayAttendanceHistory(records, summary) {
    // Update summary counts
    document.getElementById('workCount').textContent = summary.work || 0;
    document.getElementById('leaveCount').textContent = summary.leave || 0;
    document.getElementById('otCount').textContent = summary.ot || 0;

    // Update tab counts
    document.getElementById('workTabCount').textContent = summary.work || 0;
    document.getElementById('leaveTabCount').textContent = summary.leave || 0;
    document.getElementById('otTabCount').textContent = summary.ot || 0;
    document.getElementById('allTabCount').textContent = records.length;

    // Filter records by type
    const workRecords = records.filter(r => r.attendance_type === 'work');
    const leaveRecords = records.filter(r => r.attendance_type === 'leave');
    const otRecords = records.filter(r => r.attendance_type === 'ot');

    // Display each type
    displayHistoryTable('workTableBody', workRecords);
    displayHistoryTable('leaveTableBody', leaveRecords);
    displayHistoryTable('otTableBody', otRecords);
    displayHistoryTableAll('allTableBody', records);
}

// Display history table for specific type
function displayHistoryTable(tableBodyId, records) {
    const tbody = document.getElementById(tableBodyId);

    if (records.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="7" class="text-center text-muted">ไม่มีข้อมูล</td>
            </tr>
        `;
        return;
    }

    tbody.innerHTML = '';
    records.forEach((record, index) => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td class="text-center">${index + 1}</td>
            <td>${record.emp_id}</td>
            <td>${record.name}</td>
            <td>${record.workplace || '-'}</td>
            <td>${record.group_name || '-'}</td>
            <td><span class="badge bg-${record.emp_type === 'MCP' ? 'success' : 'info'}">${record.emp_type}</span></td>
            <td class="text-center"><span class="badge bg-secondary">${record.shift}</span></td>
        `;
        tbody.appendChild(row);
    });
}

// Display all records table with attendance type column
function displayHistoryTableAll(tableBodyId, records) {
    const tbody = document.getElementById(tableBodyId);

    if (records.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="8" class="text-center text-muted">ไม่มีข้อมูล</td>
            </tr>
        `;
        return;
    }

    tbody.innerHTML = '';
    records.forEach((record, index) => {
        const attendanceTypeBadge = getAttendanceTypeBadge(record.attendance_type);
        const row = document.createElement('tr');
        row.innerHTML = `
            <td class="text-center">${index + 1}</td>
            <td>${record.emp_id}</td>
            <td>${record.name}</td>
            <td>${record.workplace || '-'}</td>
            <td>${record.group_name || '-'}</td>
            <td><span class="badge bg-${record.emp_type === 'MCP' ? 'success' : 'info'}">${record.emp_type}</span></td>
            <td class="text-center"><span class="badge bg-secondary">${record.shift}</span></td>
            <td class="text-center">${attendanceTypeBadge}</td>
        `;
        tbody.appendChild(row);
    });
}

// Get attendance type badge
function getAttendanceTypeBadge(attendanceType) {
    if (attendanceType === 'work') {
        return '<span class="badge bg-success"><i class="bi bi-check-circle"></i> ทำงาน</span>';
    } else if (attendanceType === 'leave') {
        return '<span class="badge bg-warning"><i class="bi bi-calendar-x"></i> ลางาน</span>';
    } else if (attendanceType === 'ot') {
        return '<span class="badge bg-primary"><i class="bi bi-clock"></i> OT</span>';
    }
    return '-';
}

// Show Leave Details Modal
async function showLeaveDetailsModal() {
    const workDate = document.getElementById('workDate').value;

    if (!workDate) {
        alert('กรุณาเลือกวันที่ทำงาน');
        return;
    }

    try {
        const response = await fetch(`/api/attendance/leave-details?work_date=${workDate}`);
        const data = await response.json();

        if (!data.success) {
            alert('เกิดข้อผิดพลาด: ' + data.message);
            return;
        }

        // Format date for display
        const dateObj = new Date(workDate);
        const formattedDate = dateObj.toLocaleDateString('th-TH', {
            year: 'numeric',
            month: 'long',
            day: 'numeric',
            weekday: 'long'
        });

        // Update modal summary
        document.getElementById('leaveDetailsDate').textContent = formattedDate;
        document.getElementById('leaveDetailsCount').textContent = `${data.count} คน`;

        // Populate table
        const tbody = document.getElementById('leaveDetailsTableBody');

        if (data.records.length === 0) {
            tbody.innerHTML = `
                <tr>
                    <td colspan="6" style="text-align: center; padding: 40px; color: var(--muted);">
                        <i class="bi bi-inbox" style="font-size: 2rem; display: block; margin-bottom: 10px;"></i>
                        ไม่มีข้อมูลพนักงานที่ลางาน
                    </td>
                </tr>
            `;
        } else {
            tbody.innerHTML = data.records.map((record, index) => `
                <tr>
                    <td class="text-center">${index + 1}</td>
                    <td>${record.emp_id}</td>
                    <td>${record.emp_name}</td>
                    <td>${record.workplace}</td>
                    <td>${record.leave_period}</td>
                    <td>${record.leave_time}</td>
                </tr>
            `).join('');
        }

        // Show modal
        const modal = new bootstrap.Modal(document.getElementById('leaveDetailsModal'));
        modal.show();
    } catch (error) {
        console.error('Error fetching leave details:', error);
        alert('เกิดข้อผิดพลาดในการดึงข้อมูล');
    }
}

window.showLeaveOptionModal = showLeaveOptionModal;
window.confirmLeaveOption = confirmLeaveOption;
window.getLeaveOptionsForShift = getLeaveOptionsForShift;
window.showLeaveDetailsModal = showLeaveDetailsModal;
window.toggleOtherReasonInput = toggleOtherReasonInput;
