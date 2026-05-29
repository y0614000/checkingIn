// 考勤状态配置。value 用于程序存储，label 用于页面显示和 Excel 导出。
const STATUS_OPTIONS = [
  { value: "normal", label: "正常" },
  { value: "late", label: "迟到" },
  { value: "personalLeave", label: "事假" },
  { value: "sickLeave", label: "病假" },
  { value: "absent", label: "缺勤" }
];

const UNCHECKED_STATUS = "unchecked";
const STORAGE_KEY = "checkinToolStudents";
const DEFAULT_FIELD_LABELS = {
  identifier: "学号",
  name: "姓名"
};

const fileInput = document.getElementById("excelFile");
const exportBtn = document.getElementById("exportBtn");
const clearBtn = document.getElementById("clearBtn");
const resetBtn = document.getElementById("resetBtn");
const searchInput = document.getElementById("searchInput");
const studentList = document.getElementById("studentList");
const emptyTip = document.getElementById("emptyTip");

const totalCount = document.getElementById("totalCount");
const normalCount = document.getElementById("normalCount");
const lateCount = document.getElementById("lateCount");
const personalLeaveCount = document.getElementById("personalLeaveCount");
const sickLeaveCount = document.getElementById("sickLeaveCount");
const absentCount = document.getElementById("absentCount");
const uncheckedCount = document.getElementById("uncheckedCount");

let students = [];
let fieldLabels = { ...DEFAULT_FIELD_LABELS };

// 页面加载时先恢复 localStorage 中的考勤数据。
document.addEventListener("DOMContentLoaded", () => {
  const savedData = loadSavedData();
  students = savedData.students;
  fieldLabels = savedData.fieldLabels;
  render();
});

fileInput.addEventListener("change", handleImport);
exportBtn.addEventListener("click", exportResult);
clearBtn.addEventListener("click", clearRecords);
resetBtn.addEventListener("click", resetAllToUnchecked);
searchInput.addEventListener("input", renderStudentList);

function loadSavedData() {
  const saved = localStorage.getItem(STORAGE_KEY);

  if (!saved) {
    return {
      students: [],
      fieldLabels: { ...DEFAULT_FIELD_LABELS }
    };
  }

  try {
    const parsed = JSON.parse(saved);

    // 兼容旧版本：旧数据直接保存为学生数组。
    if (Array.isArray(parsed)) {
      return {
        students: parsed,
        fieldLabels: { ...DEFAULT_FIELD_LABELS }
      };
    }

    return {
      students: Array.isArray(parsed.students) ? parsed.students : [],
      fieldLabels: {
        identifier: parsed.fieldLabels?.identifier || DEFAULT_FIELD_LABELS.identifier,
        name: parsed.fieldLabels?.name || DEFAULT_FIELD_LABELS.name
      }
    };
  } catch (error) {
    console.error("读取本地考勤数据失败：", error);
    return {
      students: [],
      fieldLabels: { ...DEFAULT_FIELD_LABELS }
    };
  }
}

function saveStudents() {
  localStorage.setItem(
    STORAGE_KEY,
    JSON.stringify({
      students,
      fieldLabels
    })
  );
}

function handleImport(event) {
  const file = event.target.files[0];

  if (!file) {
    return;
  }

  if (!window.XLSX) {
    alert("Excel 解析库未加载。请确认 lib/xlsx.full.min.js 存在，或网络可以访问 CDN。");
    event.target.value = "";
    return;
  }

  const reader = new FileReader();

  reader.onload = (loadEvent) => {
    try {
      const data = new Uint8Array(loadEvent.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const firstSheetName = workbook.SheetNames[0];

      if (!firstSheetName) {
        alert("Excel 文件中没有工作表。");
        return;
      }

      const worksheet = workbook.Sheets[firstSheetName];
      const rows = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        defval: "",
        raw: false
      });

      const importResult = parseStudentRows(rows);

      if (importResult.students.length === 0) {
        alert("没有读取到学生数据。请确认 Excel 第一行为表头，后面是学生数据。");
        return;
      }

      students = importResult.students;
      fieldLabels = importResult.fieldLabels;
      saveStudents();
      searchInput.value = "";
      render();
      alert(`成功导入 ${students.length} 名学生。`);
    } catch (error) {
      console.error("导入 Excel 失败：", error);
      alert("导入失败，请确认文件是有效的 Excel。");
    } finally {
      // 清空 input 值，方便再次选择同一个文件导入。
      event.target.value = "";
    }
  };

  reader.readAsArrayBuffer(file);
}

function parseStudentRows(rows) {
  const usefulRows = rows.filter((row) => {
    return row.some((cell) => normalizeCell(cell));
  });

  if (usefulRows.length < 2) {
    return {
      students: [],
      fieldLabels: { ...DEFAULT_FIELD_LABELS }
    };
  }

  // 第一行作为动态表头，支持“学号/姓名”“序号/姓名”等不同叫法。
  const headerRow = usefulRows[0].map(normalizeCell);
  const columnConfig = getColumnConfig(headerRow);
  const dataRows = usefulRows.slice(1);

  const importedStudents = dataRows
    .map((row, index) => {
      const studentId = normalizeCell(row[columnConfig.identifierIndex]);
      const name = normalizeCell(row[columnConfig.nameIndex]);

      if (!studentId && !name) {
        return null;
      }

      return {
        id: `${Date.now()}-${index}-${studentId || name}`,
        studentId,
        name,
        status: UNCHECKED_STATUS,
        time: ""
      };
    })
    .filter(Boolean);

  return {
    students: importedStudents,
    fieldLabels: {
      identifier: columnConfig.identifierLabel,
      name: columnConfig.nameLabel
    }
  };
}

function getColumnConfig(headerRow) {
  const nameIndex = findHeaderIndex(headerRow, ["姓名", "名字", "名称", "name"]);
  const resolvedNameIndex = nameIndex >= 0 ? nameIndex : 1;
  const identifierIndex = headerRow.findIndex((header, index) => {
    return index !== resolvedNameIndex && header;
  });
  const resolvedIdentifierIndex = identifierIndex >= 0 ? identifierIndex : 0;

  return {
    identifierIndex: resolvedIdentifierIndex,
    nameIndex: resolvedNameIndex,
    identifierLabel: headerRow[resolvedIdentifierIndex] || DEFAULT_FIELD_LABELS.identifier,
    nameLabel: headerRow[resolvedNameIndex] || DEFAULT_FIELD_LABELS.name
  };
}

function findHeaderIndex(headerRow, keywords) {
  return headerRow.findIndex((header) => {
    const normalizedHeader = header.toLowerCase();
    return keywords.some((keyword) => normalizedHeader.includes(keyword.toLowerCase()));
  });
}

function normalizeCell(value) {
  return String(value ?? "").trim();
}

function render() {
  updateSearchPlaceholder();
  renderStats();
  renderStudentList();
}

function updateSearchPlaceholder() {
  searchInput.placeholder = `搜索${fieldLabels.name}或${fieldLabels.identifier}`;
}

function renderStats() {
  totalCount.textContent = students.length;
  normalCount.textContent = countByStatus("normal");
  lateCount.textContent = countByStatus("late");
  personalLeaveCount.textContent = countByStatus("personalLeave");
  sickLeaveCount.textContent = countByStatus("sickLeave");
  absentCount.textContent = countByStatus("absent");
  uncheckedCount.textContent = students.filter((student) => {
    return student.status === UNCHECKED_STATUS;
  }).length;
}

function countByStatus(status) {
  return students.filter((student) => student.status === status).length;
}

function renderStudentList() {
  const keyword = searchInput.value.trim().toLowerCase();
  const filteredStudents = students.filter((student) => {
    const studentId = String(student.studentId || "").toLowerCase();
    const name = String(student.name || "").toLowerCase();
    return studentId.includes(keyword) || name.includes(keyword);
  });

  studentList.innerHTML = "";
  emptyTip.style.display = filteredStudents.length === 0 ? "block" : "none";
  emptyTip.textContent = students.length === 0 ? "请先导入 Excel 学生名单" : "没有匹配的学生";

  const fragment = document.createDocumentFragment();

  filteredStudents.forEach((student) => {
    fragment.appendChild(createStudentCard(student));
  });

  studentList.appendChild(fragment);
}

function createStudentCard(student) {
  const card = document.createElement("article");
  card.className = "student-card";

  const info = document.createElement("div");
  info.className = "student-info";

  const textWrap = document.createElement("div");

  const name = document.createElement("h2");
  name.className = "student-name";
  name.textContent = student.name || "未填写姓名";

  const studentId = document.createElement("p");
  studentId.className = "student-id";
  studentId.textContent = `${fieldLabels.identifier}：${student.studentId || "未填写"}`;

  const badge = document.createElement("span");
  badge.className = `status-badge status-${student.status}`;
  badge.textContent = getStatusLabel(student.status);

  textWrap.append(name, studentId);
  info.append(textWrap, badge);

  const buttons = document.createElement("div");
  buttons.className = "status-buttons";

  STATUS_OPTIONS.forEach((status) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "status-btn";
    button.textContent = status.label;

    if (student.status === status.value) {
      button.classList.add("active");
    }

    button.addEventListener("click", () => {
      updateStudentStatus(student.id, status.value);
    });

    buttons.appendChild(button);
  });

  card.append(info, buttons);
  return card;
}

function updateStudentStatus(studentId, status) {
  students = students.map((student) => {
    if (student.id !== studentId) {
      return student;
    }

    return {
      ...student,
      status,
      time: formatDateTime(new Date())
    };
  });

  saveStudents();
  render();
}

function resetAllToUnchecked() {
  if (students.length === 0) {
    alert("当前没有学生名单。");
    return;
  }

  const confirmed = confirm("确定要将所有学生恢复为“未签到”吗？");

  if (!confirmed) {
    return;
  }

  students = students.map((student) => ({
    ...student,
    status: UNCHECKED_STATUS,
    time: ""
  }));

  saveStudents();
  render();
}

function clearRecords() {
  if (students.length === 0) {
    alert("当前没有可清空的记录。");
    return;
  }

  const confirmed = confirm("确定要清空本次考勤记录和学生名单吗？");

  if (!confirmed) {
    return;
  }

  students = [];
  localStorage.removeItem(STORAGE_KEY);
  searchInput.value = "";
  render();
}

function exportResult() {
  if (students.length === 0) {
    alert("当前没有可导出的考勤结果。");
    return;
  }

  if (!window.XLSX) {
    alert("Excel 导出库未加载。请确认 lib/xlsx.full.min.js 存在，或网络可以访问 CDN。");
    return;
  }

  const rows = students.map((student) => ({
    [fieldLabels.identifier]: student.studentId,
    [fieldLabels.name]: student.name,
    签到状态: getStatusLabel(student.status),
    操作时间: student.time
  }));

  const worksheet = XLSX.utils.json_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "考勤结果");
  XLSX.writeFile(workbook, `考勤结果_${formatDateForFile(new Date())}.xlsx`);
}

function getStatusLabel(status) {
  if (status === UNCHECKED_STATUS) {
    return "未签到";
  }

  const matched = STATUS_OPTIONS.find((option) => option.value === status);
  return matched ? matched.label : "未签到";
}

function formatDateTime(date) {
  const year = date.getFullYear();
  const month = padZero(date.getMonth() + 1);
  const day = padZero(date.getDate());
  const hour = padZero(date.getHours());
  const minute = padZero(date.getMinutes());
  const second = padZero(date.getSeconds());

  return `${year}-${month}-${day} ${hour}:${minute}:${second}`;
}

function formatDateForFile(date) {
  const year = date.getFullYear();
  const month = padZero(date.getMonth() + 1);
  const day = padZero(date.getDate());

  return `${year}${month}${day}`;
}

function padZero(number) {
  return String(number).padStart(2, "0");
}
