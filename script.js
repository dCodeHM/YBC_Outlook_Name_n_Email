document.getElementById('fileInput').addEventListener('change', handleFile, false);
const downloadBtn = document.getElementById('downloadBtn');
const resultTable = document.getElementById('resultTable').getElementsByTagName('tbody')[0];
let parsedData = [];

function handleFile(event) {
    const file = event.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function(e) {
        const text = e.target.result;
        parsedData = parseGraduateNames(text);
        renderTable(parsedData);
        downloadBtn.disabled = parsedData.length === 0;
    };
    reader.readAsText(file);
}

function parseGraduateNames(text) {
    const entries = text.split(';').map(e => e.trim()).filter(Boolean);
    const data = [];
    for (const entry of entries) {
        const nameMatch = entry.match(/^(.*?)\s*<.*?>$/);
        const emailMatch = entry.match(/<([^>]+)>/);
        if (nameMatch && emailMatch) {
            data.push({ name: nameMatch[1].trim(), email: emailMatch[1].trim() });
        }
    }
    return data;
}

function renderTable(data) {
    resultTable.innerHTML = '';
    for (const row of data) {
        const tr = document.createElement('tr');
        const tdName = document.createElement('td');
        tdName.textContent = row.name;
        const tdEmail = document.createElement('td');
        tdEmail.textContent = row.email;
        tr.appendChild(tdName);
        tr.appendChild(tdEmail);
        resultTable.appendChild(tr);
    }
}

downloadBtn.addEventListener('click', function() {
    if (parsedData.length === 0) return;
    const wsData = [
        ['Full Name', 'Email'],
        ...parsedData.map(row => [row.name, row.email])
    ];
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Graduates');
    XLSX.writeFile(wb, 'Graduate_Names.xlsx');
}); 