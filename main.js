let excelData = [];
const { jsPDF } = window.jspdf;

// Yükleme alanı tıklama ve sürükle-bırak
document.addEventListener("DOMContentLoaded", () => {
    const uploadArea = document.getElementById('uploadArea');
    const fileInput = document.getElementById('fileInput');

    uploadArea.addEventListener('click', () => fileInput.click());

    uploadArea.addEventListener('dragover', e => {
        e.preventDefault();
        uploadArea.classList.add('dragover');
    });

    uploadArea.addEventListener('dragleave', e => {
        uploadArea.classList.remove('dragover');
    });

    uploadArea.addEventListener('drop', e => {
        e.preventDefault();
        uploadArea.classList.remove('dragover');
        if (e.dataTransfer.files.length > 0) handleFile(e.dataTransfer.files[0]);
    });

    fileInput.addEventListener('change', e => {
        if (e.target.files.length > 0) handleFile(e.target.files[0]);
    });
});

// Excel dosyasını okuma
function handleFile(file) {
    if (!file.name.match(/\.(xlsx|xls)$/)) {
        return showStatus('Geçerli bir Excel dosyası seçin', 'error');
    }

    const reader = new FileReader();
    reader.onload = e => {
        try {
            const workbook = XLSX.read(e.target.result, { type: 'binary' });
            const sheetName = workbook.SheetNames[0];
            const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });

            if (jsonData.length < 2) {
                return showStatus('Yeterli veri yok', 'error');
            }

            processExcelData(jsonData);
            showStatus('Dosya yüklendi', 'success');
        } catch (err) {
            console.error(err);
            showStatus('Dosya okunamadı', 'error');
        }
    };
    reader.readAsBinaryString(file);
}

// Excel verisini işleme
function processExcelData(jsonData) {
    excelData = jsonData.slice(1).map(row => ({
        cariKod: row[0] || '',
        unvan: row[1] || '',
        adres: row[2] || '',
        adres2: row[3] || '',
        ilce: row[4] || '',
        il: row[5] || '',
        vergiDairesi: row[6] || '',
        vergiNo: row[7] || '',
        borc: row[8] || 0
    })).filter(r => r.cariKod);

    updateTable();
    document.getElementById('dataSection').style.display = 'block';
    document.getElementById('recordCount').textContent = excelData.length;
}

// Tabloyu güncelleme
function updateTable() {
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = '';
    excelData.forEach((d, i) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${d.cariKod}</td>
            <td>${d.unvan}</td>
            <td>${d.adres} ${d.adres2}</td>
            <td>${d.ilce}</td>
            <td>${d.il}</td>
            <td>${formatMoney(d.borc)} TL</td>
            <td><button onclick="generateSinglePDF(${i})">📄 PDF İndir</button></td>
        `;
        tbody.appendChild(tr);
    });
}

// Para formatı
function formatMoney(amount) {
    if (!amount || isNaN(amount)) return '0,00';
    return parseFloat(amount).toFixed(2).replace('.', ',');
}

// Rakamı yazıya çevirme (kısa)
function numberToWords(amount) {
    if (!amount || isNaN(amount)) return 'Sıfır TL';
    const num = parseFloat(amount);
    const lira = Math.floor(num);
    const kurus = Math.round((num - lira) * 100);
    return `${lira} TL ${kurus > 0 ? kurus + ' kuruş' : ''}`;
}

// Tekli PDF
function generateSinglePDF(index) {
    const doc = createPDF(excelData[index]);
    doc.save(`${excelData[index].cariKod}_mutabakat.pdf`);
}

// Toplu PDF ZIP
async function downloadAllPDFs() {
    const zip = new JSZip();
    for (let i = 0; i < excelData.length; i++) {
        const pdfBlob = createPDF(excelData[i]).output('blob');
        zip.file(`${excelData[i].cariKod}_mutabakat.pdf`, pdfBlob);
    }
    const blob = await zip.generateAsync({ type: 'blob' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'mutabakatlar.zip';
    a.click();
    URL.revokeObjectURL(url);
}

// PDF oluşturma
function createPDF(data) {
    const doc = new jsPDF();
    doc.setFontSize(12);
    doc.text(`Cari Kodu: ${data.cariKod}`, 10, 10);
    doc.text(`Unvan: ${data.unvan}`, 10, 20);
    doc.text(`Vergi Dairesi: ${data.vergiDairesi}`, 10, 30);
    doc.text(`Vergi No: ${data.vergiNo}`, 10, 40);
    doc.text(`Adres: ${data.adres} ${data.adres2}`, 10, 50);
    doc.text(`İlçe: ${data.ilce}`, 10, 60);
    doc.text(`İl: ${data.il}`, 10, 70);
    doc.text(`Borcunuz: ${formatMoney(data.borc)} TL (${numberToWords(data.borc)})`, 10, 90);
    return doc;
}

// Durum mesajı
function showStatus(msg, type) {
    const el = document.getElementById('statusMessage');
    el.textContent = msg;
    el.className = 'status-message status-' + type;
    el.style.display = 'block';
    if (type === 'success') setTimeout(() => el.style.display = 'none', 3000);
}
