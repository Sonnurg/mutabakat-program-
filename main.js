const { jsPDF } = window.jspdf;
let excelData = [];

// Excel dosyasını yükleme
document.getElementById('fileInput').addEventListener('change', handleFile);

function handleFile(event) {
    const file = event.target.files[0];
    if (!file) return;
    if (!file.name.match(/\.(xlsx|xls)$/)) {
        alert("Lütfen geçerli bir Excel dosyası yükleyin");
        return;
    }

    const reader = new FileReader();
    reader.onload = e => {
        try {
            const workbook = XLSX.read(e.target.result, { type: 'binary' });
            const sheetName = workbook.SheetNames[0];
            const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });
            if (jsonData.length < 2) {
                alert("Excel dosyasında yeterli veri yok");
                return;
            }
            processExcelData(jsonData);
        } catch (err) {
            alert("Dosya okunamadı");
        }
    };
    reader.readAsBinaryString(file);
}

function processExcelData(jsonData) {
    excelData = jsonData.slice(1).map(row => ({
        cariKod: row[0] || '',
        unvan: row[1] || '',
        adres: row[2] || '',
        adres2: row[3] || '',
        ilce: row[4] || '',
        il: row[5] || '',
        vergiDairesi: row[7] || '',
        vergiNo: row[8] || '',
        borc: row[10] || 0
    })).filter(r => r.cariKod);

    // İlk satırı test için PDF yapalım
    if (excelData.length > 0) {
        generatePDF(excelData[0]);
    }
}

function generatePDF(data) {
    // HTML şablonunu doldur
    document.getElementById('tpl-cariKod').textContent = data.cariKod;
    document.getElementById('tpl-unvan').textContent = data.unvan;
    document.getElementById('tpl-vergiDairesi').textContent = data.vergiDairesi;
    document.getElementById('tpl-vergiNo').textContent = data.vergiNo;
    document.getElementById('tpl-adres').textContent = `${data.adres} ${data.adres2}`;
    document.getElementById('tpl-ilce').textContent = data.ilce;
    document.getElementById('tpl-il').textContent = data.il;
    document.getElementById('tpl-unvan2').textContent = data.unvan;
    document.getElementById('tpl-tarih').textContent = new Date().toLocaleDateString('tr-TR');
    document.getElementById('tpl-borc').textContent = formatMoney(data.borc);
    document.getElementById('tpl-borcYazi').textContent = numberToWords(data.borc);

    const doc = new jsPDF('p', 'mm', 'a4');
    doc.html(document.getElementById('mutabakat-template'), {
        callback: function (doc) {
            doc.save(`${data.cariKod}_mutabakat.pdf`);
        },
        x: 10,
        y: 10,
        html2canvas: { scale: 0.5 }
    });
}

// Formatlama fonksiyonları
function formatMoney(amount) {
    if (!amount || isNaN(amount)) return '0,00';
    return parseFloat(amount).toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&.').replace('.', ',');
}

function numberToWords(amount) {
    if (!amount || isNaN(amount)) return 'Sıfır TL';
    const num = parseFloat(amount), lira = Math.floor(num), kurus = Math.round((num - lira) * 100);
    const ones = ['', 'bir', 'iki', 'üç', 'dört', 'beş', 'altı', 'yedi', 'sekiz', 'dokuz'];
    const tens = ['', '', 'yirmi', 'otuz', 'kırk', 'elli', 'altmış', 'yetmiş', 'seksen', 'doksan'];
    let res = '';
    if (lira >= 1000) res += (Math.floor(lira / 1000) === 1 ? 'bin' : ones[Math.floor(lira / 1000)] + ' bin');
    const rem = lira % 1000;
    if (rem >= 100) res += (Math.floor(rem / 100) === 1 ? 'yüz' : ones[Math.floor(rem / 100)] + ' yüz');
    const lt = rem % 100;
    if (lt >= 10) { res += ' ' + tens[Math.floor(lt / 10)]; if (lt % 10 > 0) res += ' ' + ones[lt % 10]; }
    else if (lt > 0) res += ' ' + ones[lt];
    res += ' TL';
    if (kurus > 0) res += ' ' + (kurus < 10 ? ones[kurus] : tens[Math.floor(kurus / 10)] + (kurus % 10 > 0 ? ' ' + ones[kurus % 10] : '')) + ' kuruş';
    return res.trim();
}
