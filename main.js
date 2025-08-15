// PDF kütüphanesi kontrolü
if (typeof window.jspdf === 'undefined') {
    console.error('jsPDF kütüphanesi yüklenmedi');
} else {
    console.log('jsPDF kütüphanesi yüklendi');
}

// XLSX kütüphanesi kontrolü
if (typeof XLSX === 'undefined') {
    console.error('XLSX kütüphanesi yüklenmedi');
} else {
    console.log('XLSX kütüphanesi yüklendi');
}

const { jsPDF } = window.jspdf;
let excelData = [];

// Sayfa yüklendiğinde çalışacak
document.addEventListener('DOMContentLoaded', function() {
    console.log('Sayfa yüklendi');
    
    // File input kontrolü
    const fileInput = document.getElementById('fileInput');
    if (!fileInput) {
        console.error('fileInput elementi bulunamadı');
        return;
    }
    
    fileInput.addEventListener('change', handleFile);
    console.log('Event listener eklendi');
});

function handleFile(event) {
    console.log('Dosya seçildi');
    const file = event.target.files[0];
    
    if (!file) {
        console.log('Dosya seçilmedi');
        return;
    }
    
    console.log('Dosya adı:', file.name);
    
    // Dosya uzantısı kontrolü
    if (!file.name.match(/\.(xlsx|xls)$/i)) {
        showMessage("Lütfen geçerli bir Excel dosyası yükleyin (.xlsx veya .xls)", 'error');
        return;
    }

    showMessage("Dosya okunuyor...", 'success');
    
    const reader = new FileReader();
    reader.onload = function(e) {
        console.log('Dosya okundu');
        try {
            const workbook = XLSX.read(e.target.result, { type: 'binary' });
            console.log('Workbook oluşturuldu:', workbook.SheetNames);
            
            const sheetName = workbook.SheetNames[0];
            const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });
            
            console.log('JSON data:', jsonData);
            
            if (jsonData.length < 2) {
                showMessage("Excel dosyasında yeterli veri yok", 'error');
                return;
            }
            
            processExcelData(jsonData);
        } catch (err) {
            console.error('Dosya okuma hatası:', err);
            showMessage("Dosya okunamadı: " + err.message, 'error');
        }
    };
    
    reader.onerror = function(err) {
        console.error('FileReader hatası:', err);
        showMessage("Dosya okuma hatası", 'error');
    };
    
    reader.readAsBinaryString(file);
}

function processExcelData(jsonData) {
    console.log('Excel verisi işleniyor');
    
    // İlk satır başlık olarak kabul ediliyor
    const headers = jsonData[0];
    console.log('Başlıklar:', headers);
    
    excelData = jsonData.slice(1).map((row, index) => {
        const data = {
            cariKod: row[0] || '',
            unvan: row[1] || '',
            adres: row[2] || '',
            adres2: row[3] || '',
            ilce: row[4] || '',
            il: row[5] || '',
            vergiDairesi: row[7] || '',
            vergiNo: row[8] || '',
            borc: parseFloat(row[10]) || 0
        };
        console.log(`Satır ${index + 1}:`, data);
        return data;
    }).filter(r => r.cariKod && r.cariKod.trim() !== '');

    console.log('İşlenen veri sayısı:', excelData.length);
    
    if (excelData.length === 0) {
        showMessage("Geçerli veri bulunamadı", 'error');
        return;
    }
    
    showMessage(`${excelData.length} kayıt işlendi. İlk kayıt için PDF oluşturuluyor...`, 'success');
    
    // İlk satırı test için PDF yap
    setTimeout(() => {
        generatePDF(excelData[0]);
    }, 1000);
}

function generatePDF(data) {
    console.log('PDF oluşturuluyor:', data);
    
    try {
        // Template kontrolü
        const template = document.getElementById('mutabakat-template');
        if (!template) {
            throw new Error('Template bulunamadı');
        }
        
        // HTML şablonunu doldur
        document.getElementById('tpl-cariKod').textContent = data.cariKod;
        document.getElementById('tpl-unvan').textContent = data.unvan;
        document.getElementById('tpl-vergiDairesi').textContent = data.vergiDairesi;
        document.getElementById('tpl-vergiNo').textContent = data.vergiNo;
        document.getElementById('tpl-adres').textContent = `${data.adres} ${data.adres2}`.trim();
        document.getElementById('tpl-ilce').textContent = data.ilce;
        document.getElementById('tpl-il').textContent = data.il;
        document.getElementById('tpl-unvan2').textContent = data.unvan;
        document.getElementById('tpl-tarih').textContent = new Date().toLocaleDateString('tr-TR');
        document.getElementById('tpl-borc').textContent = formatMoney(data.borc);
        document.getElementById('tpl-borcYazi').textContent = numberToWords(data.borc);

        console.log('Template dolduruldu');
        
        // PDF oluştur
        const doc = new jsPDF('p', 'mm', 'a4');
        
        // Template'i geçici olarak görünür yap
        template.style.position = 'static';
        template.style.left = 'auto';
        
        doc.html(template, {
            callback: function (doc) {
                console.log('PDF hazır');
                doc.save(`${data.cariKod}_mutabakat.pdf`);
                
                // Template'i tekrar gizle
                template.style.position = 'absolute';
                template.style.left = '-9999px';
                
                showMessage(`${data.cariKod} için PDF oluşturuldu`, 'success');
            },
            x: 10,
            y: 10,
            width: 180,
            windowWidth: 800,
            html2canvas: { 
                scale: 0.8,
                useCORS: true,
                letterRendering: true
            }
        });
        
    } catch (err) {
        console.error('PDF oluşturma hatası:', err);
        showMessage("PDF oluşturulamadı: " + err.message, 'error');
    }
}

// Para formatı
function formatMoney(amount) {
    if (!amount || isNaN(amount)) return '0,00 TL';
    return new Intl.NumberFormat('tr-TR', {
        style: 'currency',
        currency: 'TRY',
        minimumFractionDigits: 2
    }).format(amount);
}

// Sayıyı yazıya çevirme
function numberToWords(amount) {
    if (!amount || isNaN(amount)) return 'Sıfır Türk Lirası';

    const num = parseFloat(amount);
    const lira = Math.floor(num);
    const kurus = Math.round(((num - lira) * 100));

    const ones = ['', 'bir', 'iki', 'üç', 'dört', 'beş', 'altı', 'yedi', 'sekiz', 'dokuz'];
    const tens = ['', '', 'yirmi', 'otuz', 'kırk', 'elli', 'altmış', 'yetmiş', 'seksen', 'doksan'];
    const teens = ['on', 'on bir', 'on iki', 'on üç', 'on dört', 'on beş', 'on altı', 'on yedi', 'on sekiz', 'on dokuz'];

    function convertHundreds(n) {
        let result = '';
        
        if (n >= 100) {
            const hundreds = Math.floor(n / 100);
            result += (hundreds === 1 ? 'yüz' : ones[hundreds] + ' yüz');
            n %= 100;
        }
        
        if (n >= 20) {
            result += (result ? ' ' : '') + tens[Math.floor(n / 10)];
            if (n % 10 > 0) result += ' ' + ones[n % 10];
        } else if (n >= 10) {
            result += (result ? ' ' : '') + teens[n - 10];
        } else if (n > 0) {
            result += (result ? ' ' : '') + ones[n];
        }
        
        return result;
    }

    let result = '';

    // Milyonlar
    if (lira >= 1000000) {
        const millions = Math.floor(lira / 1000000);
        result += (millions === 1 ? 'bir milyon' : convertHundreds(millions) + ' milyon');
        lira %= 1000000;
    }

    // Binler
    const remainder = lira % 1000000;
    if (remainder >= 1000) {
        const thousands = Math.floor(remainder / 1000);
        result += (result ? ' ' : '') + (thousands === 1 ? 'bin' : convertHundreds(thousands) + ' bin');
    }

    // Yüzler, onlar, birler
    const lastPart = remainder % 1000;
    if (lastPart > 0) {
        result += (result ? ' ' : '') + convertHundreds(lastPart);
    }

    if (!result) result = 'sıfır';
    
    result += ' Türk Lirası';

    // Kuruş ekle
    if (kurus > 0) {
        result += ' ' + convertHundreds(kurus) + ' kuruş';
    }

    return result;
}

// Mesaj gösterme fonksiyonu
function showMessage(text, type = 'success') {
    console.log('Mesaj:', text, type);
    
    // Önceki mesajları temizle
    const oldMessages = document.querySelectorAll('.status-message');
    oldMessages.forEach(msg => msg.remove());
    
    const messageDiv = document.createElement('div');
    messageDiv.className = `status-message ${type}`;
    messageDiv.textContent = text;
    
    const uploadSection = document.querySelector('.upload-section');
    if (uploadSection) {
        uploadSection.appendChild(messageDiv);
        
        // 5 saniye sonra mesajı kaldır
        setTimeout(() => {
            if (messageDiv.parentNode) {
                messageDiv.remove();
            }
        }, 5000);
    }
}
