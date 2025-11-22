// Gerekli Kütüphaneler (Google API'leri artık YOK!)
require('dotenv').config();

const express = require('express');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path'); // Dosya yollarını yönetmek için

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware: JSON ve URL kodlu form verilerini işlemek için
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// --- API Rotası ---
app.post('/create-upload-file', async (req, res) => {
    // 1. Frontend'den (Shopify) gelen verileri alma
    const { pazaryeri, kategori, barkod_on_ek, marka_adi } = req.body;

    // Temel doğrulama
    if (!pazaryeri || !kategori || !barkod_on_ek || !marka_adi) {
        return res.status(400).send('Lütfen tüm zorunlu alanları doldurun.');
    }

    // 2. Şablon Dosyanın Yerel Yolunu Belirleme
    try {
        // Dinamik olarak şablon dosyasının tam yolunu oluştur
        // Örn: templates/trendyol/elbise.xlsx
        const templateFileName = `${kategori}.xlsx`;
        const templatePath = path.join(
            __dirname, // Projenin ana dizini
            'templates', // templates klasörü
            pazaryeri, // trendyol veya hepsiburada (Postman'den gelen değer)
            templateFileName
        );
        
        // Dosyanın gerçekten var olup olmadığını kontrol et
        // Buradaki pazaryeri ve kategori değerlerinin tam olarak klasör ve dosya adlarıyla eşleştiğinden emin olun.
        if (!fs.existsSync(templatePath)) {
             return res.status(404).send(`Hata: '${templateFileName}' şablonu, '${pazaryeri}' klasöründe bulunamadı.`);
        }

        // 3. Şablon Dosyasını Okuma (Buffer'a yüklenir)
        const fileBuffer = fs.readFileSync(templatePath);
        
        // 4. Excel Dosyasını Manipüle Etme
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(fileBuffer);
        
        const worksheet = workbook.worksheets[0]; // İlk çalışma sayfasını al

        // Şablonunuzdaki eski barkod ön ekini tanımlayın
        const eskiOnEk = "ZDX"; 
        
        // Yeni ön eki tırnak içine alarak hazırla (formül için gerekli)
        const yeniOnEkTirnakli = `"${barkod_on_ek}"`;
        const eskiOnEkTirnakli = `"${eskiOnEk}"`;


        // Dosyadaki her satırı döngüye al
        worksheet.eachRow({ includeEmpty: false, first: 2 }, (row, rowNumber) => {
            
            // A Sütunu: SKU/Barcode için formül değiştirme (="ZDX" & Kx & Yx)
            const cellA = row.getCell('A');
            if (cellA.formula) {
                let newFormula = cellA.formula.replace(eskiOnEkTirnakli, yeniOnEkTirnakli);
                cellA.value = { formula: newFormula };
            }

            // B Sütunu: Barkodlar için formül değiştirme (="ZDX" & Kx)
            const cellB = row.getCell('B');
            if (cellB.formula) {
                let newFormula = cellB.formula.replace(eskiOnEkTirnakli, yeniOnEkTirnakli);
                cellB.value = { formula: newFormula };
            }
            
            // 💡 Marka Adı Güncelleme (C sütununda Marka Adı olduğunu varsayalım)
            const cellC = row.getCell('C'); 
            if (!cellC.value || cellC.value !== marka_adi) {
                cellC.value = marka_adi; 
            }
        });

        // 5. Değiştirilmiş Dosyayı Buffer Olarak Kaydetme
        const modifiedBuffer = await workbook.xlsx.writeBuffer();

        // 6. Kullanıcıya Geri Gönderme (İndirme Başlatma)
        const outputFileName = `${pazaryeri}-${kategori}-${barkod_on_ek}-${Date.now()}.xlsx`;
        
        // Yanıt başlıklarını ayarlama (Dosya indirme başlatmak için kritik)
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${outputFileName}"`);
        
        // Buffer'ı yanıt olarak gönderme
        res.send(modifiedBuffer);

    } catch (error) {
        console.error('İşlem sırasında beklenmeyen hata:', error);
        res.status(500).send(`Dosya işlenirken sunucu hatası oluştu: ${error.message}`);
    }
});

// Sunucuyu başlatma
app.listen(PORT, () => {
  console.log(`Sunucu http://localhost:${PORT} adresinde çalışıyor.`);
});