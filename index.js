// Gerekli Kütüphaneler (Google API'leri artık YOK!)
require('dotenv').config();

const express = require('express');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const cors = require('cors'); // ✨ YENİ: CORS paketini ekledik

const app = express();
const PORT = process.env.PORT || 3000;

// 1. ✨ YENİ: CORS Middleware'i
// Tüm alan adlarından gelen isteklere izin veriyoruz.
// Production ortamında sadece Shopify alan adınızı buraya eklemeniz daha güvenlidir.
app.use(cors()); 

// 2. Middleware: JSON ve URL kodlu form verilerini işlemek için
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
        const templateFileName = `${kategori}.xlsx`;
        const templatePath = path.join(
            __dirname,
            'templates', 
            pazaryeri, 
            templateFileName
        );
        
        // Dosyanın gerçekten var olup olmadığını kontrol et
        if (!fs.existsSync(templatePath)) {
             return res.status(404).send(`Hata: '${templateFileName}' şablonu, '${pazaryeri}' klasöründe bulunamadı.`);
        }

        // 3. Şablon Dosyasını Okuma (Buffer'a yüklenir)
        const fileBuffer = fs.readFileSync(templatePath);
        
        // 4. Excel Dosyasını Manipüle Etme
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(fileBuffer);
        
        const worksheet = workbook.worksheets[0]; // İlk çalışma sayfasını al

        const eskiOnEk = "ZDX"; 
        const yeniOnEkTirnakli = `"${barkod_on_ek}"`;
        const eskiOnEkTirnakli = `"${eskiOnEk}"`;


        // Dosyadaki her satırı döngüye al
        worksheet.eachRow({ includeEmpty: false}, (row, rowNumber) => {
            if (rowNumber === 1) {
        return; 
    }
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
            
            // Marka Adı Güncelleme (C sütununda Marka Adı olduğunu varsayalım)
            // C Sütunu: Marka Adı Güncelleme (Sadece veri satırlarında)
    const cellC = row.getCell('C'); 
    if (cellC) {
        // Kullanıcının girdiği marka adını hücreye yazar
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


// ----------------------------------------------------------------------
// ✅ YENİ ROTA: /stok-guncelle (Toplu Stok Güncelleme)
// ----------------------------------------------------------------------
app.post('/stok-guncelle', async (req, res) => {
    // 1. Frontend'den gelen verileri alma (Sadece barkod_on_ek)
    const { barkod_on_ek } = req.body;

    // Temel doğrulama
    if (!barkod_on_ek) {
        return res.status(400).send('Barkod Ön Ek alanı boş bırakılamaz.');
    }

    // 2. Şablon Dosyanın Yerel Yolunu Belirleme (stok/stok_guncelleme.xlsx)
    try {
        const templateFileName = 'stok_guncelleme.xlsx';
        const templatePath = path.join(
            __dirname,
            'stok', // Yeni klasör adı (stok)
            templateFileName
        );
        
        if (!fs.existsSync(templatePath)) {
             return res.status(404).send(`Hata: Stok şablon dosyası ('${templateFileName}') bulunamadı. Lütfen /stok klasörüne eklediğinizden emin olun.`);
        }

        // 3. Şablon Dosyasını Oku
        const fileBuffer = fs.readFileSync(templatePath);
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(fileBuffer);
        const worksheet = workbook.worksheets[0]; 

        // 4. Excel Dosyasını Manipüle Etme
        const eskiOnEk = "ZDX"; // Şablonunuzdaki eski ön ek
        const yeniOnEkTirnakli = `"${barkod_on_ek}"`;
        const eskiOnEkTirnakli = `"${eskiOnEk}"`;

        // Yalnızca formül içeren A ve B sütunlarını güncellemek için döngü
        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
            
            // SADECE SATIR 1'İ ATLA (Başlıkları korumak için)
            if (rowNumber === 1) {
                return; 
            }
            
            // A Sütunu: SKU/Barcode formülü değiştirilir
            const cellA = row.getCell('A');
            if (cellA.formula) {
                let newFormula = cellA.formula.replace(eskiOnEkTirnakli, yeniOnEkTirnakli);
                cellA.value = { formula: newFormula };
            }

            // B Sütunu: Barkodlar için formül değiştirilir
            const cellB = row.getCell('B');
            if (cellB.formula) {
                let newFormula = cellB.formula.replace(eskiOnEkTirnakli, yeniOnEkTirnakli);
                cellB.value = { formula: newFormula };
            }
            
            // C Sütunu veya diğer sütunlar için hiçbir işlem YAPILMAZ (talep üzerine)
        });

        // 5. Geri Gönderme
        const modifiedBuffer = await workbook.xlsx.writeBuffer();
        const outputFileName = `stok_guncelleme_${barkod_on_ek}_${Date.now()}.xlsx`;
        
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${outputFileName}"`);
        res.send(modifiedBuffer);

    } catch (error) {
        console.error('Stok güncelleme sırasında hata:', error);
        res.status(500).send(`Stok dosyası işlenirken sunucu hatası oluştu: ${error.message}`);
    }
});

// Sunucuyu başlatma (Vercel kendi portunu atayacağı için burası production'da çalışmaz, ama yerel test için gerekli)
app.listen(PORT, () => {
  console.log(`Sunucu http://localhost:${PORT} adresinde çalışıyor.`);
});