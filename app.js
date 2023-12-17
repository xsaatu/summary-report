const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');


function processDailyData(dayFolder, filePath) {
  try {
    // Membaca file Excel harian
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Mengambil data dari sel tertentu
    const maxPrice = sheet['B2'] ? sheet['B2'].v : null;
    const avgPrice = sheet['C2'] ? sheet['C2'].v : null;

    // Mengembalikan hasil
    return { Date: dayFolder, Max: maxPrice, Avg: avgPrice };
  } catch (error) {
    console.error('Error processing daily data:', error.message);
    throw error;
  }
}

async function generateMonthlyReport(monthFolder) {
  // Mendapatkan daftar folder hari dalam folder bulan
  const dayFolders = fs.readdirSync(monthFolder).filter(item => fs.statSync(path.join(monthFolder, item)).isDirectory());

  // Inisialisasi data untuk hasil akhir
  const resultData = [];
  

  // Memproses setiap folder hari
  for (const dayFolder of dayFolders) {
    const dayFolderPath = path.join(monthFolder, dayFolder);

    // Mendapatkan daftar file dalam folder hari
    const dayFiles = fs.readdirSync(dayFolderPath);

    // Memeriksa apakah ada file di dalam folder hari
    if (dayFiles.length > 0) {
      try {
        // Membaca data harian dengan menggunakan nama folder sebagai Date
        const dailyData = await processDailyData(dayFolder, path.join(dayFolderPath, dayFiles[0]));

        // Menambahkan data harian ke hasil akhir
        resultData.push(dailyData);
      } catch (error) {
        console.error('Error processing daily folder:', dayFolder, error.message);
      }
    } else {
      console.warn('No Excel file found in the daily folder:', dayFolder);
    }
  }

  // Membuat objek hasil untuk worksheet
  const resultWorksheet = [['Date', 'Max Price', 'Average Price']];

  resultWorksheet.unshift([{ t: 's', v: 'Summary', s: { font: { size: 14, bold: true } } }]);

  // Menambahkan hasil akhir ke worksheet
  resultData.forEach(data => {
    resultWorksheet.push([data.Date, data.Max, data.Avg]);
  });

  // Menyimpan hasil akhir ke file Excel utama
  const resultWorkbook = XLSX.utils.book_new();
  const resultSheet = XLSX.utils.aoa_to_sheet(resultWorksheet);
  XLSX.utils.book_append_sheet(resultWorkbook, resultSheet, 'Monthly Report');
  XLSX.writeFile(resultWorkbook, path.join(monthFolder, 'monthly_report.xlsx'));
}

// path folder
const pathFolder = 'D:/Dev/excell/februari';

// Menghasilkan laporan bulanan
generateMonthlyReport(pathFolder).then(() => {
  console.log('Monthly report generated successfully.');
}).catch(error => {
  console.error('Error:', error.message);
});
