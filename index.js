const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

const directoryPath = './xlsxFiles'; // ディレクトリのパスを設定
const outputDirectoryPath = './jsonOutput'; // 出力先のディレクトリのパスを設定

// 指定されたディレクトリ内のすべてのxlsxファイルを読み込む関数
fs.readdir(directoryPath, (err, files) => {
    if (err) {
        console.log('Error reading directory:', err);
        return;
    }

    files.forEach(file => {
        if (path.extname(file) === '.xlsx') {
            // Excelファイルを読み込む
            const filePath = path.join(directoryPath, file);
            const workbook = xlsx.readFile(filePath);
            const sheetName = workbook.SheetNames[0]; // 最初のシートを使用
            const jsonData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

            // JSONデータをファイルに書き出す
            const jsonFilePath = path.join(outputDirectoryPath, `${path.basename(file, '.xlsx')}.json`);
            fs.writeFile(jsonFilePath, JSON.stringify(jsonData, null, 2), (err) => {
                if (err) {
                    console.log(`Error writing JSON for ${file}:`, err);
                } else {
                    console.log(`JSON for ${file} written successfully`);
                }
            });
        }
    });
});
