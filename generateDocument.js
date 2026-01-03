// generateDocument.js - с поддержкой PDF
const Docxtemplater = require('docxtemplater');
const PizZip = require('pizzip');
const fs = require('fs').promises;
const path = require('path');
const { exec } = require('child_process');
const util = require('util');
const execAsync = util.promisify(exec);

async function generateDocument(templatePath, data) {
    try {
        console.log('=== НАЧАЛО ГЕНЕРАЦИИ ДОКУМЕНТА ===');
        console.log('Шаблон:', templatePath);
        
        // 1. Читаем шаблон
        const content = await fs.readFile(templatePath, 'binary');
        const zip = new PizZip(content);
        
        // 2. Создаем документ
        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
        });

        // 3. Подготавливаем данные
        const templateData = {
            ContractNumber: data.ContractNumber || data.contractNumber || '',
            City: data.City || data.city || '',
            ContractDate: data.ContractDate || data.contractDate || '',
            ClientName: data.ClientName || data.clientName || '',
            ClientFio: data.ClientFio || data.clientFio || '',
            ExecutorName: data.ExecutorName || data.executorName || '',
            ExecutorFio: data.ExecutorFio || data.executorFio || '',
            NumOfDays: data.NumOfDays || data.numOfDays || '',
            items: data.items || []
        };

        console.log('📊 Данные для шаблона:', templateData);
        
        // 4. Устанавливаем данные
        doc.setData(templateData);
        
        // 5. Рендерим документ
        doc.render();
        
        // 6. Генерируем Word документ
        const buffer = doc.getZip().generate({
            type: 'nodebuffer',
            compression: 'DEFLATE',
        });

        // 7. Сохраняем Word документ
        const wordFilename = `document_${Date.now()}.docx`;
        const wordPath = path.join(__dirname, 'downloads', wordFilename);
        
        // Создаем папку если не существует
        try {
            await fs.access(path.join(__dirname, 'downloads'));
        } catch {
            await fs.mkdir(path.join(__dirname, 'downloads'), { recursive: true });
        }
        
        await fs.writeFile(wordPath, buffer);
        console.log(`✅ Word документ сохранен: ${wordFilename}`);
        
        // 8. Пробуем сгенерировать PDF
        let pdfFilename = null;
        
        try {
            pdfFilename = await convertToPDF(wordPath);
            console.log(`✅ PDF документ создан: ${pdfFilename}`);
        } catch (pdfError) {
            console.log('⚠️ PDF не сгенерирован:', pdfError.message);
            console.log('📋 Инструкция по установке LibreOffice:');
            console.log('1. Установите Homebrew: /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"');
            console.log('2. Установите LibreOffice: brew install --cask libreoffice');
            console.log('3. Перезапустите сервер');
        }

        return {
            wordFilename: wordFilename,
            pdfFilename: pdfFilename
        };
        
    } catch (error) {
        console.error('❌ Ошибка генерации:', error);
        throw error;
    }
}

// Функция конвертации в PDF
async function convertToPDF(wordPath) {
    console.log('🔄 Конвертация в PDF...');
    
    const pdfFilename = path.basename(wordPath).replace('.docx', '.pdf');
    const pdfPath = path.join(__dirname, 'downloads', pdfFilename);
    
    // Проверяем разные пути к LibreOffice
    const libreofficePaths = [
        '/Applications/LibreOffice.app/Contents/MacOS/soffice',
        '/Applications/LibreOffice.app/Contents/MacOS/soffice.bin',
        '/opt/homebrew/bin/soffice',
        '/usr/local/bin/soffice'
    ];
    
    let libreofficePath = null;
    
    // Ищем LibreOffice
    for (const path of libreofficePaths) {
        try {
            await fs.access(path);
            libreofficePath = path;
            console.log(`✅ Найден LibreOffice: ${path}`);
            break;
        } catch {
            continue;
        }
    }
    
    if (!libreofficePath) {
        throw new Error('LibreOffice не найден. Установите: brew install --cask libreoffice');
    }
    
    // Команда для конвертации
    const command = `"${libreofficePath}" --headless --convert-to pdf --outdir "${path.dirname(pdfPath)}" "${wordPath}"`;
    
    console.log('🔄 Выполняем команду:', command);
    
    try {
        const { stdout, stderr } = await execAsync(command, { timeout: 30000 });
        
        if (stderr) {
            console.log('⚠️ Предупреждение LibreOffice:', stderr);
        }
        
        console.log('✅ Конвертация завершена:', stdout);
        
        // Проверяем что PDF создан
        await fs.access(pdfPath);
        
        return pdfFilename;
        
    } catch (execError) {
        console.error('❌ Ошибка конвертации:', execError.message);
        
        // Альтернативный способ через модуль
        try {
            console.log('🔄 Пробуем альтернативный способ конвертации...');
            return await convertToPDFAlternative(wordPath, pdfPath);
        } catch (altError) {
            throw new Error(`Не удалось конвертировать в PDF: ${execError.message}`);
        }
    }
}

// Альтернативный способ конвертации
async function convertToPDFAlternative(wordPath, pdfPath) {
    try {
        // Пробуем использовать libreoffice-convert если установлен
        const libre = require('libreoffice-convert');
        libre.convertAsync = util.promisify(libre.convert);
        
        const input = await fs.readFile(wordPath);
        const pdfBuffer = await libre.convertAsync(input, '.pdf', undefined);
        
        await fs.writeFile(pdfPath, pdfBuffer);
        
        return path.basename(pdfPath);
    } catch (error) {
        throw new Error(`Альтернативный способ также не сработал: ${error.message}`);
    }
}

module.exports = { generateDocument };