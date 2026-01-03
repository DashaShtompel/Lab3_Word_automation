// generateLargeTable.js
const PizZip = require('pizzip');
const fs = require('fs').promises;
const path = require('path');

async function generateLargeTable(templatePath, rowCount) {
    try {
        // Читаем шаблон
        const content = await fs.readFile(templatePath, 'binary');
        const zip = new PizZip(content);
        
        // Получаем XML документа
        let xml = zip.files['word/document.xml'].asText();
        
        // Генерируем таблицу
        const tableXML = generateTableXML(rowCount);
        
        // Вставляем таблицу в документ
        // Ищем метку <!--TABLE--> в шаблоне
        if (xml.includes('<!--TABLE-->')) {
            xml = xml.replace('<!--TABLE-->', tableXML);
        } else {
            // Если метки нет, вставляем в конец тела документа
            const bodyEnd = xml.indexOf('</w:body>');
            xml = xml.slice(0, bodyEnd) + tableXML + xml.slice(bodyEnd);
        }
        
        // Обновляем документ
        zip.file('word/document.xml', xml);
        
        // Генерируем новый документ
        const buffer = zip.generate({ type: 'nodebuffer' });
        
        // Сохраняем
        const wordFilename = `large_table_${rowCount}_${Date.now()}.docx`;
        const wordPath = path.join(__dirname, 'downloads', wordFilename);
        await fs.writeFile(wordPath, buffer);

        return { wordFilename };
    } catch (error) {
        throw new Error(`Ошибка генерации большой таблицы: ${error.message}`);
    }
}

function generateTableXML(rowCount) {
    // Создаем заголовок таблицы
    let xml = `
        <w:tbl>
            <w:tblPr>
                <w:tblW w:w="0" w:type="auto"/>
                <w:jc w:val="center"/>
                <w:tblBorders>
                    <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                    <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                    <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                    <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                </w:tblBorders>
            </w:tblPr>
            <w:tblGrid>
                <w:gridCol w:w="2000"/>
                <w:gridCol w:w="2000"/>
                <w:gridCol w:w="2000"/>
                <w:gridCol w:w="2000"/>
                <w:gridCol w:w="2000"/>
            </w:tblGrid>
    `;
    
    // Заголовок таблицы
    xml += `
        <w:tr>
            <w:tc>
                <w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
                <w:p><w:r><w:t>Наименование товара</w:t></w:r></w:p>
            </w:tc>
            <w:tc>
                <w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
                <w:p><w:r><w:t>Дата поставки</w:t></w:r></w:p>
            </w:tc>
            <w:tc>
                <w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
                <w:p><w:r><w:t>Количество</w:t></w:r></w:p>
            </w:tc>
            <w:tc>
                <w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
                <w:p><w:r><w:t>Цена за единицу</w:t></w:r></w:p>
            </w:tc>
            <w:tc>
                <w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
                <w:p><w:r><w:t>Сумма</w:t></w:r></w:p>
            </w:tc>
        </w:tr>
    `;
    
    // Генерируем строки
    for (let i = 0; i < rowCount; i++) {
        const productName = `Товар ${i + 1}`;
        const date = getRandomDate();
        const quantity = Math.floor(Math.random() * 100) + 1;
        const price = (Math.random() * 1000 + 10).toFixed(2);
        const total = (quantity * parseFloat(price)).toFixed(2);
        
        xml += `
            <w:tr>
                <w:tc>
                    <w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
                    <w:p><w:r><w:t>${productName}</w:t></w:r></w:p>
                </w:tc>
                <w:tc>
                    <w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
                    <w:p><w:r><w:t>${date}</w:t></w:r></w:p>
                </w:tc>
                <w:tc>
                    <w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
                    <w:p><w:r><w:t>${quantity}</w:t></w:r></w:p>
                </w:tc>
                <w:tc>
                    <w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
                    <w:p><w:r><w:t>${price}</w:t></w:r></w:p>
                </w:tc>
                <w:tc>
                    <w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
                    <w:p><w:r><w:t>${total}</w:t></w:r></w:p>
                </w:tc>
            </w:tr>
        `;
    }
    
    xml += '</w:tbl>';
    return xml;
}

function getRandomDate() {
    const start = new Date(2023, 0, 1);
    const end = new Date();
    const date = new Date(start.getTime() + Math.random() * (end.getTime() - start.getTime()));
    return date.toLocaleDateString('ru-RU');
}

module.exports = { generateLargeTable };