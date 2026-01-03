// applyLandscape.js
const Docxtemplater = require('docxtemplater');
const PizZip = require('pizzip');
const fs = require('fs').promises;
const path = require('path');

async function applyLandscapeOrientation(filePath) {
    try {
        // Читаем документ
        const content = await fs.readFile(filePath, 'binary');
        const zip = new PizZip(content);
        
        // Получаем XML документа
        let xml = zip.files['word/document.xml'].asText();
        
        // Находим все секции и меняем ориентацию
        const sectionRegex = /<w:sectPr>[\s\S]*?<\/w:sectPr>/g;
        
        xml = xml.replace(sectionRegex, (section) => {
            // Заменяем размер страницы на альбомный
            return section.replace(
                /<w:pgSz[^>]*>/g,
                '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>'
            );
        });
        
        // Если секций нет, добавляем
        if (!xml.includes('<w:sectPr>')) {
            const bodyEnd = xml.indexOf('</w:body>');
            const landscapeSection = `
                <w:p>
                    <w:pPr>
                        <w:sectPr>
                            <w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>
                            <w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417"/>
                        </w:sectPr>
                    </w:pPr>
                </w:p>
            `;
            xml = xml.slice(0, bodyEnd) + landscapeSection + xml.slice(bodyEnd);
        }
        
        // Сохраняем изменения
        zip.file('word/document.xml', xml);
        
        // Генерируем новый документ
        const buffer = zip.generate({ type: 'nodebuffer' });
        
        // Сохраняем с новым именем
        const filename = `landscape_${Date.now()}.docx`;
        const newFilePath = path.join(__dirname, 'downloads', filename);
        await fs.writeFile(newFilePath, buffer);
        
        return { filename };
    } catch (error) {
        throw new Error(`Ошибка поворота страницы: ${error.message}`);
    }
}

module.exports = { applyLandscapeOrientation };