// server.js
const express = require('express');
const cors = require('cors');
const multer = require('multer');
const path = require('path');
const fs = require('fs').promises;
const { generateDocument } = require('./generateDocument');
const { generateLargeTable } = require('./generateLargeTable');

const app = express();
const port = 3000;

// Настройка загрузки файлов
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, 'uploads/');
    },
    filename: (req, file, cb) => {
        cb(null, Date.now() + '-' + file.originalname);
    }
});

const upload = multer({ 
    storage: storage,
    fileFilter: (req, file, cb) => {
        const allowedTypes = ['.dotx', '.docx'];
        const ext = path.extname(file.originalname).toLowerCase();
        if (allowedTypes.includes(ext)) {
            cb(null, true);
        } else {
            cb(new Error('Разрешены только файлы .dotx и .docx'));
        }
    }
});

// Создаем необходимые папки
const initFolders = async () => {
    const folders = ['uploads', 'downloads', 'templates'];
    for (const folder of folders) {
        try {
            await fs.access(folder);
        } catch {
            await fs.mkdir(folder, { recursive: true });
        }
    }
};

// Middleware
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));
app.use('/downloads', express.static('downloads'));
app.use(express.static('.'));

// Health check endpoint
app.get('/health', (req, res) => {
    res.json({ 
        status: 'ok', 
        timestamp: new Date().toISOString(),
        message: 'Сервер работает нормально'
    });
});

// Главная страница
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

// Эндпоинт для тестирования
app.get('/api/test', (req, res) => {
    res.json({
        message: 'API работает',
        endpoints: [
            '/api/generate - генерация документа',
            '/api/generate-large - генерация с большой таблицей',
            '/api/landscape - альбомная ориентация'
        ]
    });
});

// Генерация документа с пользовательскими данными
app.post('/api/generate', upload.single('template'), async (req, res) => {
    try {
        console.log('Получен запрос на генерацию документа');
        
        if (!req.file) {
            throw new Error('Файл шаблона не загружен');
        }

        console.log('Загружен файл:', req.file.originalname);
        
        const userData = req.body.data ? JSON.parse(req.body.data) : {};
        const tableData = req.body.tableData ? JSON.parse(req.body.tableData) : [];
        
        console.log('Данные пользователя:', userData);
        console.log('Данные таблицы (первые 3 строки):', tableData.slice(0, 3));

        // Объединяем данные
        const allData = {
            ...userData,
            items: tableData
        };

        console.log('Все данные для документа:', allData);
        
        const result = await generateDocument(req.file.path, allData);
        
        // Удаляем временный файл
        await fs.unlink(req.file.path);
        
        console.log('Документ успешно создан:', result);
        
        res.json({
            success: true,
            wordUrl: `/downloads/${result.wordFilename}`,
            pdfUrl: `/downloads/${result.pdfFilename}`,
            message: 'Документ успешно создан'
        });
    } catch (error) {
        console.error('Ошибка генерации документа:', error);
        res.status(500).json({ 
            success: false, 
            error: error.message,
            details: error.stack
        });
    }
});

// Генерация документа с большой таблицей
app.post('/api/generate-large', upload.single('template'), async (req, res) => {
    try {
        console.log('Генерация документа с большой таблицей');
        
        if (!req.file) {
            throw new Error('Файл шаблона не загружен');
        }

        const rows = parseInt(req.body.rows) || 10000;
        const result = await generateLargeTable(req.file.path, rows);
        
        // Удаляем временный файл
        await fs.unlink(req.file.path);
        
        res.json({
            success: true,
            wordUrl: `/downloads/${result.wordFilename}`,
            pdfUrl: `/downloads/${result.pdfFilename}`,
            message: `Документ с ${rows} строками успешно создан`
        });
    } catch (error) {
        console.error('Ошибка генерации большого документа:', error);
        res.status(500).json({ 
            success: false, 
            error: error.message 
        });
    }
});

// Применение альбомной ориентации
app.post('/api/landscape', upload.single('template'), async (req, res) => {
    try {
        if (!req.file) {
            throw new Error('Файл шаблона не загружен');
        }

        // Сначала создаем обычный документ
        const userData = req.body.data ? JSON.parse(req.body.data) : {};
        const allData = {
            ...userData,
            items: []
        };

        const documentResult = await generateDocument(req.file.path, allData);
        
        // Теперь применяем альбомную ориентацию
        const { applyLandscapeOrientation } = require('./applyLandscape');
        const landscapeResult = await applyLandscapeOrientation(
            path.join(__dirname, 'downloads', documentResult.wordFilename)
        );
        
        res.json({
            success: true,
            wordUrl: `/downloads/${landscapeResult.filename}`,
            message: 'Документ повернут в альбомную ориентацию'
        });
    } catch (error) {
        console.error('Ошибка поворота страницы:', error);
        res.status(500).json({ 
            success: false, 
            error: error.message 
        });
    }
});

// Запуск сервера
const startServer = async () => {
    await initFolders();
    
    app.listen(port, '0.0.0.0', () => {
        console.log(`
╔═══════════════════════════════════════════════════════╗
║                                                       ║
║   🚀 Сервер запущен!                                 ║
║                                                       ║
║   📍 Адрес: http://localhost:${port}                   ║
║   🌐 Также доступен по IP                            ║
║   📅 ${new Date().toLocaleString('ru-RU')}            ║
║                                                       ║
╚═══════════════════════════════════════════════════════╝
        `);
        
        console.log('\n📁 Структура папок:');
        console.log('├── uploads/    - для временных файлов');
        console.log('├── downloads/  - готовые документы');
        console.log('└── templates/  - шаблоны Word\n');
        
        console.log('📝 Для тестирования:');
        console.log('1. Перейдите на http://localhost:3000');
        console.log('2. Загрузите шаблон .dotx/.docx');
        console.log('3. Заполните форму и нажмите "Создать документ"\n');
    });
};

startServer().catch(console.error);