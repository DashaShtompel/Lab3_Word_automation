// app.js - Основная логика фронтенда

class DocumentAutomationApp {
    constructor() {
        this.initializeElements();
        this.bindEvents();
        this.checkServerStatus();
        this.generateTableData(5);
    }

    initializeElements() {
        // Форма
        this.templateUpload = document.getElementById('templateUpload');
        this.contractNumber = document.getElementById('contractNumber');
        this.city = document.getElementById('city');
        this.contractDate = document.getElementById('contractDate');
        this.clientName = document.getElementById('clientName');
        this.clientFio = document.getElementById('clientFio');
        this.executorName = document.getElementById('executorName');
        this.executorFio = document.getElementById('executorFio');
        this.numOfDays = document.getElementById('numOfDays');
        
        // Кнопки формы
        this.generateRandomDataBtn = document.getElementById('generateRandomData');
        this.clearFormBtn = document.getElementById('clearForm');
        
        // Таблица
        this.rowCount = document.getElementById('rowCount');
        this.generateTableBtn = document.getElementById('generateTable');
        this.tableBody = document.getElementById('tableBody');
        this.dataTable = document.getElementById('dataTable');
        
        // Генерация документов
        this.generateDocBtn = document.getElementById('generateDoc');
        this.generateLargeBtn = document.getElementById('generateLarge');
        this.applyLandscapeBtn = document.getElementById('applyLandscape');
        
        // Результаты
        this.resultsSection = document.getElementById('resultsSection');
        this.resultsContainer = document.getElementById('resultsContainer');
        
        // Статус
        this.statusIndicator = document.getElementById('statusIndicator');
        this.statusText = document.getElementById('statusText');
        
        // Модальное окно
        this.loadingModal = document.getElementById('loadingModal');
        this.loadingMessage = document.getElementById('loadingMessage');
        this.progressBar = document.getElementById('progressBar');
        
        // Файл
        this.fileInfo = document.getElementById('fileInfo');
    }

    bindEvents() {
        // Форма
        this.generateRandomDataBtn.addEventListener('click', () => this.fillRandomData());
        this.clearFormBtn.addEventListener('click', () => this.clearForm());
        
        // Загрузка файла
        this.templateUpload.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (file) {
                this.fileInfo.innerHTML = `
                    <i class="fas fa-file-word"></i> ${file.name} (${(file.size / 1024).toFixed(2)} KB)
                `;
            }
        });
        
        // Таблица
        this.generateTableBtn.addEventListener('click', () => {
            const count = parseInt(this.rowCount.value) || 5;
            this.generateTableData(count);
        });
        
        // Генерация документов
        this.generateDocBtn.addEventListener('click', () => this.generateDocument());
        this.generateLargeBtn.addEventListener('click', () => this.generateLargeDocument());
        this.applyLandscapeBtn.addEventListener('click', () => this.applyLandscape());
        
        // Автосохранение при вводе
        ['contractNumber', 'city', 'clientName', 'clientFio', 'executorName', 'executorFio'].forEach(id => {
            document.getElementById(id).addEventListener('input', () => this.autoSave());
        });
        
        // Сортировка по клику на заголовки
        this.setupTableSorting();
    }

    setupTableSorting() {
        const headers = this.dataTable.querySelectorAll('thead th');
        headers.forEach((header, index) => {
            header.style.cursor = 'pointer';
            header.addEventListener('click', () => {
                this.sortTableByColumn(index);
            });
        });
    }

    checkServerStatus() {
        fetch('http://localhost:3000/health')
            .then(response => {
                if (response.ok) {
                    this.statusIndicator.classList.add('connected');
                    this.statusText.textContent = 'Сервер запущен';
                } else {
                    throw new Error('Сервер не отвечает');
                }
            })
            .catch(() => {
                this.statusIndicator.classList.remove('connected');
                this.statusText.textContent = 'Сервер не запущен';
            });
    }

    sortTableByColumn(columnIndex) {
        const rows = Array.from(this.tableBody.querySelectorAll('tr'));
        
        // Определяем тип данных в колонке
        const dataType = this.getColumnDataType(columnIndex);
        
        // Сортируем строки
        rows.sort((a, b) => {
            const aCell = a.cells[columnIndex];
            const bCell = b.cells[columnIndex];
            
            if (!aCell || !bCell) return 0;
            
            let aValue = aCell.textContent.trim();
            let bValue = bCell.textContent.trim();
            
            // Для пустых значений
            if (aValue === '' && bValue === '') return 0;
            if (aValue === '') return -1;
            if (bValue === '') return 1;
            
            // Сортировка в зависимости от типа данных
            if (dataType === 'number') {
                const aNum = parseFloat(aValue.replace(',', '.'));
                const bNum = parseFloat(bValue.replace(',', '.'));
                return aNum - bNum;
            } else if (dataType === 'date') {
                const aDate = this.parseDate(aValue);
                const bDate = this.parseDate(bValue);
                return aDate - bDate;
            } else {
                // Строковая сортировка
                return aValue.localeCompare(bValue, 'ru');
            }
        });
        
        // Обновляем таблицу
        this.tableBody.innerHTML = '';
        rows.forEach(row => this.tableBody.appendChild(row));
        
        this.showNotification(`Таблица отсортирована по колонке ${columnIndex + 1}`, 'info');
    }

    getColumnDataType(columnIndex) {
        // Определяем тип данных по заголовку или содержимому
        const headers = ['Наименование услуги', 'Дата выполнения', 'Количество', 'Стоимость за единицу', 'Сумма'];
        
        if (columnIndex === 0) return 'string'; // Название
        if (columnIndex === 1) return 'date';   // Дата
        return 'number';                        // Количество, цена, сумма
    }

    parseDate(dateString) {
        // Парсим дату в формате ДД.ММ.ГГГГ
        const parts = dateString.split('.');
        if (parts.length === 3) {
            return new Date(parts[2], parts[1] - 1, parts[0]).getTime();
        }
        return new Date(dateString).getTime();
    }

    fillRandomData() {
        const cities = ['Москва', 'Санкт-Петербург', 'Новосибирск', 'Екатеринбург', 'Казань'];
        const clients = ['ООО "Ромашка"', 'АО "Вектор"', 'ИП "Старт"', 'ЗАО "Технологии"', 'ООО "Прогресс"'];
        const names = ['Иванов Иван Иванович', 'Петров Петр Петрович', 'Сидоров Сидор Сидорович', 'Алексеев Алексей Алексеевич'];
        
        this.contractNumber.value = `Д-${Math.floor(Math.random() * 1000)}/2024`;
        this.city.value = `г. ${cities[Math.floor(Math.random() * cities.length)]}`;
        this.contractDate.value = this.getRandomDate();
        this.clientName.value = clients[Math.floor(Math.random() * clients.length)];
        this.clientFio.value = names[Math.floor(Math.random() * names.length)];
        this.executorName.value = 'ИП "Автоматизация Сервис"';
        this.executorFio.value = 'Сергеев Сергей Сергеевич';
        this.numOfDays.value = Math.floor(Math.random() * 30) + 1;
        
        this.showNotification('Форма заполнена случайными данными', 'success');
    }

    clearForm() {
        const formElements = [
            this.contractNumber, this.city, this.contractDate,
            this.clientName, this.clientFio, this.executorName,
            this.executorFio, this.numOfDays
        ];
        
        formElements.forEach(element => element.value = '');
        this.contractDate.value = new Date().toISOString().split('T')[0];
        this.numOfDays.value = '30';
        
        this.showNotification('Форма очищена', 'info');
    }

    generateTableData(count) {
        this.tableBody.innerHTML = '';
        
        // Массивы для генерации осмысленных данных
        const services = ['Консультация', 'Разработка', 'Тестирование', 'Поддержка', 'Обучение'];
        const adjectives = ['Базовый', 'Расширенный', 'Профессиональный', 'Корпоративный', 'Индивидуальный'];
        
        for (let i = 0; i < count; i++) {
            const row = document.createElement('tr');
            
            // Генерация данных
            const serviceName = `${adjectives[i % adjectives.length]} ${services[i % services.length]}`;
            const date = this.getRandomDate();
            const quantity = Math.floor(Math.random() * 100) + 1;
            const price = (Math.random() * 1000 + 10).toFixed(2);
            const total = (quantity * parseFloat(price)).toFixed(2);
            
            row.innerHTML = `
                <td>${serviceName}</td>
                <td>${this.formatDate(date)}</td>
                <td>${quantity}</td>
                <td>${price}</td>
                <td>${total}</td>
            `;
            
            this.tableBody.appendChild(row);
        }
        
        this.showNotification(`Сгенерировано ${count} строк таблицы`, 'success');
    }

    async generateDocument() {
        if (!this.validateForm()) return;
        
        const formData = new FormData();
        const file = this.templateUpload.files[0];
        
        if (!file) {
            this.showNotification('Пожалуйста, выберите шаблон Word', 'error');
            return;
        }
        
        formData.append('template', file);
        formData.append('data', JSON.stringify(this.getFormData()));
        formData.append('tableData', JSON.stringify(this.getTableData()));
        
        this.showLoading('Создание документа...');
        
        try {
            const response = await fetch('http://localhost:3000/api/generate', {
                method: 'POST',
                body: formData
            });
            
            const result = await response.json();
            
            if (result.success) {
                this.showResults(result);
                this.showNotification('Документ успешно создан!', 'success');
                
                // Автоматически скачиваем Word документ
                setTimeout(() => {
                    const link = document.createElement('a');
                    link.href = `http://localhost:3000${result.wordUrl}`;
                    link.download = 'document.docx';
                    document.body.appendChild(link);
                    link.click();
                    document.body.removeChild(link);
                }, 1000);
            } else {
                throw new Error(result.error || 'Произошла ошибка');
            }
        } catch (error) {
            this.showNotification(`Ошибка: ${error.message}`, 'error');
            console.error('Детали ошибки:', error);
        } finally {
            this.hideLoading();
        }
    }

    async generateLargeDocument() {
        const formData = new FormData();
        const file = this.templateUpload.files[0];
        
        if (!file) {
            this.showNotification('Пожалуйста, выберите шаблон Word', 'error');
            return;
        }
        
        formData.append('template', file);
        formData.append('rows', 10000);
        
        await this.sendRequest('/api/generate-large', formData, 'Генерация большого документа');
    }

    async applyLandscape() {
        const formData = new FormData();
        const file = this.templateUpload.files[0];
        
        if (!file) {
            this.showNotification('Пожалуйста, выберите шаблон Word', 'error');
            return;
        }
        
        formData.append('template', file);
        
        await this.sendRequest('/api/landscape', formData, 'Применение альбомной ориентации');
    }

    async sendRequest(endpoint, formData, message) {
        this.showLoading(message);
        
        try {
            const response = await fetch(`http://localhost:3000${endpoint}`, {
                method: 'POST',
                body: formData
            });
            
            const result = await response.json();
            
            if (result.success) {
                this.showResults(result);
                this.showNotification('Операция выполнена успешно!', 'success');
            } else {
                throw new Error(result.error || 'Произошла ошибка');
            }
        } catch (error) {
            this.showNotification(`Ошибка: ${error.message}`, 'error');
        } finally {
            this.hideLoading();
        }
    }

    showResults(result) {
        this.resultsSection.style.display = 'block';
        
        let html = '';
        
        if (result.wordUrl) {
            html += `
                <div class="result-item">
                    <i class="fas fa-file-word"></i>
                    <div>
                        <a href="http://localhost:3000${result.wordUrl}" download>
                            Скачать Word документ
                        </a>
                        <small>${result.message}</small>
                    </div>
                </div>
            `;
        }
        
        if (result.pdfUrl) {
            html += `
                <div class="result-item">
                    <i class="fas fa-file-pdf"></i>
                    <div>
                        <a href="http://localhost:3000${result.pdfUrl}" download>
                            Скачать PDF документ
                        </a>
                        <small>Версия для печати</small>
                    </div>
                </div>
            `;
        }
        
        this.resultsContainer.innerHTML = html;
    }

    getFormData() {
        return {
            ContractNumber: this.contractNumber.value,
            City: this.city.value,
            ContractDate: this.contractDate.value ? 
                new Date(this.contractDate.value).toLocaleDateString('ru-RU') : '',
            ClientName: this.clientName.value,
            ClientFio: this.clientFio.value,
            ExecutorName: this.executorName.value,
            ExecutorFio: this.executorFio.value,
            NumOfDays: this.numOfDays.value
        };
    }

    getTableData() {
        const rows = [];
        this.tableBody.querySelectorAll('tr').forEach(row => {
            rows.push({
                product: row.cells[0].textContent,
                date: row.cells[1].textContent,
                quantity: row.cells[2].textContent,
                price: row.cells[3].textContent,
                total: row.cells[4].textContent
            });
        });
        return rows;
    }

    validateForm() {
        const required = [
            this.contractNumber, this.city, this.contractDate,
            this.clientName, this.clientFio, this.executorName,
            this.executorFio
        ];
        
        for (const field of required) {
            if (!field.value.trim()) {
                this.showNotification(`Заполните поле: ${field.previousElementSibling.textContent}`, 'error');
                field.focus();
                return false;
            }
        }
        
        return true;
    }

    getRandomDate() {
        const start = new Date(2023, 0, 1);
        const end = new Date();
        const date = new Date(start.getTime() + Math.random() * (end.getTime() - start.getTime()));
        return date.toISOString().split('T')[0];
    }

    formatDate(dateString) {
        const date = new Date(dateString);
        return date.toLocaleDateString('ru-RU');
    }

    showLoading(message) {
        this.loadingMessage.textContent = message;
        this.loadingModal.style.display = 'flex';
        this.progressBar.style.width = '30%';
        
        // Симуляция прогресса
        const interval = setInterval(() => {
            const currentWidth = parseFloat(this.progressBar.style.width);
            if (currentWidth < 90) {
                this.progressBar.style.width = (currentWidth + 10) + '%';
            }
        }, 500);
        
        this.loadingInterval = interval;
    }

    hideLoading() {
        if (this.loadingInterval) {
            clearInterval(this.loadingInterval);
        }
        this.progressBar.style.width = '100%';
        
        setTimeout(() => {
            this.loadingModal.style.display = 'none';
            this.progressBar.style.width = '0%';
        }, 500);
    }

    showNotification(message, type = 'info') {
        // Создаем уведомление
        const notification = document.createElement('div');
        notification.className = `notification notification-${type}`;
        notification.innerHTML = `
            <i class="fas fa-${type === 'error' ? 'exclamation-circle' : type === 'success' ? 'check-circle' : 'info-circle'}"></i>
            <span>${message}</span>
            <button class="notification-close"><i class="fas fa-times"></i></button>
        `;
        
        // Стили для уведомления
        notification.style.cssText = `
            position: fixed;
            top: 20px;
            right: 20px;
            background: ${type === 'error' ? '#e74c3c' : type === 'success' ? '#2ecc71' : '#3498db'};
            color: white;
            padding: 15px 20px;
            border-radius: 5px;
            display: flex;
            align-items: center;
            gap: 10px;
            z-index: 10000;
            animation: slideInRight 0.3s ease-out;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        `;
        
        document.body.appendChild(notification);
        
        // Кнопка закрытия
        notification.querySelector('.notification-close').addEventListener('click', () => {
            notification.style.animation = 'slideOutRight 0.3s ease-out';
            setTimeout(() => notification.remove(), 300);
        });
        
        // Автоматическое скрытие
        setTimeout(() => {
            if (notification.parentNode) {
                notification.style.animation = 'slideOutRight 0.3s ease-out';
                setTimeout(() => notification.remove(), 300);
            }
        }, 5000);
    }

    autoSave() {
        // Сохранение данных в localStorage
        const data = this.getFormData();
        localStorage.setItem('documentFormData', JSON.stringify(data));
    }

    loadSavedData() {
        const saved = localStorage.getItem('documentFormData');
        if (saved) {
            const data = JSON.parse(saved);
            Object.keys(data).forEach(key => {
                const element = document.getElementById(key);
                if (element) element.value = data[key];
            });
        }
    }
}

// Инициализация приложения при загрузке страницы
document.addEventListener('DOMContentLoaded', () => {
    window.app = new DocumentAutomationApp();
    window.app.loadSavedData();
    
    // Добавляем стили для анимаций уведомлений
    const style = document.createElement('style');
    style.textContent = `
        @keyframes slideInRight {
            from { transform: translateX(100%); opacity: 0; }
            to { transform: translateX(0); opacity: 1; }
        }
        @keyframes slideOutRight {
            from { transform: translateX(0); opacity: 1; }
            to { transform: translateX(100%); opacity: 0; }
        }
        .notification-close {
            background: none;
            border: none;
            color: white;
            cursor: pointer;
            padding: 0;
            margin-left: 10px;
        }
        
        /* Стили для кликабельных заголовков */
        th {
            cursor: pointer;
            position: relative;
            user-select: none;
        }
        
        th:hover {
            background-color: rgba(0, 0, 0, 0.1);
        }
    `;
    document.head.appendChild(style);
});