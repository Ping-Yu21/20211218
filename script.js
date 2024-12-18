document.addEventListener('DOMContentLoaded', function() {
    const borrowForm = document.getElementById('borrowForm');
    const borrowList = document.getElementById('borrowList').getElementsByTagName('tbody')[0];
    const importDialog = document.getElementById('importDialog');
    const importButton = document.getElementById('importItems');
    const confirmImport = document.getElementById('confirmImport');
    const cancelImport = document.getElementById('cancelImport');
    const itemsList = document.getElementById('itemsList');
    const itemSelect = document.getElementById('itemName');
    const excelFile = document.getElementById('excelFile');
    const previewList = document.getElementById('previewList');
    const returnDialog = document.getElementById('returnDialog');
    const actualReturnDate = document.getElementById('actualReturnDate');
    const confirmReturn = document.getElementById('confirmReturn');
    const cancelReturn = document.getElementById('cancelReturn');
    
    // 從 localStorage 載入資料
    let borrowRecords = JSON.parse(localStorage.getItem('borrowRecords')) || [];
    let items = JSON.parse(localStorage.getItem('items')) || [];
    
    let currentReturnId = null; // 用於存儲當前要歸還的記錄 ID
    
    // 初始化顯示
    updateItemsList();
    displayBorrowRecords();
    
    // 初始化行事曆
    const calendarEl = document.getElementById('calendar');
    const calendar = new FullCalendar.Calendar(calendarEl, {
        initialView: 'dayGridMonth',
        locale: 'zh-tw',
        headerToolbar: {
            left: 'prev,next today',
            center: 'title',
            right: 'dayGridMonth,dayGridWeek'
        },
        events: getCalendarEvents(),
        eventClick: function(info) {
            alert(`
                物品：${info.event.title}
                借用人：${info.event.extendedProps.borrower}
                借用日期：${new Date(info.event.start).toLocaleString()}
                歸還日期：${new Date(info.event.extendedProps.returnDate).toLocaleString()}
                狀態：${info.event.extendedProps.status}
            `);
        },
        eventDidMount: function(info) {
            // 根據狀態設定不同的樣式
            if (info.event.extendedProps.status === '已歸還') {
                info.el.classList.add('fc-event-returned');
            } else {
                info.el.classList.add('fc-event-borrowed');
            }
        }
    });
    
    calendar.render();
    
    // 取得行事曆事件資料
    function getCalendarEvents() {
        return borrowRecords.map(record => ({
            title: `${record.itemName} (${record.borrower})`,
            start: record.borrowDate,
            end: record.returnDate,
            extendedProps: {
                borrower: record.borrower,
                status: record.status,
                returnDate: record.returnDate
            }
        }));
    }
    
    // 匯入按鈕事件
    importButton.addEventListener('click', function(e) {
        e.preventDefault();
        importDialog.style.display = 'block';
        console.log('開啟對話框');
    });
    
    // Excel 檔案處理
    excelFile.addEventListener('change', function(e) {
        const file = e.target.files[0];
        if (file) {
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    
                    // 獲取第一個工作表
                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    
                    // 將工作表轉換為陣列
                    const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
                    
                    // 提取物品名稱（假設在第一欄）
                    excelItems = jsonData
                        .map(row => row[0]) // 獲取第一欄
                        .filter(item => item && typeof item === 'string') // 過濾空值
                        .map(item => item.trim()); // 清理空白
                    
                    // 顯示預覽
                    showPreview(excelItems);
                    
                } catch (error) {
                    console.error('Excel 檔案處理錯誤:', error);
                    alert('Excel 檔案處理發生錯誤，請確認檔案格式是否正確。');
                }
            };
            reader.readAsArrayBuffer(file);
        }
    });
    
    // 顯示預覽清單
    function showPreview(items) {
        previewList.innerHTML = items.length > 0 
            ? `<ul>${items.map(item => `<li>${item}</li>`).join('')}</ul>`
            : '<p>無有效的物品資料</p>';
    }
    
    // 修改確認匯入事件
    confirmImport.addEventListener('click', function() {
        if (excelItems.length > 0) {
            // 合併新舊物品並去除重複
            // 合���新舊物品並去除重複
            items = [...new Set([...items, ...excelItems])];
            
            // 儲存到 localStorage
            localStorage.setItem('items', JSON.stringify(items));
            
            // 更新下拉選單
            updateItemsList();
            
            // 重置
            excelItems = [];
            excelFile.value = '';
            previewList.innerHTML = '';
            importDialog.style.display = 'none';
            
            alert('物品清單匯入成功！');
        } else {
            alert('請先選擇 Excel 檔案');
        }
    });
    
    // 取消匯入事件
    cancelImport.addEventListener('click', function() {
        importDialog.style.display = 'none';
        excelFile.value = '';
        itemsList.value = '';
    });
    
    // 點擊對話框外部關閉
    window.addEventListener('click', function(event) {
        if (event.target === importDialog) {
            importDialog.style.display = 'none';
            excelFile.value = '';
            itemsList.value = '';
        }
    });
    
    // 表單提交事件
    borrowForm.addEventListener('submit', function(e) {
        e.preventDefault();
        
        const newRecord = {
            id: Date.now(),
            itemName: document.getElementById('itemName').value,
            borrower: document.getElementById('borrower').value,
            borrowDate: document.getElementById('borrowDate').value,
            returnDate: document.getElementById('returnDate').value,
            status: '借出中'  // 添加初始狀態
        };
        
        borrowRecords.push(newRecord);
        localStorage.setItem('borrowRecords', JSON.stringify(borrowRecords));
        
        displayBorrowRecords();
        calendar.removeAllEvents();
        calendar.addEventSource(getCalendarEvents());
        
        borrowForm.reset();
    });
    
    // 更新物品列表
    function updateItemsList() {
        console.log('更新物品下拉選單');
        itemSelect.innerHTML = '<option value="">請選擇物品</option>';
        items.sort().forEach(item => {
            const option = document.createElement('option');
            option.value = item;
            option.textContent = item;
            itemSelect.appendChild(option);
        });
    }
    
    // 顯示借用記錄
    function displayBorrowRecords() {
        borrowList.innerHTML = '';
        
        borrowRecords.forEach(record => {
            const row = document.createElement('tr');
            const statusClass = record.status === '已歸還' ? 'status-returned' : 'status-borrowed';
            
            row.innerHTML = `
                <td>${record.itemName}</td>
                <td>${record.borrower}</td>
                <td>${formatDateTime(record.borrowDate)}</td>
                <td>${record.status === '已歸還' ? 
                    formatDateTime(record.actualReturnDate) : 
                    formatDateTime(record.returnDate)}</td>
                <td class="${statusClass}">${record.status}</td>
                <td>
                    ${record.status === '借出中' ? 
                        `<button class="button-return" onclick="returnItem(${record.id})">歸還</button>` : 
                        ''}
                    <button class="button-delete" onclick="deleteRecord(${record.id})">刪除</button>
                </td>
            `;
            borrowList.appendChild(row);
        });
    }
    
    // 格式化日期時間
    function formatDateTime(dateTimeString) {
        return new Date(dateTimeString).toLocaleString('zh-TW');
    }
    
    // 刪除記錄
    window.deleteRecord = function(id) {
        if (confirm('確定要刪除這筆記錄嗎？')) {
            borrowRecords = borrowRecords.filter(record => record.id !== id);
            localStorage.setItem('borrowRecords', JSON.stringify(borrowRecords));
            displayBorrowRecords();
            
            // 更新行事曆
            calendar.removeAllEvents();
            calendar.addEventSource(getCalendarEvents());
        }
    }
    
    // 修改歸還功能
    window.returnItem = function(id) {
        currentReturnId = id;
        // 設定預計歸還時間為當前時間
        const now = new Date();
        now.setMinutes(now.getMinutes() - now.getTimezoneOffset());
        actualReturnDate.value = now.toISOString().slice(0, 16);
        // 顯示歸還對話框
        returnDialog.style.display = 'block';
    };

    // 確認歸還
    confirmReturn.addEventListener('click', function() {
        if (currentReturnId && actualReturnDate.value) {
            const recordIndex = borrowRecords.findIndex(record => record.id === currentReturnId);
            if (recordIndex !== -1) {
                // 更新記錄
                borrowRecords[recordIndex].status = '已歸還';
                borrowRecords[recordIndex].actualReturnDate = actualReturnDate.value;
                
                // 儲存到 localStorage
                localStorage.setItem('borrowRecords', JSON.stringify(borrowRecords));
                
                // 更新顯示
                displayBorrowRecords();
                
                // 更新行事曆
                calendar.removeAllEvents();
                calendar.addEventSource(getCalendarEvents());
                
                // 重置和關閉對話框
                currentReturnId = null;
                returnDialog.style.display = 'none';
                actualReturnDate.value = '';
            }
        }
    });

    // 取消歸還
    cancelReturn.addEventListener('click', function() {
        currentReturnId = null;
        returnDialog.style.display = 'none';
        actualReturnDate.value = '';
    });

    // 點擊對話框外部關閉
    window.addEventListener('click', function(event) {
        if (event.target === returnDialog) {
            currentReturnId = null;
            returnDialog.style.display = 'none';
            actualReturnDate.value = '';
        }
    });
});