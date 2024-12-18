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
    
    // 確保所有元素都存在
    if (!importDialog || !excelFile || !previewList || !confirmImport || !cancelImport) {
        console.error('找不到必要的 DOM 元素:', {
            importDialog: !!importDialog,
            excelFile: !!excelFile,
            previewList: !!previewList,
            confirmImport: !!confirmImport,
            cancelImport: !!cancelImport
        });
        return;
    }
    
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
            const returnDateText = info.event.extendedProps.status === '已歸還' 
                ? `實際歸還日期：${formatDateTime(info.event.extendedProps.actualReturnDate)}`
                : `預計歸還日期：${formatDateTime(info.event.extendedProps.returnDate)}`;

            alert(`
                物品：${info.event.title}
                借用人：${info.event.extendedProps.borrower}
                借用日期：${formatDateTime(info.event.start)}
                ${returnDateText}
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
            end: record.status === '已歸還' ? record.actualReturnDate : record.returnDate,
            extendedProps: {
                borrower: record.borrower,
                status: record.status,
                returnDate: record.returnDate,
                actualReturnDate: record.actualReturnDate
            }
        }));
    }
    
    // 匯入按鈕事件
    importButton.addEventListener('click', function(e) {
        e.preventDefault();
        importDialog.style.display = 'block';
        console.log('開啟對話框');
    });
    
    let excelItems = []; // 用於存儲 Excel 匯入的項目
    
    // Excel 檔案處理
    excelFile.addEventListener('change', function(e) {
        const file = e.target.files[0];
        if (!file) {
            previewList.innerHTML = '<p>未選擇檔案</p>';
            return;
        }

        console.log('檔案資訊:', {
            name: file.name,
            type: file.type,
            size: file.size
        });

        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                const binaryString = e.target.result;
                const workbook = XLSX.read(binaryString, {
                    type: 'binary',
                    cellDates: true,
                    cellNF: false,
                    cellText: false
                });
                
                console.log('工作表名稱:', workbook.SheetNames);
                
                // 獲取第一個工作表
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                
                // 將工作表轉換為陣列
                const jsonData = XLSX.utils.sheet_to_json(firstSheet, { 
                    header: 1,
                    raw: true,
                    blankrows: false
                });
                
                console.log('原始數據:', jsonData);
                
                // 提取物品名稱（第一欄）
                excelItems = jsonData
                    .filter(row => Array.isArray(row) && row.length > 0) // 確保是有效的行
                    .map(row => {
                        const value = row[0];
                        // 處理不同類型的值
                        if (typeof value === 'string') return value.trim();
                        if (typeof value === 'number') return String(value);
                        if (value && typeof value === 'object' && value.toString) return value.toString().trim();
                        return null;
                    })
                    .filter(item => item && item !== ''); // 過濾空值
                
                console.log('處理後的物品列表:', excelItems);
                
                if (excelItems.length === 0) {
                    throw new Error('沒有找到有效的物品資料');
                }
                
                // 更新預覽前先檢查元素
                if (previewList) {
                    showPreview(excelItems);
                } else {
                    console.error('預覽區域元素不存在');
                }
                
            } catch (error) {
                console.error('詳細錯誤:', error);
                if (previewList) {
                    previewList.innerHTML = `
                        <div class="preview-error">
                            <p>處理檔案時發生錯誤</p>
                            <p>錯誤訊息: ${error.message}</p>
                        </div>
                    `;
                }
                excelFile.value = '';
            }
        };

        reader.onerror = function(error) {
            console.error('檔案讀取錯誤:', error);
            if (previewList) {
                previewList.innerHTML = '<div class="preview-error">檔案讀取失敗，請重試</div>';
            }
        };

        reader.readAsBinaryString(file);
    });
    
    // 顯示預覽清單
    function showPreview(items) {
        if (!previewList) {
            console.error('預覽區域元素不存在');
            return;
        }

        if (items && items.length > 0) {
            const listHtml = items
                .map((item, index) => `<li>${index + 1}. ${item}</li>`)
                .join('');
            previewList.innerHTML = `
                <div class="preview-header">找到 ${items.length} 個物品：</div>
                <ul class="preview-list">${listHtml}</ul>
            `;
        } else {
            previewList.innerHTML = `
                <div class="preview-error">
                    <p>無有效的物品資料</p>
                    <p>請確認：</p>
                    <ul>
                        <li>Excel 檔案中的物品名稱是否在第一欄（A欄）</li>
                        <li>檔案是否為有效的 Excel 格式（.xls 或 .xlsx）</li>
                        <li>工作表中是否包含資料</li>
                    </ul>
                </div>
            `;
        }
    }
    
    // 確認匯入按鈕事件
    confirmImport.addEventListener('click', function() {
        if (excelItems && excelItems.length > 0) {
            try {
                // 獲取現有物品清單
                let items = JSON.parse(localStorage.getItem('items')) || [];
                console.log('現有物品:', items);
                
                // 合併新舊物品並去除重複
                items = [...new Set([...items, ...excelItems])];
                console.log('更��後的物品:', items);
                
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
            } catch (error) {
                console.error('儲存物品清單時發生錯誤:', error);
                alert('儲存失敗，請重試。');
            }
        } else {
            alert('請先選擇 Excel 檔案並確認有有效的物品資料');
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
        try {
            const items = JSON.parse(localStorage.getItem('items')) || [];
            const itemSelect = document.getElementById('itemName');
            
            // 清空現有選項
            itemSelect.innerHTML = '<option value="">請選擇物品</option>';
            
            // 添加新選項
            items.sort().forEach(item => {
                const option = document.createElement('option');
                option.value = item;
                option.textContent = item;
                itemSelect.appendChild(option);
            });
            
            console.log('下拉選單更新完成，共', items.length, '個物品');
        } catch (error) {
            console.error('更新下拉選單時發生錯誤:', error);
        }
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
                <td>${formatDateTime(record.returnDate)}</td>
                <td>${record.actualReturnDate ? formatDateTime(record.actualReturnDate) : '-'}</td>
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
        const recordIndex = borrowRecords.findIndex(record => record.id === id);
        if (recordIndex !== -1) {
            // 更新記錄狀態
            borrowRecords[recordIndex].status = '已歸還';
            borrowRecords[recordIndex].actualReturnDate = document.getElementById('actualReturnDate').value;
            
            // 儲存到 localStorage
            localStorage.setItem('borrowRecords', JSON.stringify(borrowRecords));
            
            // 更新顯示
            displayBorrowRecords();
            
            // 更新行事曆
            calendar.removeAllEvents();
            calendar.addEventSource(getCalendarEvents());
            
            // 關閉對話框
            returnDialog.style.display = 'none';
            
            console.log('物品已歸還:', borrowRecords[recordIndex]);
        }
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