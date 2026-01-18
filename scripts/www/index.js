// ==================== Tauri API 封装 ====================
let tauriAvailable = false;
let invokeFunc = null;

// 检测Tauri环境
if (window.__TAURI__) {
    tauriAvailable = true;
    invokeFunc = window.__TAURI__.tauri.invoke;
    console.log('Tauri环境已检测到');
} else {
    console.warn('未检测到Tauri环境，将在浏览器中运行');
    // 浏览器模拟函数
    invokeFunc = async (cmd, args) => {
        console.log(`模拟调用: ${cmd}`, args);
        if (cmd === 'save_excel_file') {
            return new Promise((resolve, reject) => {
                setTimeout(() => {
                    alert('在Tauri环境中才能保存文件');
                    reject('需要Tauri环境');
                }, 100);
            });
        }
        if (cmd === 'import_excel_file') {
            // 浏览器中模拟文件选择
            return new Promise((resolve, reject) => {
                const input = document.createElement('input');
                input.type = 'file';
                input.accept = '.xlsx,.xls';
                input.onchange = (e) => {
                    const file = e.target.files[0];
                    if (!file) {
                        reject('未选择文件');
                        return;
                    }
                    const reader = new FileReader();
                    reader.onload = (e) => {
                        // 这里简化处理，实际应调用原有解析逻辑
                        alert('在浏览器中只能预览，请使用Tauri APP版获得完整功能');
                        resolve({ success: false, message: '请在APP中使用此功能' });
                    };
                    reader.readAsArrayBuffer(file);
                };
                input.click();
            });
        }
    };
}

// ==================== 核心功能重写 ====================

// 1. 重写导出Excel功能
function setupExportExcel() {
    const exportBtn = document.getElementById('exportExcel');
    if (!exportBtn) return;
    
    exportBtn.addEventListener('click', async function() {
        try {
            // 获取当前数据
            const chargingRecords = JSON.parse(localStorage.getItem('chargingRecords') || '[]');
            const parkingRecords = JSON.parse(localStorage.getItem('parkingRecords') || '[]');
            
            if (chargingRecords.length === 0 && parkingRecords.length === 0) {
                alert('没有数据可以导出');
                return;
            }
            
            // 创建Excel工作簿
            const wb = XLSX.utils.book_new();
            
            // 充电记录工作表
            if (chargingRecords.length > 0) {
                const chargingData = chargingRecords.map(record => ({
                    '日期': record.date ? new Date(record.date).toLocaleDateString('zh-CN') : '',
                    '里程(公里)': record.mileage || 0,
                    '充电量(度)': record.amount || 0,
                    '电费单价(元/度)': record.price || 0,
                    '充电费用(元)': record.cost || 0,
                    '是否充满': record.isFull ? '是' : '否'
                }));
                const chargingWs = XLSX.utils.json_to_sheet(chargingData);
                XLSX.utils.book_append_sheet(wb, chargingWs, '充电记录');
            }
            
            // 停车记录工作表
            if (parkingRecords.length > 0) {
                const parkingData = parkingRecords.map(record => ({
                    '日期': record.date ? new Date(record.date).toLocaleDateString('zh-CN') : '',
                    '停车费用(元)': record.cost || 0
                }));
                const parkingWs = XLSX.utils.json_to_sheet(parkingData);
                XLSX.utils.book_append_sheet(wb, parkingWs, '停车记录');
            }
            
            // 生成Excel二进制数据
            const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
            
            // 转换为Base64
            const base64Data = arrayBufferToBase64(excelBuffer);
            
            // 调用Tauri保存文件
            if (tauriAvailable) {
                await invokeFunc('save_excel_file', { 
                    excelData: base64Data,
                    fileName: `汽车费用记录_${new Date().toISOString().slice(0,10).replace(/-/g, '')}.xlsx`
                });
                alert('导出成功！文件已保存到您选择的位置。');
                updateExportInfo();
            } else {
                // 浏览器中直接下载
                XLSX.writeFile(wb, `汽车费用记录_${new Date().toISOString().slice(0,10)}.xlsx`);
                alert('导出成功！文件已下载。');
            }
            
        } catch (error) {
            console.error('导出失败:', error);
            if (error !== '用户取消了保存') {
                alert('导出失败: ' + error);
            }
        }
    });
}

// 2. 重写导入Excel功能
function setupImportExcel() {
    const importBtn = document.getElementById('importBtn');
    const fileInput = document.getElementById('importExcel');
    
    if (!importBtn) return;
    
    // 隐藏原来的文件输入
    if (fileInput) fileInput.style.display = 'none';
    
    importBtn.addEventListener('click', async function() {
        try {
            if (!tauriAvailable) {
                // 浏览器中回退到原文件输入
                if (fileInput) {
                    fileInput.click();
                    return;
                }
                alert('请在APP中使用完整导入功能');
                return;
            }
            
            // 调用Tauri导入功能
            const result = await invokeFunc('import_excel_file', {});
            
            if (result.success) {
                // 解析并处理数据
                const chargingRecords = [];
                const parkingRecords = [];
                
                // 处理充电记录
                if (result.chargingData && result.chargingData.length > 0) {
                    result.chargingData.forEach(item => {
                        // 解析日期
                        let dateStr = item['日期'] || '';
                        let dateObj = null;
                        
                        if (dateStr) {
                            // 尝试多种日期格式
                            const dateMatch = dateStr.match(/(\d{4})[年/-](\d{1,2})[月/-](\d{1,2})/);
                            if (dateMatch) {
                                dateObj = new Date(dateMatch[1], dateMatch[2]-1, dateMatch[3]);
                            } else {
                                dateObj = new Date(dateStr);
                            }
                        }
                        
                        if (isNaN(dateObj?.getTime())) {
                            dateObj = new Date();
                        }
                        
                        chargingRecords.push({
                            id: `import_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
                            date: dateObj.toISOString().split('T')[0],
                            mileage: parseFloat(item['里程(公里)']) || parseInt(item['里程']) || 0,
                            amount: parseFloat(item['充电量(度)']) || parseFloat(item['充电量']) || 0,
                            price: parseFloat(item['电费单价(元/度)']) || parseFloat(item['电费单价']) || 0,
                            cost: parseFloat(item['充电费用(元)']) || parseFloat(item['本次充电总费用']) || 0,
                            isFull: item['是否充满'] === '是' || item['是否充满'] === true || false
                        });
                    });
                }
                
                // 处理停车记录
                if (result.parkingData && result.parkingData.length > 0) {
                    result.parkingData.forEach(item => {
                        let dateStr = item['日期'] || '';
                        let dateObj = null;
                        
                        if (dateStr) {
                            const dateMatch = dateStr.match(/(\d{4})[年/-](\d{1,2})[月/-](\d{1,2})/);
                            if (dateMatch) {
                                dateObj = new Date(dateMatch[1], dateMatch[2]-1, dateMatch[3]);
                            } else {
                                dateObj = new Date(dateStr);
                            }
                        }
                        
                        if (isNaN(dateObj?.getTime())) {
                            dateObj = new Date();
                        }
                        
                        parkingRecords.push({
                            id: `import_parking_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
                            date: dateObj.toISOString().split('T')[0],
                            cost: parseFloat(item['停车费用(元)']) || parseFloat(item['停车费用']) || 0
                        });
                    });
                }
                
                // 合并数据
                if (chargingRecords.length > 0) {
                    const existingCharging = JSON.parse(localStorage.getItem('chargingRecords') || '[]');
                    const newCharging = [...existingCharging, ...chargingRecords];
                    localStorage.setItem('chargingRecords', JSON.stringify(newCharging));
                }
                
                if (parkingRecords.length > 0) {
                    const existingParking = JSON.parse(localStorage.getItem('parkingRecords') || '[]');
                    const newParking = [...existingParking, ...parkingRecords];
                    localStorage.setItem('parkingRecords', JSON.stringify(newParking));
                }
                
                const totalImported = chargingRecords.length + parkingRecords.length;
                alert(`导入成功！\n充电记录: ${chargingRecords.length} 条\n停车记录: ${parkingRecords.length} 条\n总计: ${totalImported} 条`);
                
                // 刷新页面显示
                if (window.app && typeof window.app.render === 'function') {
                    window.app.render();
                } else {
                    location.reload();
                }
                
            } else {
                throw new Error(result.message || '导入失败');
            }
            
        } catch (error) {
            console.error('导入失败:', error);
            if (error.message !== '用户取消了选择') {
                alert('导入失败: ' + error.message);
            }
        }
    });
}

// 3. 数据清零功能
function setupClearData() {
    const clearBtn = document.getElementById('clearAllData');
    if (!clearBtn) return;
    
    clearBtn.addEventListener('click', async function(e) {
        e.preventDefault();
        
        const confirmed = confirm('⚠️ 警告！确定要清空所有数据吗？此操作不可恢复！');
        if (!confirmed) return;
        
        // 清空本地存储
        localStorage.removeItem('chargingRecords');
        localStorage.removeItem('parkingRecords');
        localStorage.removeItem('lastExportTime');
        
        alert('所有数据已成功清空！');
        location.reload();
    });
}

// 4. 更新导出信息
function updateExportInfo() {
    const lastExportTimeStr = localStorage.getItem('lastExportTime');
    const lastExportTimeEl = document.getElementById('lastExportTime');
    const daysSinceExportEl = document.getElementById('daysSinceExport');
    
    if (!lastExportTimeEl || !daysSinceExportEl) return;
    
    if (lastExportTimeStr) {
        const lastExportTime = new Date(lastExportTimeStr);
        const now = new Date();
        
        // 格式化显示
        const formattedDate = `${lastExportTime.getFullYear()}年${(lastExportTime.getMonth()+1).toString().padStart(2, '0')}月${lastExportTime.getDate().toString().padStart(2, '0')}日`;
        lastExportTimeEl.textContent = formattedDate;
        
        // 计算天数差
        lastExportTime.setHours(0, 0, 0, 0);
        now.setHours(0, 0, 0, 0);
        const diffDays = Math.floor((now - lastExportTime) / (1000 * 60 * 60 * 24));
        daysSinceExportEl.textContent = diffDays;
    } else {
        lastExportTimeEl.textContent = '从未导出';
        daysSinceExportEl.textContent = '0';
    }
}

// ==================== 工具函数 ====================

// ArrayBuffer转Base64
function arrayBufferToBase64(buffer) {
    let binary = '';
    const bytes = new Uint8Array(buffer);
    for (let i = 0; i < bytes.byteLength; i++) {
        binary += String.fromCharCode(bytes[i]);
    }
    return btoa(binary);
}

// Base64转ArrayBuffer
function base64ToArrayBuffer(base64) {
    const binaryString = atob(base64);
    const len = binaryString.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) {
        bytes[i] = binaryString.charCodeAt(i);
    }
    return bytes.buffer;
}

// ==================== 应用初始化 ====================

// 初始化应用
window.initializeApp = function() {
    console.log('初始化汽车费用追踪应用...');
    
    // 设置导出信息
    updateExportInfo();
    
    // 设置事件监听
    setupExportExcel();
    setupImportExcel();
    setupClearData();
    
    // 初始化时间选择器
    const timeBtns = document.querySelectorAll('#dashboard .time-selector [data-time]');
    timeBtns.forEach(btn => {
        btn.addEventListener('click', function() {
            timeBtns.forEach(b => b.classList.remove('active'));
            this.classList.add('active');
            if (window.app && typeof window.app.switchTimeRange === 'function') {
                window.app.switchTimeRange(this.dataset.time);
            }
        });
    });
    
    // 初始化标签页切换
    const navTabs = document.querySelectorAll('.nav-tab');
    navTabs.forEach(tab => {
        tab.addEventListener('click', function() {
            const tabId = this.dataset.tab;
            
            // 更新标签页
            navTabs.forEach(t => t.classList.remove('active'));
            this.classList.add('active');
            
            // 更新内容
            document.querySelectorAll('.tab-content').forEach(content => {
                content.classList.remove('active');
            });
            document.getElementById(tabId).classList.add('active');
            
            // 如果是统计页面，渲染图表
            if (tabId === 'statistics' && window.app && typeof window.app.renderCharts === 'function') {
                setTimeout(() => window.app.renderCharts(), 100);
            }
        });
    });
    
    console.log('应用初始化完成');
};

// 全局app对象（兼容原有代码）
window.app = window.app || {
    switchTimeRange: function(timeRange) {
        console.log('切换时间范围:', timeRange);
        // 这里可以添加刷新数据的逻辑
    },
    renderCharts: function() {
        console.log('渲染图表');
        // 这里可以添加图表渲染逻辑
    },
    render: function() {
        console.log('重新渲染页面');
        updateExportInfo();
        // 这里可以添加其他渲染逻辑
    }
};