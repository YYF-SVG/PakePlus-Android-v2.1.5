// index.js - 前端调用Tauri后端
const { invoke, dialog, fs } = window.__TAURI__;

// 1. 替换原“导出Excel”按钮的点击事件
document.addEventListener('DOMContentLoaded', function() {
    const exportBtn = document.getElementById('exportExcel');
    if (exportBtn) {
        exportBtn.addEventListener('click', async function() {
            try {
                // 调用你原来的数据生成函数，生成Excel数据 (ArrayBuffer)
                const wb = window.generateExcelWorkbook(); // 你需要确保这个函数存在并返回XLSX的workbook对象
                const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
                
                // 将ArrayBuffer转换为Base64，方便传输
                const base64Data = arrayBufferToBase64(excelBuffer);
                
                // 调用Tauri后端命令保存文件
                await invoke('save_excel_file', { data: base64Data });
                alert('导出成功！文件已保存至您选择的位置。');
            } catch (error) {
                console.error('导出失败:', error);
                alert('导出失败: ' + error);
            }
        });
    }

    // 2. 替换原“导入Excel”按钮的点击事件
    const importBtn = document.getElementById('importBtn');
    if (importBtn) {
        importBtn.addEventListener('click', async function() {
            try {
                // 调用Tauri后端命令选择并读取文件
                const fileData = await invoke('open_excel_file');
                // 假设后端返回Base64，进行解析
                const arrayBuffer = base64ToArrayBuffer(fileData);
                const workbook = XLSX.read(arrayBuffer, { type: 'array' });
                
                // 调用你原来的数据解析和导入函数
                window.parseExcelWorkbook(workbook); // 你需要确保这个函数存在
                alert('导入成功！');
                // 重新渲染页面
                if (window.app && window.app.render) {
                    window.app.render();
                }
            } catch (error) {
                console.error('导入失败:', error);
                if (error !== '用户取消') {
                    alert('导入失败: ' + error);
                }
            }
        });
    }

    // 3. 替换“数据清零”按钮的二次确认（使用原生对话框）
    const clearBtn = document.getElementById('clearAllData');
    if (clearBtn && window.__TAURI__) {
        clearBtn.addEventListener('click', async function(e) {
            e.preventDefault();
            const confirmed = await dialog.confirm('确定要清空所有数据吗？此操作不可恢复！', { title: '警告', type: 'warning' });
            if (confirmed) {
                // 执行清空操作
                localStorage.clear();
                alert('所有数据已成功清空！');
                location.reload();
            }
        });
    }
});

// 工具函数：ArrayBuffer 转 Base64
function arrayBufferToBase64(buffer) {
    let binary = '';
    const bytes = new Uint8Array(buffer);
    const len = bytes.byteLength;
    for (let i = 0; i < len; i++) {
        binary += String.fromCharCode(bytes[i]);
    }
    return btoa(binary);
}

// 工具函数：Base64 转 ArrayBuffer
function base64ToArrayBuffer(base64) {
    const binaryString = atob(base64);
    const len = binaryString.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) {
        bytes[i] = binaryString.charCodeAt(i);
    }
    return bytes.buffer;
}