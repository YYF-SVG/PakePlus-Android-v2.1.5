// PakePlus 文件存储工具类
class PakePlusStorage {
    static STORAGE_FILE_NAME = 'car_expense_data.json';
    static EXPORT_FILE_NAME = '车辆费用记录_';
    static EXPORT_FILE_EXT = '.xlsx';
    
    // 获取应用数据目录
    static async getAppDataDir() {
        try {
            if (window.pakeplus && window.pakeplus.file) {
                const appDir = await window.pakeplus.file.getAppDataDir();
                return appDir;
            }
            return '';
        } catch (error) {
            console.error('获取应用数据目录失败:', error);
            return '';
        }
    }
    
    // 获取数据文件完整路径
    static async getDataFilePath() {
        const appDir = await this.getAppDataDir();
        if (!appDir) return '';
        
        // 根据不同平台构建路径
        if (process.platform === 'win32') {
            return `${appDir}\\${this.STORAGE_FILE_NAME}`;
        } else {
            return `${appDir}/${this.STORAGE_FILE_NAME}`;
        }
    }
    
    // 读取数据文件
    static async readDataFile() {
        try {
            const filePath = await this.getDataFilePath();
            if (!filePath) return null;
            
            const content = await window.pakeplus.file.readFile(filePath);
            return JSON.parse(content);
        } catch (error) {
            console.error('读取数据文件失败:', error);
            return null;
        }
    }
    
    // 写入数据文件
    static async writeDataFile(data) {
        try {
            const filePath = await this.getDataFilePath();
            if (!filePath) return false;
            
            const content = JSON.stringify(data, null, 2);
            await window.pakeplus.file.writeFile(filePath, content);
            return true;
        } catch (error) {
            console.error('写入数据文件失败:', error);
            return false;
        }
    }
    
    // 获取充电记录
    static async getChargingRecords() {
        const data = await this.readDataFile();
        return data?.chargingRecords || [];
    }
    
    // 保存充电记录
    static async saveChargingRecords(records) {
        const data = await this.readDataFile() || {};
        data.chargingRecords = records;
        return await this.writeDataFile(data);
    }
    
    // 获取停车记录
    static async getParkingRecords() {
        const data = await this.readDataFile();
        return data?.parkingRecords || [];
    }
    
    // 保存停车记录
    static async saveParkingRecords(records) {
        const data = await this.readDataFile() || {};
        data.parkingRecords = records;
        return await this.writeDataFile(data);
    }
    
    // 获取导出信息
    static async getLastExportTime() {
        const data = await this.readDataFile();
        return data?.lastExportTime || null;
    }
    
    // 保存导出时间
    static async saveLastExportTime(time) {
        const data = await this.readDataFile() || {};
        data.lastExportTime = time;
        return await this.writeDataFile(data);
    }
    
    // 清空所有数据
    static async clearAllData() {
        return await this.writeDataFile({});
    }
    
    // 导出Excel文件
    static async exportToExcel(chargingRecords, parkingRecords) {
        try {
            const now = new Date();
            const dateStr = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
            const fileName = `${this.EXPORT_FILE_NAME}${dateStr}${this.EXPORT_FILE_EXT}`;
            
            // 使用SheetJS生成Excel数据
            const wb = XLSX.utils.book_new();
            
            // 充电记录工作表
            const chargingData = chargingRecords.map(record => ({
                日期: Utils.formatDate(record.date),
                里程: Math.round(Utils.safeParseFloat(record.mileage)),
                充电量: Utils.safeParseFloat(record.amount).toFixed(2),
                电费单价: Utils.safeParseFloat(record.price).toFixed(2),
                本次充电总费用: Utils.safeParseFloat(record.cost).toFixed(2),
                是否充满: record.isFull ? '是' : '否',
                百公里电耗: ''
            }));
            
            const chargingWs = ExcelProcessor.createWorksheet(chargingData, [
                '日期', '里程', '充电量', '电费单价', '本次充电总费用', '是否充满', '百公里电耗'
            ]);
            XLSX.utils.book_append_sheet(wb, chargingWs, '充电记录');
            
            // 停车记录工作表
            const parkingData = parkingRecords.map(record => ({
                日期: Utils.formatDate(record.date),
                停车费用: record.cost
            }));
            
            const parkingWs = ExcelProcessor.createWorksheet(parkingData, ['日期', '停车费用']);
            XLSX.utils.book_append_sheet(wb, parkingWs, '停车记录');
            
            // 生成Excel文件内容
            const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
            const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            const content = await new Response(blob).arrayBuffer();
            
            // 使用PakePlus的saveFile API保存文件
            const saveResult = await window.pakeplus.file.saveFile({
                content: content,
                filename: fileName,
                mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            });
            
            if (saveResult.success) {
                await this.saveLastExportTime(now.toISOString());
                return true;
            }
            return false;
        } catch (error) {
            console.error('导出Excel失败:', error);
            return false;
        }
    }
    
    // 导入Excel文件
    static async importFromExcel() {
        try {
            // 使用PakePlus的openFile API选择文件
            const fileResult = await window.pakeplus.file.openFile({
                accept: '.xlsx, .xls'
            });
            
            if (!fileResult.success || !fileResult.content) {
                return null;
            }
            
            // 读取文件内容
            const content = fileResult.content;
            const data = new Uint8Array(content);
            
            // 使用SheetJS解析Excel文件
            const wb = XLSX.read(data, {
                type: 'array',
                cellDates: true,
                cellNF: false,
                cellText: true
            });
            
            // 调用原有的ExcelProcessor.importFromExcel逻辑来解析数据
            return await ExcelProcessor.parseExcelData(wb);
        } catch (error) {
            console.error('导入Excel失败:', error);
            throw error;
        }
    }
}