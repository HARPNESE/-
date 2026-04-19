// Excel模板配置
const excelConfig = {
    // 模板1配置
    nova: {
        startRow: 8, // 数据开始行
        columns: {
            item: 'A', // 序号列
            name: 'B', // 设备名称列
            brand: 'C', // 品牌列
            model: 'D', // 型号列
            quantity: 'E', // 数量列
            unit: 'F', // 单位列
            price: 'G', // 单价列
            total: 'H', // 金额列
            params: 'I', // 技术参数列
            remark: 'J' // 备注列
        }
    },
    // 模板配置
    haidong: {
        startRow: 8, // 数据开始行
        columns: {
            item: 'A', // 序号列
            name: 'B', // 设备名称列
            brand: 'C', // 品牌列
            model: 'D', // 型号列
            quantity: 'E', // 数量列
            unit: 'F', // 单位列
            price: 'G', // 单价列
            total: 'H', // 金额列
            params: 'I', // 技术参数列
            remark: 'J' // 备注列
        }
    }
};


// 设置模板1
async function setupNovaTemplate(worksheet, workbook) {
// 设置工作表名称
    worksheet.name = '报价单';
    
    // 设置列宽
    worksheet.columns = [
        { key: 'item', width: 6 },
        { key: 'name', width: 22 },
        { key: 'brand', width: 23 },
        { key: 'model', width: 35 },
        { key: 'quantity', width: 8 },
        { key: 'unit', width: 8 },
        { key: 'price', width: 10 },
        { key: 'total', width: 12 },
        { key: 'params', width: 50 },
        { key: 'remark', width: 12 },
        { key: 'empty', width: 4 } // K列空列
    ];
    
    // 设置行高
    worksheet.getRow(1).height = 45; // 标题行
    for (let i = 2; i <= 6; i++) { // 信息栏
        worksheet.getRow(i).height = 27.65;
    }
    worksheet.getRow(7).height = 25; // 表头行
    
    // 1. 标题区（第1行）
    worksheet.getCell('A1').value = '模板1产品报价单';
    worksheet.getCell('A1').font = { name: '隶书', size: 24, bold: true };
    worksheet.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.mergeCells('A1:J1');



    // 添加logo图片（使用 fetch 加载）
    try {
        console.log('开始加载logo图片...');
        const response = await fetch('logo1.png');
        if (!response.ok) {
            throw new Error(`logo图片加载失败: ${response.status} ${response.statusText}`);
        }
        const logoBuffer = await response.arrayBuffer();
        const uint8Array = new Uint8Array(logoBuffer);

        const imageId = workbook.addImage({
            buffer: uint8Array,
            extension: 'png',
        });
        console.log('logo图片添加到工作簿成功，ID:', imageId);

        worksheet.addImage(imageId, {
            // 锚定A1单元格，精准像素偏移
            tl: { 
                col: 0, 
                row: 0, 
            },
            // 图片大小：完全匹配你截图的跨列+行高
            ext: { 
                width: 250,  // 宽度：刚好覆盖A+B两列
                height: 55   // 高度：刚好填满第1行，下边缘对齐第2行上边框
            },
            // 核心：绝对悬浮，不随单元格移动
            editAs: 'absolute',
            floating: true,
            behindDocument: false,
        });
        console.log('logo图片添加成功');
    } catch (error) {
        console.log('添加logo失败:', error);
    }
    
    // 2. 基础信息区（第2-6行）
    // 左侧客户信息区
    const leftInfo = [
        '项目名称：',
        '客户名称：',
        '客户地址：',
        '联  系  人：',
        '联系电话：'
    ];
    
    // 右侧供应商信息区
    const rightInfo = [
        '供  应  商：',
        '地        址：',
        '报价日期：***年***月***日',
        '联  系  人：***',
        '联系电话：***'
    ];
    
    for (let i = 0; i < 5; i++) {
        const row = i + 2;
        // 左侧
        worksheet.getCell(`A${row}`).value = leftInfo[i];
        worksheet.mergeCells(`A${row}:E${row}`);
        // 右侧
        worksheet.getCell(`F${row}`).value = rightInfo[i];
        worksheet.mergeCells(`F${row}:J${row}`);
        
        // 设置样式
        const leftCell = worksheet.getCell(`A${row}`);
        const rightCell = worksheet.getCell(`F${row}`);
        leftCell.font = { name: '等线', size: 12, bold: true };
        rightCell.font = { name: '等线', size: 12, bold: true };
        leftCell.alignment = { horizontal: 'left', vertical: 'middle' };
        rightCell.alignment = { horizontal: 'left', vertical: 'middle' };
        leftCell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
        rightCell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
    }
    
    // 3. 表头区（第7行）
    const headers = ['序号', '设备名称', '品牌', '型号', '数量', '单位', '单价', '金额', '技术参数', '备注'];
    for (let i = 0; i < headers.length; i++) {
        const cell = worksheet.getCell(7, i + 1);
        cell.value = headers[i];
        cell.font = { name: '等线', size: 12, bold: true };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFD9D9D9' } // 白色，背景1，深色15%
        };
        cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
    }
}

// 设置模板2
async function setupHaidongTemplate(worksheet, workbook) {
    // 设置工作表名称
    worksheet.name = '报价单';
    
    // 设置列宽
    worksheet.columns = [
        { key: 'item', width: 6 },
        { key: 'name', width: 22 },
        { key: 'brand', width: 23 },
        { key: 'model', width: 35 },
        { key: 'quantity', width: 8 },
        { key: 'unit', width: 8 },
        { key: 'price', width: 10 },
        { key: 'total', width: 12 },
        { key: 'params', width: 50 },
        { key: 'remark', width: 12 },
        { key: 'empty', width: 4 } // K列空列
    ];
    
    // 设置行高
    worksheet.getRow(1).height = 45; // 标题行
    for (let i = 2; i <= 6; i++) { // 信息栏
        worksheet.getRow(i).height = 27.65;
    }
    worksheet.getRow(7).height = 25; // 表头行
    
    // 1. 标题区（第1行）
    worksheet.getCell('A1').value = '模板2产品报价单';
    worksheet.getCell('A1').font = { name: '隶书', size: 24, bold: true };
    worksheet.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.mergeCells('A1:J1');
    
    // 添加模板logo图片（使用 fetch 加载）
    try {
        console.log('开始加载模板logo图片...');
        const response = await fetch('logo2.png');
        if (!response.ok) {
            throw new Error(`图片加载失败: ${response.status} ${response.statusText}`);
        }
        const logoBuffer = await response.arrayBuffer();
        const uint8Array = new Uint8Array(logoBuffer);

        const imageId = workbook.addImage({
            buffer: uint8Array,
            extension: 'png',
        });
        console.log('模板logo图片添加到工作簿成功，ID:', imageId);

        worksheet.addImage(imageId, {
            // 锚定A1单元格，精准像素偏移
            tl: { 
                col: 0, 
                row: 0, 
            },
            // 图片大小：完全匹配你截图的跨列+行高
            ext: { 
                width: 250,  // 宽度：刚好覆盖A+B两列
                height: 55   // 高度：刚好填满第1行，下边缘对齐第2行上边框
            },
            // 核心：绝对悬浮，不随单元格移动
            editAs: 'absolute',
            floating: true,
            behindDocument: false,
        });
        console.log('模板logo图片添加成功');
    } catch (error) {
        console.log('添加模板logo失败:', error);
    }
    
    // 2. 基础信息区（第2-6行）
    // 左侧客户信息区
    const leftInfo = [
        '项目名称：',
        '客户名称：',
        '客户地址：',
        '联  系  人：',
        '联系电话：'
    ];
    
    // 右侧供应商信息区
    const rightInfo = [
        '供  应  商：',
        '地        址：',
        '报价日期：***年***月***日',
        '联  系  人：***',
        '联系电话：***'
    ];
    
    for (let i = 0; i < 5; i++) {
        const row = i + 2;
        // 左侧
        worksheet.getCell(`A${row}`).value = leftInfo[i];
        worksheet.mergeCells(`A${row}:E${row}`);
        // 右侧
        worksheet.getCell(`F${row}`).value = rightInfo[i];
        worksheet.mergeCells(`F${row}:J${row}`);
        
        // 设置样式
        const leftCell = worksheet.getCell(`A${row}`);
        const rightCell = worksheet.getCell(`F${row}`);
        leftCell.font = { name: '等线', size: 12, bold: true };
        rightCell.font = { name: '等线', size: 12, bold: true };
        leftCell.alignment = { horizontal: 'left', vertical: 'middle' };
        rightCell.alignment = { horizontal: 'left', vertical: 'middle' };
        leftCell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
        rightCell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
    }
    
    // 3. 表头区（第7行）
    const headers = ['序号', '设备名称', '品牌', '型号', '数量', '单位', '单价', '金额', '技术参数', '备注'];
    for (let i = 0; i < headers.length; i++) {
        const cell = worksheet.getCell(7, i + 1);
        cell.value = headers[i];
        cell.font = { name: '等线', size: 12, bold: true };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFD9D9D9' } // 白色，背景1，深色15%
        };
        cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
    }
}

// 将数字转换为中文大写金额
function numberToChinese(amount) {
    const digits = ['零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖'];
    const units = ['', '拾', '佰', '仟'];
    const bigUnits = ['', '万', '亿'];
    
    if (amount === 0) return '零元整';
    
    // 处理整数部分和小数部分
    const integerPart = Math.floor(amount);
    const decimalPart = Math.round((amount - integerPart) * 100);
    
    let result = '';
    let unitIndex = 0;
    let bigUnitIndex = 0;
    let temp = integerPart;
    
    // 处理整数部分
    while (temp > 0) {
        const section = temp % 10000;
        if (section > 0) {
            let sectionResult = '';
            let sectionTemp = section;
            let sectionUnitIndex = 0;
            
            while (sectionTemp > 0) {
                const digit = sectionTemp % 10;
                if (digit > 0) {
                    sectionResult = digits[digit] + units[sectionUnitIndex] + sectionResult;
                } else {
                    // 处理连续的零
                    if (sectionResult.length > 0 && sectionResult.charAt(0) !== '零') {
                        sectionResult = '零' + sectionResult;
                    }
                }
                sectionTemp = Math.floor(sectionTemp / 10);
                sectionUnitIndex++;
            }
            
            result = sectionResult + bigUnits[bigUnitIndex] + result;
        }
        temp = Math.floor(temp / 10000);
        bigUnitIndex++;
    }
    
    result += '元';
    
    // 处理小数部分
    if (decimalPart === 0) {
        result += '整';
    } else {
        const jiao = Math.floor(decimalPart / 10);
        const fen = decimalPart % 10;
        if (jiao > 0) {
            result += digits[jiao] + '角';
        }
        if (fen > 0) {
            result += digits[fen] + '分';
        }
    }
    
    return result;
}

// 填充数据到工作表
function fillWorksheetData(worksheet, currentTemplate, products) {
    console.log('进入 fillWorksheetData 函数');
    console.log('worksheet:', worksheet);
    console.log('currentTemplate:', currentTemplate);
    console.log('products:', products);
    
    // 填充数据
    let rowIndex = 8; // 数据开始行
    let groupItemIndex = 1;
    let groupTotals = [];
    
    const groups = document.querySelectorAll('.group');
    console.log('找到的分组数量:', groups.length);
    
    groups.forEach((group) => {
        console.log('处理分组:', group);
        const groupName = group.querySelector('.group-header input').value;
        console.log('分组名称:', groupName);
        let groupTotal = 0;
        
        // 4. 分类标题行
        const groupCell = worksheet.getCell('A' + rowIndex);
        groupCell.value = groupName;
        groupCell.font = { name: '等线', size: 11, bold: true };
        groupCell.alignment = { horizontal: 'left', vertical: 'middle' };
        worksheet.mergeCells('A' + rowIndex + ':J' + rowIndex);
        worksheet.getRow(rowIndex).height = 25; // 分类标题行高
        rowIndex++;
        
        // 5. 设备数据行
        const rows = group.querySelectorAll('tbody tr');
        console.log('分组内的行数:', rows.length);
        
        rows.forEach(row => {
            console.log('处理行:', row);
            const productId = row.querySelector('.product-id').value;
            console.log('productId:', productId);
            const price = parseFloat(row.querySelector('.price').value) || 0;
            const quantity = parseInt(row.querySelector('.quantity').value) || 0;
            const amount = price * quantity;
            
            // 获取用户在界面上输入的所有值（包括自定义修改的值）
            const name = row.querySelector('.name').value || '';
            const brand = row.querySelector('.brand').value || '';
            const modelElement = row.querySelector('.select-btn');
            const model = modelElement ? modelElement.textContent || '' : '';
            const params = row.querySelector('.params').value || '';
            const remark = row.querySelector('.remark').value || '';
            const unit = row.querySelector('.unit').value || '个';
            
            // 只要有型号或名称就导出（包括自定义设备）
            if (model || name) {
                // 填充产品信息（使用用户实际输入的值）
                worksheet.getCell('A' + rowIndex).value = groupItemIndex;
                worksheet.getCell('B' + rowIndex).value = name;
                worksheet.getCell('C' + rowIndex).value = brand;
                worksheet.getCell('D' + rowIndex).value = model;
                worksheet.getCell('E' + rowIndex).value = quantity;
                worksheet.getCell('F' + rowIndex).value = unit;
                worksheet.getCell('G' + rowIndex).value = price;
                worksheet.getCell('H' + rowIndex).value = amount;
                worksheet.getCell('I' + rowIndex).value = params;
                worksheet.getCell('J' + rowIndex).value = remark;
                
                // 设置样式
                for (let col = 1; col <= 10; col++) {
                    const cell = worksheet.getCell(rowIndex, col);
                    cell.font = { name: '等线', size: 11 };
                    if (col === 9 ) { // 技术参数列
                        cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
                    } else if (col === 10) { // 备注列
                        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                    } else {
                        cell.alignment = { horizontal: 'center', vertical: 'middle' };
                    }
                }
                
                // 设置数字格式
                worksheet.getCell('G' + rowIndex).numFmt = '#,##0.00'; // 单价
                worksheet.getCell('H' + rowIndex).numFmt = '#,##0.00'; // 金额
                
                groupTotal += amount;
                worksheet.getRow(rowIndex).height = 25; // 设备行高
                rowIndex++;
                groupItemIndex++;
            }
        });
        
        // 添加分组小计行
        if (groupTotal > 0) {
            // 填充小计信息（不合并单元格）
            worksheet.getCell('A' + rowIndex).value = ''; // 序号列留空
            worksheet.getCell('B' + rowIndex).value = '小计'; // 设备名称列写"小计"
            worksheet.getCell('C' + rowIndex).value = ''; // 品牌列留空
            worksheet.getCell('D' + rowIndex).value = ''; // 型号列留空
            worksheet.getCell('E' + rowIndex).value = ''; // 数量列留空
            worksheet.getCell('F' + rowIndex).value = ''; // 单位列留空
            worksheet.getCell('G' + rowIndex).value = ''; // 单价列留空
            worksheet.getCell('H' + rowIndex).value = groupTotal; // 金额列
            worksheet.getCell('I' + rowIndex).value = ''; // 技术参数列留空
            worksheet.getCell('J' + rowIndex).value = 'RMB（元）'; // 备注列写"RMB（元）"
            
            // 设置小计行样式
            for (let col = 1; col <= 10; col++) {
                const cell = worksheet.getCell(rowIndex, col);
                cell.font = { name: '等线', size: 11, bold: true };
                // 只修改"小计"为水平居左对齐，垂直居中
                if (col === 2) { // B列（设备名称列）
                    cell.alignment = { horizontal: 'left', vertical: 'middle' };
                } else {
                    cell.alignment = { horizontal: 'center', vertical: 'middle' };
                }
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                };
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            }
            
            // 设置小计金额格式
            worksheet.getCell('H' + rowIndex).numFmt = '#,##0.00';
            worksheet.getRow(rowIndex).height = 25; // 小计行高
            rowIndex++;
        }
        
        groupItemIndex = 1; // 重置分组内序号
        groupTotals.push(groupTotal);
    });
    
    // 6. 合计区
    const total = groupTotals.reduce((sum, value) => sum + value, 0);
    const totalRow = rowIndex;
    
    // 合计(大写):
    worksheet.getCell('A' + totalRow).value = '合计(大写):';
    worksheet.mergeCells('A' + totalRow + ':B' + totalRow);
    worksheet.getRow(totalRow).height = 28; // 合计行高
    
    // 大写总金额
    worksheet.getCell('C' + totalRow).value =  numberToChinese(total);
    worksheet.mergeCells('C' + totalRow + ':E' + totalRow);
    
    // 合计(小写）:
    worksheet.getCell('F' + totalRow).value = '合计(小写）:';
    worksheet.mergeCells('F' + totalRow + ':G' + totalRow);
    
    // 小写总金额
    worksheet.getCell('H' + totalRow).value = total;
    worksheet.getCell('H' + totalRow).numFmt = '#,##0.00'; // 合计（小写）
    worksheet.mergeCells('H' + totalRow + ':J' + totalRow);
    
    // 设置合计区样式
    for (let col = 1; col <= 10; col++) {
        const cell = worksheet.getCell(totalRow, col);
        cell.font = { name: '等线', size: 11, bold: true };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFD9D9D9' } // 白色，背景1，深色15%
        };
        cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
    }
    
    // 7. 备注区
    const remarkRow = totalRow + 1;
    worksheet.getCell('A' + remarkRow).value = '备注：\n1、本报价单为含税报价，包含13%增值税；\n2、本报价单有效期：30天。';
    worksheet.mergeCells('A' + remarkRow + ':J' + remarkRow);
    worksheet.getCell('A' + remarkRow).font = { name: '等线', size: 11, bold: true };
    worksheet.getCell('A' + remarkRow).alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
    worksheet.getRow(remarkRow).height = 60; // 备注行高
    
    // 8. 为所有单元格添加细边框
    const lastRow = remarkRow;
    for (let row = 1; row <= lastRow; row++) {
        for (let col = 1; col <= 10; col++) {
            const cell = worksheet.getCell(row, col);
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        }
    }
    
    return total;
}

