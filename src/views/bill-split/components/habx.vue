<template>
  <div class="bill-split-container">
    <div class="upload-section">
      <el-upload
        class="upload-dragger"
        drag
        :auto-upload="false"
        :on-change="handleFileChange"
        :before-upload="beforeUpload"
        accept=".xlsx,.xls"
        :show-file-list="false"
      >
        <el-icon class="el-icon--upload"><upload-filled /></el-icon>
        <div class="el-upload__text">
          将Excel文件拖到此处,或<em>点击上传</em>
        </div>
        <template #tip>
          <div class="el-upload__tip">
            只能上传 xlsx/xls 文件,且不超过 10MB
          </div>
        </template>
      </el-upload>
    </div>

    <!-- 数据展示区域 -->
    <div v-if="showData" class="data-section">
      <div class="data-header">
        <h3>华安保险账单数据</h3>
        <div class="header-buttons">
          <el-button
            type="success"
            :loading="generating"
            @click="generateExcelFiles"
            :disabled="!showData"
          >
            {{ generating ? "生成中..." : "生成ZIP包" }}
          </el-button>
        </div>
      </div>

      <div class="data-summary">

        <el-alert
          v-if="showData && getTotalRows() > 0"
          title="数据概览"
          type="info"
          :description="`已读取 ${allSheetData.length} 个工作表,共 ${getTotalRows()} 行数据`"
          show-icon
          style="margin-top: 10px"
        />
      </div>


      <!-- 部门拆分结果表格 -->
      <div v-if="generatedFiles.length > 0" class="department-results">
        <h4>账单部门拆分结果</h4>
        <el-table :data="generatedFiles" border style="width: 100%">
          <el-table-column prop="departmentName" label="部门名称" width="200" />
          <el-table-column prop="rowCount" label="数据行数" width="120" />
          <el-table-column prop="fileName" label="生成文件名" />
          <el-table-column label="类型" width="120">
            <template #default="scope">
              <el-tag
                :type="scope.row.departmentName.includes('火车') ? 'warning' :
                       scope.row.departmentName.includes('酒店') ? 'danger' : 'success'"
                size="small"
              >
                {{ scope.row.departmentName.includes('火车') ? '火车票' :
                   scope.row.departmentName.includes('酒店') ? '酒店' : '机票' }}
              </el-tag>
            </template>
          </el-table-column>
        </el-table>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref } from "vue";
import { ElMessage } from "element-plus";
import { UploadFilled } from "@element-plus/icons-vue";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import JSZip from "jszip";

defineOptions({
  name: "HabxBillSplit"
});

const uploadedFile = ref<File | null>(null);
const allSheetData = ref<any[]>([]);
const loading = ref(false);
const showData = ref(false);
const generating = ref(false);
const originalWorkbook = ref<any>(null);
const generatedFiles = ref<any[]>([]); // 记录生成的文件

// 智能列宽设置函数
const setSmartColumnWidths = (worksheet: any, headers: string[], data: any[] = [], tableType: string = '') => {
  console.log(`开始设置智能列宽，表格类型: ${tableType}`);
  console.log('表头列表:', headers);

  headers.forEach((header, columnIndex) => {
    const columnNumber = columnIndex + 1;
    let maxLength = header.toString().length; // 先以表头长度为基准

    // 遍历数据行，找到该列最长的内容
    data.forEach(row => {
      const cellValue = row[columnIndex] || '';
      const textLength = cellValue.toString().length;
      if (textLength > maxLength) {
        maxLength = textLength;
      }
    });

    // 根据内容长度设置列宽，使用系数调整
    let columnWidth;
    if (maxLength <= 5) {
      columnWidth = maxLength * 2.5; // 短内容使用较大系数
    } else if (maxLength <= 10) {
      columnWidth = maxLength * 2.0; // 中等内容
    } else if (maxLength <= 20) {
      columnWidth = maxLength * 1.5; // 较长内容使用较小系数
    } else {
      columnWidth = maxLength * 1.2; // 很长内容使用更小系数，并限制最大宽度
      columnWidth = Math.min(columnWidth, 50); // 最大宽度限制为50
    }

    // 为特定列增加额外的宽度系数
    const specialColumns = [
      '旅客直属部门', '应还款总金额', '酒店开票类型',
      '平均客房单价', '预订/退款日期', '座位编号'
    ];

    // 需要更大宽度系数的特殊列
    const extraWideColumns = ['出发时间'];
    // 火车票表格的行程列需要特别大的宽度
    if (tableType === 'train' && header === '行程') {
      columnWidth *= 2.5; // 火车票行程列增加150%的宽度
      console.log(`第${columnNumber}列"${header}"是火车票行程列，增加150%宽度，基础宽度：${columnWidth / 2.5}，最终宽度：${columnWidth}`);
    } else if (extraWideColumns.includes(header)) {
      columnWidth *= 1.8; // 为这些列增加80%的宽度
      console.log(`第${columnNumber}列"${header}"是需要大宽度的列，增加80%宽度，基础宽度：${columnWidth / 1.8}，最终宽度：${columnWidth}`);
    } else if (specialColumns.includes(header)) {
      columnWidth *= 1.3; // 为其他特殊列增加30%的宽度
      console.log(`第${columnNumber}列"${header}"是特殊列，增加30%宽度，基础宽度：${columnWidth / 1.3}，最终宽度：${columnWidth}`);
    } else {
      console.log(`第${columnNumber}列"${header}"使用标准宽度：${columnWidth}`);
    }

    // 设置最小宽度
    columnWidth = Math.max(columnWidth, 8);

    worksheet.getColumn(columnNumber).width = columnWidth;
    console.log(`第${columnNumber}列"${header}": 最长内容${maxLength}字符, 设置宽度${columnWidth.toFixed(1)}`);
  });
};

const handleFileChange = (uploadFile: any) => {
  const file = uploadFile.raw;
  if (!file) return;

  uploadedFile.value = file;
  readFile(file);
};

const readFile = (file: File) => {
  loading.value = true;

  const reader = new FileReader();
  reader.onload = e => {
    try {
      const buffer = e.target?.result as ArrayBuffer;
      const workbook = new ExcelJS.Workbook();

      workbook.xlsx
        .load(buffer)
        .then(async () => {
          console.log('=== Excel文件加载成功 ===');
          console.log('所有工作表:', workbook.worksheets.map(ws => ws.name));

          const sheetInfoArray: any[] = [];

          workbook.worksheets.forEach(worksheet => {
            // 读取数据为二维数组
            const jsonData: any[][] = [];
            worksheet.eachRow((row, rowNumber) => {
              const rowData: any[] = [];
              row.eachCell((cell, colNumber) => {
                rowData.push(cell.value);
              });
              jsonData.push(rowData);
            });

            sheetInfoArray.push({
              name: worksheet.name,
              rowCount: jsonData.length,
              columnCount: worksheet.columnCount,
              data: jsonData
            });

            console.log(`工作表 "${worksheet.name}": ${jsonData.length} 行, ${worksheet.columnCount} 列`);
          });

          allSheetData.value = sheetInfoArray;
          originalWorkbook.value = workbook;

          // 异步处理机票和火车票数据
          const processAllData = async () => {
            // 处理国内和国际机票按部门拆分
            const domesticResult = processDomesticFlights(workbook);
            const internationalResult = processInternationalFlights(workbook);
            const trainResult = processDomesticTrains(workbook);
            const hotelResult = processDomesticHotels(workbook);

            // 合并机票、火车票和酒店的数据
            const mergedDepartmentData: { [key: string]: Array<{data: any[], type: string}> } = {};
            const allDepartments = new Set([
              ...Object.keys(domesticResult.departmentData),
              ...Object.keys(internationalResult.departmentData),
              ...Object.keys(trainResult.departmentData),
              ...Object.keys(hotelResult.departmentData)
            ]);

            allDepartments.forEach(department => {
              mergedDepartmentData[department] = [];

              // 添加国内机票数据
              if (domesticResult.departmentData[department]) {
                domesticResult.departmentData[department].forEach(row => {
                  mergedDepartmentData[department].push({
                    data: row,
                    type: 'domestic-flight'
                  });
                });
              }

              // 添加国际机票数据
              if (internationalResult.departmentData[department]) {
                internationalResult.departmentData[department].forEach(row => {
                  mergedDepartmentData[department].push({
                    data: row,
                    type: 'international-flight'
                  });
                });
              }

              // 添加火车票数据
              if (trainResult.departmentData[department]) {
                trainResult.departmentData[department].forEach(row => {
                  mergedDepartmentData[department].push({
                    data: row,
                    type: 'train'
                  });
                });
              }

              // 添加酒店数据
              if (hotelResult.departmentData[department]) {
                hotelResult.departmentData[department].forEach(row => {
                  mergedDepartmentData[department].push({
                    data: row,
                    type: 'hotel'
                  });
                });
              }
            });

            // 生成合并后的部门报告
            const columnMappings = {
              domestic: domesticResult.columnMapping,
              international: internationalResult.columnMapping,
              train: trainResult.columnMapping,
              hotel: hotelResult.columnMapping
            };

            // 分别处理火车票、机票和酒店数据
            const departments = Object.keys(mergedDepartmentData);
            for (const dept of departments) {
              const deptData = mergedDepartmentData[dept];
              if (deptData.length > 0) {
                // 分离火车票、机票和酒店数据
                const trainData = deptData.filter(item => item.type === 'train');
                const flightData = deptData.filter(item => item.type === 'domestic-flight' || item.type === 'international-flight');
                const hotelData = deptData.filter(item => item.type === 'hotel');

                // 生成火车票报告
                if (trainData.length > 0) {
                  const trainRows = trainData.map(item => item.data);
                  await generateTrainDepartmentReport(dept, trainRows, columnMappings.train);
                }

                // 生成机票报告
                if (flightData.length > 0) {
                  await generateFlightDepartmentReport(dept, flightData, columnMappings);
                }

                // 生成酒店报告
                if (hotelData.length > 0) {
                  const hotelRows = hotelData.map(item => item.data);
                  // 重新获取酒店表头用于退订费用查找
                  const hotelSheet = workbook.getWorksheet('国内酒店') || workbook.getWorksheet('酒店') || workbook.getWorksheet('酒店票') || workbook.getWorksheet('国内酒店票');
                  let hotelHeaders = [];
                  if (hotelSheet) {
                    const headerRow = hotelSheet.getRow(1);
                    headerRow.eachCell({ includeEmpty: true }, (cell: any) => {
                      hotelHeaders.push(cell.value);
                    });
                  }
                  await generateHotelDepartmentReport(dept, hotelRows, columnMappings.hotel, hotelHeaders);
                }
              }
            }

            const totalProcessedRows = domesticResult.processedRows + internationalResult.processedRows + trainResult.processedRows + hotelResult.processedRows;
            ElMessage.success(`处理完成！共处理 ${totalProcessedRows} 行数据（国内机票 ${domesticResult.processedRows} 行，国际机票 ${internationalResult.processedRows} 行，国内火车 ${trainResult.processedRows} 行，国内酒店 ${hotelResult.processedRows} 行），分成 ${departments.length} 个部门`);
          };

          await processAllData();

          showData.value = true;
          loading.value = false;

          ElMessage.success(
            `成功读取 ${sheetInfoArray.length} 个工作表！数据处理完成。`
          );
        })
        .catch(error => {
          console.error("读取Excel文件失败:", error);
          ElMessage.error("读取Excel文件失败,请检查文件格式是否正确");
          loading.value = false;
        });
    } catch (error) {
      console.error("文件处理失败:", error);
      ElMessage.error("文件处理失败");
      loading.value = false;
    }
  };

  reader.readAsArrayBuffer(file);
};

const beforeUpload = (file: File) => {
  const isExcel = file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
                  file.type === 'application/vnd.ms-excel' ||
                  file.name.endsWith('.xlsx') ||
                  file.name.endsWith('.xls');

  if (!isExcel) {
    ElMessage.error('只能上传Excel文件！');
    return false;
  }

  const isLt10M = file.size / 1024 / 1024 < 10;
  if (!isLt10M) {
    ElMessage.error("文件大小不能超过10MB！");
    return false;
  }

  return true;
};

const getTotalRows = () => {
  return allSheetData.value.reduce((sum, sheet) => sum + sheet.rowCount, 0);
};

const previewSheet = (sheetInfo: any) => {
  console.log('查看工作表数据:', sheetInfo.name);
  console.log('数据内容:', sheetInfo.data);
  ElMessage.info(`工作表 "${sheetInfo.name}" 数据已输出到控制台`);
};

// 处理行程数据，提取出发地-目的地
const processItinerary = (itinerary: string): string => {
  if (!itinerary) return '';

  // 将行程转换为字符串
  const itineraryStr = itinerary.toString();

  // 匹配模式：出发地-目的地，后面可能跟着航班信息
  // 例如：海口-长春（航班HU448） -> 海口-长春
  // 例如：海口-长春 -> 海口-长春
  const match = itineraryStr.match(/^([^\-（]+)-([^\-（]+)/);
  if (match) {
    return `${match[1].trim()}-${match[2].trim()}`;
  }

  // 如果没有匹配到标准格式，尝试其他模式
  // 例如：海口-长春航班HU448 -> 海口-长春
  const altMatch = itineraryStr.match(/^([^\-]+)-([^\s]+)\s*/);
  if (altMatch) {
    return `${altMatch[1].trim()}-${altMatch[2].trim()}`;
  }

  // 如果都没有匹配到，返回原字符串
  return itineraryStr;
};

// 处理出发时间，合并日期和时间
const processDepartureTime = (date: any, time: any): string => {
  if (!date && !time) return '';

  const dateStr = date ? date.toString().trim() : '';
  const timeStr = time ? time.toString().trim() : '';

  if (!dateStr) return timeStr;
  if (!timeStr) return dateStr;

  // 如果日期已经包含时间信息，直接返回
  if (dateStr.includes(':') || dateStr.includes('时')) {
    return dateStr;
  }

  // 合并日期和时间
  return `${dateStr} ${timeStr}`;
};

// 转换英文姓名为中文姓名
const convertEnglishNameToChinese = (englishName: string): string => {
  if (!englishName) return '';

  const nameMap: { [key: string]: string } = {
    'ZHAO/QUAN': '赵权',
    'LIU/XIANGONG': '刘现功',
    'LI/GUANGRONG': '李光荣'
  };

  const upperName = englishName.toString().trim().toUpperCase();
  return nameMap[upperName] || englishName;
};



// 处理国际机票数据
const processInternationalFlights = (workbook: any) => {
  console.log('=== 开始处理国际机票数据 ===');

  // 获取国际机票工作表
  const flightSheet = workbook.getWorksheet('国际机票');
  if (!flightSheet) {
    console.log('未找到"国际机票"工作表');
    return { departmentData: {}, columnMapping: {}, processedRows: 0 };
  }

  console.log('找到"国际机票"工作表');

  // 读取国际机票数据（包含空单元格）
  const flightData: any[][] = [];
  flightSheet.eachRow((row: any, rowNumber: number) => {
    const rowData: any[] = [];
    row.eachCell({ includeEmpty: true }, (cell: any, colNumber: number) => {
      rowData.push(cell.value);
    });
    flightData.push(rowData);
  });

  console.log(`国际机票数据行数: ${flightData.length}`);

  if (flightData.length < 2) {
    console.log('国际机票数据行数不足，跳过处理');
    return { departmentData: {}, columnMapping: {}, processedRows: 0 };
  }

  // 获取表头
  const headers = flightData[0];
  console.log('国际机票表头:', headers);

  // 查找"费用归属"列的索引
  const costBelongIndex = headers.findIndex((h: any) =>
    h && h.toString().includes('费用归属')
  );

  if (costBelongIndex === -1) {
    console.error('未找到国际机票"费用归属"列');
    ElMessage.warning('国际机票工作表格式不正确，缺少"费用归属"列');
    return { departmentData: {}, columnMapping: {}, processedRows: 0 };
  }

  console.log(`国际机票费用归属列索引: ${costBelongIndex}`);

  // 查找所有需要映射的列（新表头 -> 原表头映射）
  const columnMapping: { [key: string]: number } = {
    '票号': headers.findIndex((h: any) => h && h.toString().includes('票号')),
    '机票状态': headers.findIndex((h: any) => h && h.toString().includes('订单状态')),
    '预订人': headers.findIndex((h: any) => h && h.toString().includes('预订人')),
    '旅客姓名': headers.findIndex((h: any) => h && h.toString().includes('乘机人')),
    '旅客直属部门': costBelongIndex, // 使用费用归属列
    '航程': headers.findIndex((h: any) => h && h.toString().includes('航程')),
    '航班号': headers.findIndex((h: any) => h && h.toString().includes('航班号')),
    '出发时间': headers.findIndex((h: any) => h && h.toString().includes('出发时间')),
    '票销售价': headers.findIndex((h: any) => h && h.toString().includes('票面价')),
    '税费': headers.findIndex((h: any) => h && h.toString().includes('税费')),
    '燃油费(国内)': headers.findIndex((h: any) => h && h.toString().includes('燃油')),
    '改签费': headers.findIndex((h: any) => h && h.toString().includes('改签费')),
    '退票费': headers.findIndex((h: any) => h && h.toString().includes('退票费')),
    '服务费': headers.findIndex((h: any) => h && h.toString().includes('系统使用费')),
    '应还款总金额': headers.findIndex((h: any) => h && h.toString().includes('总金额'))
  };

  console.log('国际机票列映射:', columnMapping);

  // 按部门分组数据
  const departmentData: { [key: string]: any[][] } = {};
  let processedRows = 0;

  // 从第二行开始处理数据（跳过表头）
  for (let i = 1; i < flightData.length; i++) {
    const row = flightData[i];
    const costBelong = row[costBelongIndex]?.toString().trim();

    if (!costBelong) {
      console.log(`国际机票第 ${i + 1} 行费用归属为空，跳过`);
      continue;
    }

    // 提取部门名称（去掉"商务-机票-"前缀）
    let departmentName = costBelong;
    if (costBelong.startsWith('商务-机票-')) {
      departmentName = costBelong.replace('商务-机票-', '');
    } else if (costBelong.includes('-机票-')) {
      // 处理可能的其他格式，如"其他-机票-部门名"
      const parts = costBelong.split('-机票-');
      if (parts.length === 2) {
        departmentName = parts[1]; // 取第二部分作为部门名
      }
    }

    console.log(`国际机票第 ${i + 1} 行: 费用归属="${costBelong}" -> 部门="${departmentName}"`);

    // 如果部门不存在，创建新的数组
    if (!departmentData[departmentName]) {
      departmentData[departmentName] = [];
    }

    // 将数据行添加到对应部门
    departmentData[departmentName].push(row);
    processedRows++;
  }

  console.log('国际机票部门分组结果:', departmentData);
  console.log(`处理了国际机票 ${processedRows} 行数据，分成了 ${Object.keys(departmentData).length} 个部门`);

  return { departmentData, columnMapping, processedRows };
};

// 处理国内机票按部门拆分
const processDomesticFlights = (workbook: any) => {
  console.log('=== 开始处理国内机票按部门拆分 ===');

  // 获取国内机票工作表
  const flightSheet = workbook.getWorksheet('国内机票');
  if (!flightSheet) {
    console.log('未找到"国内机票"工作表');
    return { departmentData: {}, columnMapping: {}, processedRows: 0 };
  }

  console.log('找到"国内机票"工作表');

  // 读取国内机票数据（包含空单元格）
  const flightData: any[][] = [];
  flightSheet.eachRow((row: any, rowNumber: number) => {
    const rowData: any[] = [];
    row.eachCell({ includeEmpty: true }, (cell: any, colNumber: number) => {
      rowData.push(cell.value);
    });
    flightData.push(rowData);
  });

  console.log(`国内机票数据行数: ${flightData.length}`);

  if (flightData.length < 2) {
    console.log('国内机票数据行数不足，跳过处理');
    return { departmentData: {}, columnMapping: {}, processedRows: 0 };
  }

  // 获取表头
  const headers = flightData[0];
  console.log('国内机票表头:', headers);

  // 查找"费用归属"列的索引
  const costBelongIndex = headers.findIndex((h: any) =>
    h && h.toString().includes('费用归属')
  );

  if (costBelongIndex === -1) {
    console.error('未找到国内机票"费用归属"列');
    ElMessage.warning('国内机票工作表格式不正确，缺少"费用归属"列');
    return { departmentData: {}, columnMapping: {}, processedRows: 0 };
  }

  console.log(`国内机票费用归属列索引: ${costBelongIndex}`);

  // 查找所有需要映射的列（新表头 -> 原表头映射）
  const columnMapping: { [key: string]: number } = {
    '票号': headers.findIndex((h: any) => h && h.toString().includes('票号')),
    '机票状态': headers.findIndex((h: any) => h && h.toString().includes('订单状态')),
    '预订人': headers.findIndex((h: any) => h && h.toString().includes('预订人')),
    '旅客姓名': headers.findIndex((h: any) => h && h.toString().includes('乘机人')),
    '旅客直属部门': costBelongIndex, // 使用费用归属列
    '行程': headers.findIndex((h: any) => h && h.toString().includes('行程')),
    '航班号': headers.findIndex((h: any) => h && h.toString().includes('航班号')),
    '出发日期': headers.findIndex((h: any) => h && h.toString().includes('出发日期')),
    '出发时间': headers.findIndex((h: any) => h && h.toString().includes('出发时间')),
    '票销售价': headers.findIndex((h: any) => h && h.toString().includes('票面价')),
    '机建费(国内)': headers.findIndex((h: any) => h && h.toString().includes('机建')),
    '燃油费(国内)': headers.findIndex((h: any) => h && h.toString().includes('燃油')),
    '改签费': headers.findIndex((h: any) => h && h.toString().includes('改签费')),
    '退票费': headers.findIndex((h: any) => h && h.toString().includes('退票费')),
    '服务费': headers.findIndex((h: any) => h && h.toString().includes('系统使用费')),
    '应还款总金额': headers.findIndex((h: any) => h && h.toString().includes('总金额'))
  };

  console.log('国内机票列映射:', columnMapping);

  // 按部门分组数据
  const departmentData: { [key: string]: any[][] } = {};
  let processedRows = 0;

  // 从第二行开始处理数据（跳过表头）
  for (let i = 1; i < flightData.length; i++) {
    const row = flightData[i];
    const costBelong = row[costBelongIndex]?.toString().trim();

    if (!costBelong) {
      console.log(`国内机票第 ${i + 1} 行费用归属为空，跳过`);
      continue;
    }

    // 提取部门名称（去掉"商务-机票-"前缀）
    let departmentName = costBelong;
    if (costBelong.startsWith('商务-机票-')) {
      departmentName = costBelong.replace('商务-机票-', '');
    } else if (costBelong.includes('-机票-')) {
      // 处理可能的其他格式，如"其他-机票-部门名"
      const parts = costBelong.split('-机票-');
      if (parts.length === 2) {
        departmentName = parts[1]; // 取第二部分作为部门名
      }
    }

    console.log(`国内机票第 ${i + 1} 行: 费用归属="${costBelong}" -> 部门="${departmentName}"`);

    // 如果部门不存在，创建新的数组
    if (!departmentData[departmentName]) {
      departmentData[departmentName] = [];
    }

    // 将数据行添加到对应部门
    departmentData[departmentName].push(row);
    processedRows++;
  }

  console.log('国内机票部门分组结果:', departmentData);
  console.log(`处理了国内机票 ${processedRows} 行数据，分成了 ${Object.keys(departmentData).length} 个部门`);

  return { departmentData, columnMapping, processedRows };
};

// 处理国内火车票按部门拆分
const processDomesticTrains = (workbook: any) => {
  console.log('=== 开始处理国内火车票按部门拆分 ===');

  // 尝试多种可能的火车票工作表名称
  const possibleSheetNames = ['国内火车票'];
  let trainSheet = null;
  let foundSheetName = '';

  for (const sheetName of possibleSheetNames) {
    trainSheet = workbook.getWorksheet(sheetName);
    if (trainSheet) {
      foundSheetName = sheetName;
      break;
    }
  }

  if (!trainSheet) {
    console.log('未找到火车票工作表，尝试的名称:', possibleSheetNames);
    console.log('当前所有工作表:', workbook.worksheets.map(ws => ws.name));
    ElMessage.warning('未找到火车票工作表，支持的工作表名称：国内火车、火车票、火车、国内火车票');
    return { departmentData: {}, columnMapping: {}, processedRows: 0 };
  }

  console.log(`找到火车票工作表: "${foundSheetName}"`);

  // 读取国内火车票数据（包含空单元格）
  const trainData: any[][] = [];
  trainSheet.eachRow((row: any, rowNumber: number) => {
    const rowData: any[] = [];
    row.eachCell({ includeEmpty: true }, (cell: any, colNumber: number) => {
      rowData.push(cell.value);
    });
    trainData.push(rowData);
  });

  console.log(`国内火车票数据行数: ${trainData.length}`);

  if (trainData.length < 2) {
    console.log('国内火车票数据行数不足，跳过处理');
    return { departmentData: {}, columnMapping: {}, processedRows: 0 };
  }

  // 获取表头
  const headers = trainData[0];
  console.log('国内火车票表头:', headers);

  // 查找"费用归属"列的索引
  const costBelongIndex = headers.findIndex((h: any) =>
    h && h.toString().includes('费用归属')
  );

  if (costBelongIndex === -1) {
    console.error('未找到国内火车票"费用归属"列');
    ElMessage.warning('国内火车票工作表格式不正确，缺少"费用归属"列');
    return { departmentData: {}, columnMapping: {}, processedRows: 0 };
  }

  console.log(`国内火车票费用归属列索引: ${costBelongIndex}`);

  // 查找所有需要映射的列（新表头 -> 原表头映射）
  const columnMapping: { [key: string]: number } = {
    '记账日期': headers.findIndex((h: any) => h && h.toString().includes('记账日期')),
    '订单状态': headers.findIndex((h: any) => h && h.toString().includes('订单状态')),
    '预订人': headers.findIndex((h: any) => h && h.toString().includes('预订人')),
    '乘车人': headers.findIndex((h: any) => h && h.toString().includes('乘车人')),
    '费用归属': costBelongIndex,
    '出发站': headers.findIndex((h: any) => h && h.toString().includes('出发站')),
    '到达站': headers.findIndex((h: any) => h && h.toString().includes('到达站')),
    '车次': headers.findIndex((h: any) => h && h.toString().includes('车次')),
    '出发日期': headers.findIndex((h: any) => h && h.toString().includes('出发日期')),
    '出发时间': headers.findIndex((h: any) => h && h.toString().includes('出发时间')),
    '坐席类型': headers.findIndex((h: any) => h && h.toString().includes('坐席类型')),
    '座位号': headers.findIndex((h: any) => h && h.toString().includes('座位号')),
    '车票费': headers.findIndex((h: any) => h && h.toString().includes('车票费')),
    '改签费': headers.findIndex((h: any) => h && h.toString().includes('改签费')),
    '退票费': headers.findIndex((h: any) => h && h.toString().includes('退票费')),
    '系统使用费': headers.findIndex((h: any) => h && h.toString().includes('系统使用费')),
    '总金额': headers.findIndex((h: any) => h && h.toString().includes('总金额'))
  };

  console.log('国内火车票列映射:', columnMapping);

  // 按部门分组数据
  const departmentData: { [key: string]: any[][] } = {};
  let processedRows = 0;

  // 从第二行开始处理数据（跳过表头）
  for (let i = 1; i < trainData.length; i++) {
    const row = trainData[i];
    const costBelong = row[costBelongIndex]?.toString().trim();

    if (!costBelong) {
      console.log(`国内火车票第 ${i + 1} 行费用归属为空，跳过`);
      continue;
    }

    // 提取部门名称，兼容多种格式
    let departmentName = costBelong;
    console.log(`原始费用归属: "${costBelong}"`);

    // 尝试多种格式的部门名称提取
    if (costBelong.startsWith('商务-火车-')) {
      departmentName = costBelong.replace('商务-火车-', '');
    } else if (costBelong.includes('-火车-')) {
      // 处理其他格式，如"其他-火车-部门名"
      const parts = costBelong.split('-火车-');
      if (parts.length === 2) {
        departmentName = parts[1];
      }
    } else if (costBelong.startsWith('商务-机票-')) {
      // 如果是机票格式，提取部门名
      departmentName = costBelong.replace('商务-机票-', '');
    } else if (costBelong.includes('-机票-')) {
      const parts = costBelong.split('-机票-');
      if (parts.length === 2) {
        departmentName = parts[1];
      }
    } else {
      // 如果都没有匹配到，直接使用原值
      departmentName = costBelong;
    }

    console.log(`提取的部门名称: "${departmentName}"`);

    console.log(`国内火车票第 ${i + 1} 行: 费用归属="${costBelong}" -> 部门="${departmentName}"`);

    // 如果部门不存在，创建新的数组
    if (!departmentData[departmentName]) {
      departmentData[departmentName] = [];
    }

    // 将数据行添加到对应部门
    departmentData[departmentName].push(row);
    processedRows++;
  }

  console.log('国内火车票部门分组结果:', departmentData);
  console.log(`处理了国内火车票 ${processedRows} 行数据，分成了 ${Object.keys(departmentData).length} 个部门`);

  if (processedRows === 0) {
    console.warn('警告：没有处理到任何国内火车票数据');
    ElMessage.warning('未找到有效的国内火车票数据，请检查工作表名称和数据格式');
  }

  return { departmentData, columnMapping, processedRows };
};

// 处理国内酒店按部门拆分
const processDomesticHotels = (workbook: any) => {
  console.log('=== 开始处理国内酒店按部门拆分 ===');

  // 尝试多种可能的酒店工作表名称
  const possibleSheetNames = ['国内酒店', '酒店', '酒店票', '国内酒店票'];
  let hotelSheet = null;
  let foundSheetName = '';

  for (const sheetName of possibleSheetNames) {
    hotelSheet = workbook.getWorksheet(sheetName);
    if (hotelSheet) {
      foundSheetName = sheetName;
      break;
    }
  }

  if (!hotelSheet) {
    console.log('未找到酒店工作表，尝试的名称:', possibleSheetNames);
    console.log('当前所有工作表:', workbook.worksheets.map((ws: any) => ws.name));
    ElMessage.warning('未找到酒店工作表，支持的工作表名称：国内酒店、酒店、酒店票、国内酒店票');
    return { departmentData: {}, columnMapping: {}, processedRows: 0 };
  }

  console.log(`找到酒店工作表: "${foundSheetName}"`);

  // 读取国内酒店数据（包含空单元格）
  const hotelData: any[][] = [];
  hotelSheet.eachRow((row: any, rowNumber: number) => {
    const rowData: any[] = [];
    row.eachCell({ includeEmpty: true }, (cell: any, colNumber: number) => {
      rowData.push(cell.value);
    });
    hotelData.push(rowData);
  });

  console.log(`国内酒店数据行数: ${hotelData.length}`);

  if (hotelData.length < 2) {
    console.log('国内酒店数据行数不足，跳过处理');
    return { departmentData: {}, columnMapping: {}, processedRows: 0 };
  }

  // 获取表头
  const headers = hotelData[0];
  console.log('国内酒店表头:', headers);

  // 查找"费用归属"列的索引
  const costBelongIndex = headers.findIndex((h: any) =>
    h && h.toString().includes('费用归属')
  );

  if (costBelongIndex === -1) {
    console.error('未找到国内酒店"费用归属"列');
    ElMessage.warning('国内酒店工作表格式不正确，缺少"费用归属"列');
    return { departmentData: {}, columnMapping: {}, processedRows: 0 };
  }

  console.log(`国内酒店费用归属列索引: ${costBelongIndex}`);

  // 查找所有需要映射的列（新表头 -> 原表头映射）
  const columnMapping: { [key: string]: number } = {
    '记账日期': headers.findIndex((h: any) => h && h.toString().includes('记账日期')),
    '订单状态': headers.findIndex((h: any) => h && h.toString().includes('订单状态')),
    '预订人': headers.findIndex((h: any) => h && (h.toString().includes('预订人') || h.toString().includes('预定人'))),
    '入住人': headers.findIndex((h: any) => h && h.toString().includes('入住人')),
    '费用归属': costBelongIndex,
    '入住日期': headers.findIndex((h: any) => h && h.toString().includes('入住日期')),
    '离店日期': headers.findIndex((h: any) => h && h.toString().includes('离店日期')),
    '酒店城市': headers.findIndex((h: any) => h && h.toString().includes('酒店城市')),
    '酒店名称': headers.findIndex((h: any) => h && h.toString().includes('酒店名称')),
    '夜数': headers.findIndex((h: any) => h && h.toString().includes('夜数')),
    '订房费用': headers.findIndex((h: any) => h && h.toString().includes('订房费用')),
    '系统使用费': headers.findIndex((h: any) => h && h.toString().includes('系统使用费')),
    '酒店托管费': headers.findIndex((h: any) => h && h.toString().includes('酒店托管费')),
    '代购费': headers.findIndex((h: any) => h && h.toString().includes('代购费')),
    '总金额': headers.findIndex((h: any) => h && h.toString().includes('总金额'))
  };

  console.log('国内酒店列映射:', columnMapping);

  // 按部门分组数据
  const departmentData: { [key: string]: any[][] } = {};
  let processedRows = 0;

  // 从第二行开始处理数据（跳过表头）
  for (let i = 1; i < hotelData.length; i++) {
    const row = hotelData[i];
    const costBelong = row[costBelongIndex]?.toString().trim();

    if (!costBelong) {
      console.log(`国内酒店第 ${i + 1} 行费用归属为空，跳过`);
      continue;
    }

    // 提取部门名称，兼容多种格式
    let departmentName = costBelong;
    console.log(`原始费用归属: "${costBelong}"`);

    // 尝试多种格式的部门名称提取
    if (costBelong.startsWith('商务-酒店-')) {
      departmentName = costBelong.replace('商务-酒店-', '');
    } else if (costBelong.includes('-酒店-')) {
      // 处理其他格式，如"其他-酒店-部门名"
      const parts = costBelong.split('-酒店-');
      if (parts.length === 2) {
        departmentName = parts[1];
      }
    } else if (costBelong.startsWith('商务-机票-')) {
      // 如果是机票格式，提取部门名
      departmentName = costBelong.replace('商务-机票-', '');
    } else if (costBelong.includes('-机票-')) {
      const parts = costBelong.split('-机票-');
      if (parts.length === 2) {
        departmentName = parts[1];
      }
    } else {
      // 如果都没有匹配到，直接使用原值
      departmentName = costBelong;
    }

    console.log(`提取的部门名称: "${departmentName}"`);

    // 如果部门不存在，创建新的数组
    if (!departmentData[departmentName]) {
      departmentData[departmentName] = [];
    }

    // 将数据行添加到对应部门
    departmentData[departmentName].push(row);
    processedRows++;
  }

  console.log('国内酒店部门分组结果:', departmentData);
  console.log(`处理了国内酒店 ${processedRows} 行数据，分成了 ${Object.keys(departmentData).length} 个部门`);

  if (processedRows === 0) {
    console.warn('警告：没有处理到任何国内酒店数据');
    ElMessage.warning('未找到有效的国内酒店数据，请检查工作表名称和数据格式');
  }

  return { departmentData, columnMapping, processedRows };
};

// 生成火车票部门报告
const generateTrainDepartmentReport = async (departmentName: string, trainData: any[], columnMapping: any) => {
  const fullDepartmentName = `商务-火车-${departmentName}`;

  try {
    console.log(`=== 生成 ${fullDepartmentName} 火车票部门报告 ===`);

    // 创建新的工作簿
    const newWorkbook = new ExcelJS.Workbook();
    const worksheet = newWorkbook.addWorksheet(fullDepartmentName);

    // 计算上个月日期
    const now = new Date();
    const lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    const year = lastMonth.getFullYear();
    const month = lastMonth.getMonth() + 1;
    const monthStr = month.toString().padStart(2, '0');

    // 第一行：标题
    const titleRow = worksheet.addRow([`华安保险${year}年${month}月火车对账单`]);
    titleRow.font = { bold: true, size: 22, name: '微软雅黑' };
    titleRow.alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.mergeCells(1, 1, 1, 17);
    worksheet.getRow(1).height = 53;

    // 第二行：部门信息
    const deptRow = worksheet.addRow([`部门：${departmentName}`]);
    deptRow.font = { bold: true, size: 12 };
    deptRow.alignment = { vertical: 'middle' };
    worksheet.mergeCells(2, 1, 2, 17);
    worksheet.getRow(2).height = 30;

    // 第三行：火车票表头
    const trainHeaders = [
      '预订/退款日期', '订单状态', '预订人', '旅客姓名', '旅客直属部门', '行程', '车次', '出发时间',
      '坐席', '座位编号', '车票单价', '改签费', '退票费', '销售总价', '企业支付', '服务费', '应还款总金额'
    ];
    const headerRow = worksheet.addRow(trainHeaders);
    headerRow.font = { bold: true, color: { argb: 'FFFFFF' } }; // 白色字体
    headerRow.alignment = { horizontal: 'center', vertical: 'middle' }; // 居中对齐
    worksheet.getRow(3).height = 30;

    // 设置表头行的填充颜色和边框
    for (let i = 1; i <= 17; i++) {
      const cell = headerRow.getCell(i);
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFF8E22' } // #FF8E22 橙色
      };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    }

    // 处理火车票数据
    const processedTrainData = trainData.map((originalRow) => {
      // 合并出发站和到达站形成行程
      const departureStation = originalRow[columnMapping['出发站']] || '';
      const arrivalStation = originalRow[columnMapping['到达站']] || '';
      const itinerary = `${departureStation || ''}${arrivalStation ? '-' + arrivalStation : ''}`;

      // 合并出发日期和时间
      const departureDate = originalRow[columnMapping['出发日期']] || '';
      const departureTime = originalRow[columnMapping['出发时间']] || '';
      const fullDepartureTime = processDepartureTime(departureDate, departureTime);

      return {
        processedRow: [
          originalRow[columnMapping['记账日期']] || '', // 预订/退款日期 -> 记账日期
          originalRow[columnMapping['订单状态']] || '', // 订单状态
          originalRow[columnMapping['预订人']] || '', // 预订人
          originalRow[columnMapping['乘车人']] || '', // 旅客姓名 -> 乘车人
          originalRow[columnMapping['费用归属']] || '', // 旅客直属部门 -> 费用归属
          itinerary, // 行程 -> 出发站+到达站
          originalRow[columnMapping['车次']] || '', // 车次
          fullDepartureTime, // 出发时间 -> 出发日期+出发时间
          originalRow[columnMapping['坐席类型']] || '', // 坐席 -> 坐席类型
          originalRow[columnMapping['座位号']] || '', // 座位编号 -> 座位号
          originalRow[columnMapping['车票费']] || 0, // 车票单价 -> 车票费
          originalRow[columnMapping['改签费']] || 0, // 改签费
          originalRow[columnMapping['退票费']] || 0, // 退票费
          undefined, // 销售总价 - 留空
          undefined, // 企业支付 - 留空
          originalRow[columnMapping['系统使用费']] || 0, // 服务费 -> 系统使用费
          originalRow[columnMapping['总金额']] || 0, // 应还款总金额 -> 总金额
        ],
        passengerName: originalRow[columnMapping['乘车人']] || '',
        departureTime: fullDepartureTime,
        originalRow: originalRow
      };
    });

    // 按旅客姓名分组和排序
    const passengerNames = Array.from(new Set(processedTrainData.map(item => item.passengerName || '未知乘客')));
    passengerNames.sort((a, b) => a.localeCompare(b, 'zh-CN'));

    const groupedByPassenger: { [key: string]: typeof processedTrainData } = {};
    passengerNames.forEach(passengerName => {
      const passengerData = processedTrainData.filter(item =>
        (item.passengerName || '未知乘客') === passengerName
      );
      groupedByPassenger[passengerName] = passengerData;
    });

    // 对每个分组按出发时间排序
    Object.keys(groupedByPassenger).forEach(passengerName => {
      groupedByPassenger[passengerName].sort((a, b) => {
        const timeA = new Date(a.departureTime || '').getTime();
        const timeB = new Date(b.departureTime || '').getTime();
        return timeA - timeB;
      });
    });

    // 添加数据行和小计
    let currentRowNumber = 4; // 数据从第4行开始
    Object.keys(groupedByPassenger).forEach(passengerName => {
      const passengerData = groupedByPassenger[passengerName];

      // 添加该乘客的所有数据行
      passengerData.forEach(item => {
        const newRow = worksheet.addRow(item.processedRow);
        worksheet.getRow(newRow.number).height = 30;

        // 设置销售总价和企业支付公式
        // 销售总价（第14列）= 车票单价 + 改签费 + 退票费
        const salesTotalCell = newRow.getCell(14);
        salesTotalCell.value = {
          formula: `=K${currentRowNumber}+L${currentRowNumber}+M${currentRowNumber}`,
          result: 0
        };

        // 企业支付（第15列）= 销售总价
        const enterprisePaymentCell = newRow.getCell(15);
        enterprisePaymentCell.value = {
          formula: `=N${currentRowNumber}`,
          result: 0
        };
        currentRowNumber++;
      });

      // 添加小计行
      if (passengerData.length > 0) {
        const startRow = currentRowNumber - passengerData.length;
        const endRow = currentRowNumber - 1;

        const subtotalRow = worksheet.addRow([
          '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', // 前16列全部留空，将合并显示"小计"
          { // 应还款总金额（第17列）设置求和公式
            formula: `=SUM(Q${startRow}:Q${endRow})`,
            result: 0
          }
        ]);
        worksheet.getRow(subtotalRow.number).height = 30;

        // 合并小计行的第1列到第16列，用于显示"小计"文字
        worksheet.mergeCells(subtotalRow.number, 1, subtotalRow.number, 16);

        // 设置小计行样式
        subtotalRow.font = { bold: true };
        subtotalRow.getCell(1).value = '小计'; // 在合并后的第一个单元格中设置"小计"文字
        subtotalRow.getCell(1).alignment = { horizontal: 'right', vertical: 'middle' }; // 小计文字右对齐垂直居中
        subtotalRow.getCell(17).numFmt = '#,##0.00'; // 设置应还款总金额的数字格式

        currentRowNumber++;
      }
    });

    // 添加总计行
    const totalStartRow = 4; // 数据开始行
    const totalEndRow = currentRowNumber - 1; // 最后一行数据（包括小计行）

    const totalRow = worksheet.addRow([
      '', '', '', '', // 预订/退款日期、订单状态、预订人、旅客姓名留空
      '', // 旅客直属部门留空
      '', '', '', '', // 行程、车次、出发时间、坐席留空
      '', '', '', '', '', '', // 座位编号、车票单价、改签费、退票费、销售总价、企业支付留空
      '', // 服务费留空
      { // 应还款总金额（第17列）对小计行求和
        formula: `=SUMIF(A${totalStartRow}:A${totalEndRow},"小计",Q${totalStartRow}:Q${totalEndRow})`,
        result: 0
      }
    ]);
    worksheet.getRow(totalRow.number).height = 30;

    // 设置总计行样式
    totalRow.font = { bold: true };
    totalRow.getCell(5).alignment = { horizontal: 'right' }; // 总计文字右对齐
    totalRow.getCell(17).numFmt = '#,##0.00'; // 应还款总金额数字格式

    // 合并总计行的第1列到第16列
    worksheet.mergeCells(totalRow.number, 1, totalRow.number, 16);

    // 添加签名行（最后一行的下一行）
    const signatureRow = worksheet.addRow([
      '', '经办人：', '', '', '审核人：', '', '', '日期：', '', '', '部门负责人审批：', '', '', ''
    ]);

    // 设置签名行样式
    signatureRow.alignment = { vertical: 'middle' }; // 签名行垂直居中
    worksheet.getRow(signatureRow.number).height = 66; // 签名行高度设置为66磅

    // 合并单元格用于签名信息
    worksheet.mergeCells(signatureRow.number, 2, signatureRow.number, 4); // 合并B-D列（第2-4列）经办人
    worksheet.mergeCells(signatureRow.number, 5, signatureRow.number, 7); // 合并E-G列（第5-7列）审核人
    worksheet.mergeCells(signatureRow.number, 8, signatureRow.number, 10); // 合并H-J列（第8-10列）日期
    worksheet.mergeCells(signatureRow.number, 11, signatureRow.number, 13); // 合并K-M列（第11-13列）部门负责人审批

    // 设置第四行到总计行的样式（跳过第3行表头）
    for (let i = 4; i <= totalRow.number; i++) {
      const row = worksheet.getRow(i);

      // 检查是否是小计行，如果是则设置行高为30磅
      const cell = row.getCell(1); // 第1列是小计标识列
      if (cell.value && cell.value.toString() === '小计') {
        row.height = 30;
      }

      // 设置每个单元格的样式
      for (let j = 1; j <= 17; j++) {
        const cell = row.getCell(j);

        // 如果是小计行的第一个单元格，保持原有的右对齐设置
        const isSubtotalRow = cell.value && cell.value.toString() === '小计';

        if (!isSubtotalRow) {
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        }

        cell.font = { size: 10 };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      }
    }

    // 收集所有数据行用于智能列宽计算
    const allTrainData = trainData.map(row => [
      row['预订/退款日期'] || '',
      row['订单状态'] || '',
      row['预订人'] || '',
      row['旅客姓名'] || '',
      row['旅客直属部门'] || '',
      row['行程'] || '',
      row['车次'] || '',
      row['出发时间'] || '',
      row['坐席'] || '',
      row['座位编号'] || '',
      row['车票单价'] || 0,
      row['改签费'] || 0,
      row['退票费'] || 0,
      row['销售总价'] || 0,
      row['企业支付'] || 0,
      row['服务费'] || 0,
      row['应还款总金额'] || 0
    ]);

    setSmartColumnWidths(worksheet, trainHeaders, allTrainData, 'train');

    // 设置金额列的数字格式
    worksheet.getColumn(11).numFmt = '#,##0.00'; // 车票单价
    worksheet.getColumn(12).numFmt = '#,##0.00'; // 改签费
    worksheet.getColumn(13).numFmt = '#,##0.00'; // 退票费
    worksheet.getColumn(16).numFmt = '#,##0.00'; // 服务费
    worksheet.getColumn(17).numFmt = '#,##0.00'; // 应还款总金额

    // 生成文件
    const fileName = `${fullDepartmentName}.xlsx`;
    generatedFiles.value.push({
      fileName,
      departmentName: fullDepartmentName,
      rowCount: trainData.length,
      workbook: newWorkbook
    });

    console.log(`已准备生成火车票文件: ${fileName}, 包含 ${trainData.length} 条数据`);

  } catch (error) {
    console.error(`生成 ${fullDepartmentName} 火车票部门报告失败:`, error);
  }
};

// 生成机票部门报告
const generateFlightDepartmentReport = async (departmentName: string, flightData: Array<{data: any[], type: string}>, columnMappings: { domestic: any, international: any }) => {
  // 处理贵宾部门名称的特殊逻辑
  const isVip = departmentName === '贵宾';
  const displayDepartmentName = isVip ? '无' : departmentName;

  // 对于贵宾部门，需要先获取旅客姓名来动态设置工作表名和文件名
  let passengerName = '';
  if (isVip && flightData.length > 0) {
    const columnMapping = flightData[0].type === 'international-flight' ? columnMappings.international : columnMappings.domestic;
    passengerName = flightData[0].data[columnMapping['旅客姓名']] || '';
  }

  const worksheetName = isVip ? `商务-机票-${passengerName}` : `商务-机票-${departmentName}`;
  const fullDepartmentName = `商务-机票-${departmentName}`;

  try {
    console.log(`=== 生成 ${fullDepartmentName} 机票部门报告 ===`);

    // 创建新的工作簿
    const newWorkbook = new ExcelJS.Workbook();
    const worksheet = newWorkbook.addWorksheet(worksheetName);

    // 计算上个月日期
    const now = new Date();
    const lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    const year = lastMonth.getFullYear();
    const month = lastMonth.getMonth() + 1;
    const monthStr = month.toString().padStart(2, '0');

    // 第一行：标题
    const titleRow = worksheet.addRow([`华安保险${year}年${month}月机票对账单`]);
    titleRow.font = { bold: true, size: 22, name: '微软雅黑' };
    titleRow.alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.mergeCells(1, 1, 1, 20);
    worksheet.getRow(1).height = 53;

    // 第二行：部门信息
    const deptRow = worksheet.addRow([`部门：${displayDepartmentName}`]);
    deptRow.font = { bold: true, size: 12 };
    deptRow.alignment = { vertical: 'middle' };
    worksheet.mergeCells(2, 1, 2, 20);
    worksheet.getRow(2).height = 30;

    // 第三行：表头
    const headers = [
      '动支号', '票号', '机票状态', '预订人', '旅客姓名', '旅客直属部门', '行程', '航班号',
      '起飞时间', '票销售价', '机建费(国内)', '燃油费(国内)', '改签费', '升舱费', '退票费', '销售总价',
      '企业支付', '服务费', '应还款总金额', '签字确认'
    ];
    const headerRow = worksheet.addRow(headers);
    headerRow.font = { bold: true, color: { argb: 'FFFFFF' } }; // 白色字体
    headerRow.alignment = { horizontal: 'center', vertical: 'middle' }; // 居中对齐
    worksheet.getRow(3).height = 30;

    // 设置表头行的填充颜色和边框
    for (let i = 1; i <= 20; i++) {
      const cell = headerRow.getCell(i);
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFF8E22' } // #FF8E22 橙色
      };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    }

    // 处理机票数据
    const processedAllData = flightData.map(({ data: originalRow, type }) => {
      const isInternational = type === 'international-flight';
      const columnMapping = isInternational ? columnMappings.international : columnMappings.domestic;

      let processedItinerary = '';
      let processedDepartureTime = '';
      let processedPassengerName = '';

      if (isInternational) {
        // 国际机票处理
        const originalItinerary = originalRow[columnMapping['航程']] || '';
        processedItinerary = originalItinerary; // 国际机票航程直接使用

        const originalDepartureTime = originalRow[columnMapping['出发时间']] || '';
        processedDepartureTime = originalDepartureTime; // 国际机票直接使用出发时间

        // 转换英文姓名为中文姓名
        const originalPassengerName = originalRow[columnMapping['旅客姓名']] || '';
        processedPassengerName = convertEnglishNameToChinese(originalPassengerName);
      } else {
        // 国内机票处理
        const originalItinerary = originalRow[columnMapping['行程']] || '';
        processedItinerary = processItinerary(originalItinerary);

        const departureDate = originalRow[columnMapping['出发日期']] || '';
        const departureTime = originalRow[columnMapping['出发时间']] || '';
        processedDepartureTime = processDepartureTime(departureDate, departureTime);

        // 国内机票直接使用原姓名
        processedPassengerName = originalRow[columnMapping['旅客姓名']] || '';
      }

      // 动支号列都是空的，直接设置为空字符串
      const dongzhikaoValue = '';

      return {
        processedRow: [
          dongzhikaoValue, // 动支号
          originalRow[columnMapping['票号']] || '', // 票号
          originalRow[columnMapping['机票状态']] || '', // 机票状态 -> 订单状态
          originalRow[columnMapping['预订人']] || '', // 预订人
          processedPassengerName, // 转换后的旅客姓名
          originalRow[columnMapping['旅客直属部门']] || '', // 旅客直属部门 -> 费用归属
          processedItinerary, // 处理后的行程/航程
          originalRow[columnMapping['航班号']] || '', // 航班号
          processedDepartureTime, // 处理后的起飞时间
          originalRow[columnMapping['票销售价']] || 0, // 票销售价 -> 票面价
          isInternational
            ? (originalRow[columnMapping['税费']] || 0) // 国际机票：税费 -> 机建费(国内)
            : (originalRow[columnMapping['机建费(国内)']] || 0), // 国内机票：机建费(国内)
          originalRow[columnMapping['燃油费(国内)']] || 0, // 燃油费(国内) -> 燃油
          originalRow[columnMapping['改签费']] || 0, // 改签费
          '', // 升舱费 - 留空
          originalRow[columnMapping['退票费']] || 0, // 退票费
          undefined, // 销售总价 - 稍后设置公式
          undefined, // 企业支付 - 稍后设置公式（引用P列）
          originalRow[columnMapping['服务费']] || 0, // 服务费 -> 系统使用费
          originalRow[columnMapping['应还款总金额']] || 0, // 应还款总金额 -> 总金额
          ''  // 签字确认 - 留空
        ],
        passengerName: processedPassengerName,
        departureTime: processedDepartureTime,
        originalRow: originalRow,
        isInternational: isInternational
      };
    });

    // 先获取所有旅客姓名并按中文拼音排序
    const passengerNames = Array.from(new Set(processedAllData.map(item => item.passengerName || '未知乘客')));

    // 中文拼音排序比较函数
    const chineseCompare = (a: string, b: string) => {
      // 使用 localeCompare 进行中文拼音排序
      return a.localeCompare(b, 'zh-CN');
    };

    // 对旅客姓名进行排序
    passengerNames.sort(chineseCompare);

    // 按排序后的旅客姓名分组
    const groupedByPassenger: { [key: string]: typeof processedAllData } = {};
    passengerNames.forEach(passengerName => {
      const passengerData = processedAllData.filter(item =>
        (item.passengerName || '未知乘客') === passengerName
      );
      groupedByPassenger[passengerName] = passengerData;
    });

    // 对每个分组分别按机票类型和起飞时间排序
    Object.keys(groupedByPassenger).forEach(passengerName => {
      groupedByPassenger[passengerName].sort((a, b) => {
        // 首先按机票类型排序：国内机票在前，国际机票在后
        if (a.isInternational !== b.isInternational) {
          return a.isInternational ? 1 : -1; // 国内(false)在前，国际(true)在后
        }

        // 相同机票类型下按起飞时间排序（升序）
        const timeA = new Date(a.departureTime || '').getTime();
        const timeB = new Date(b.departureTime || '').getTime();
        return timeA - timeB;
      });
    });

    // 添加排序后的数据行和分组小计
    let currentRowNumber = 4; // 数据从第4行开始

    Object.keys(groupedByPassenger).forEach(passengerName => {
      const passengerData = groupedByPassenger[passengerName];

      // 添加该乘客的所有数据行
      passengerData.forEach(item => {
        const newRow = worksheet.addRow(item.processedRow);
        worksheet.getRow(newRow.number).height = 30;

        // 设置销售总价公式（第16列）
        const salesTotalCell = newRow.getCell(16);
        salesTotalCell.value = {
          formula: `=J${currentRowNumber}+K${currentRowNumber}+L${currentRowNumber}+M${currentRowNumber}+O${currentRowNumber}`,
          result: 0
        };

        // 设置企业支付公式（第17列），直接引用销售总价（P列）
        const enterprisePaymentCell = newRow.getCell(17);
        enterprisePaymentCell.value = {
          formula: `=P${currentRowNumber}`,
          result: 0
        };

        currentRowNumber++;
      });

      // 添加小计行
      if (passengerData.length > 0) {
        const startRow = currentRowNumber - passengerData.length;
        const endRow = currentRowNumber - 1;

        const subtotalRow = worksheet.addRow([
          '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', // 前18列全部留空，将合并显示"小计"
          { // 应还款总金额（第19列）设置求和公式
            formula: `=SUM(S${startRow}:S${endRow})`,
            result: 0
          },
          '' // 签字确认留空（第20列）
        ]);
        worksheet.getRow(subtotalRow.number).height = 30;

        // 合并小计行的第1列到第18列，用于显示"小计"文字
        worksheet.mergeCells(subtotalRow.number, 1, subtotalRow.number, 18);

        // 设置小计行样式
        subtotalRow.font = { bold: true };
        subtotalRow.getCell(1).value = '小计'; // 在合并后的第一个单元格中设置"小计"文字
        subtotalRow.getCell(1).alignment = { horizontal: 'right', vertical: 'middle' }; // 小计文字右对齐垂直居中
        subtotalRow.getCell(19).numFmt = '#,##0.00'; // 设置应还款总金额的数字格式

        currentRowNumber++;
      }
    });

    // 在所有数据行添加完成后，需要记录每个分组的数据行范围，以便后续合并动支号列
    console.log('准备合并动支号列，当前行数:', currentRowNumber);

    // 记录每个分组的起始行和结束行
    const passengerGroups: Array<{ startRow: number, endRow: number, passengerName: string }> = [];
    let tempRowNumber = 4; // 重置临时行号计数器

    Object.keys(groupedByPassenger).forEach(passengerName => {
      const passengerData = groupedByPassenger[passengerName];
      const startRow = tempRowNumber;
      const endRow = tempRowNumber + passengerData.length - 1;

      if (passengerData.length > 1) { // 只有多行数据时才合并动支号
        passengerGroups.push({ startRow, endRow, passengerName });
        console.log(`乘客 ${passengerName}: 行 ${startRow} 到 ${endRow}`);
      }

      tempRowNumber += passengerData.length + 1; // +1 是因为每个乘客后面都有一个小计行
    });

    // 合并每个乘客分组内第1列（动支号列）的单元格
    passengerGroups.forEach(group => {
      worksheet.mergeCells(group.startRow, 1, group.endRow, 1);
      console.log(`合并乘客 ${group.passengerName} 的动支号列：第${group.startRow}行到第${group.endRow}行`);

      // 设置合并后的单元格样式
      const mergedCell = worksheet.getCell(group.startRow, 1);
      mergedCell.alignment = { horizontal: 'center', vertical: 'middle' };
    });

    // 添加总计行
    const totalStartRow = 4; // 数据开始行
    const totalEndRow = currentRowNumber - 1; // 最后一行数据（包括小计行）

    const totalRow = worksheet.addRow([
      '', '', '', '', // 动支号、票号、机票状态、预订人留空
      '', // 旅客姓名列留空
      '', '', '', '', // 旅客直属部门、行程、航班号、起飞时间留空
      { // 票销售价（第10列）求和
        formula: `=SUM(J${totalStartRow}:J${totalEndRow})`,
        result: 0
      },
      { // 机建费(国内)（第11列）求和
        formula: `=SUM(K${totalStartRow}:K${totalEndRow})`,
        result: 0
      },
      { // 燃油费(国内)（第12列）求和
        formula: `=SUM(L${totalStartRow}:L${totalEndRow})`,
        result: 0
      },
      { // 改签费（第13列）求和
        formula: `=SUM(M${totalStartRow}:M${totalEndRow})`,
        result: 0
      },
      { // 升舱费（第14列）求和
        formula: `=SUM(N${totalStartRow}:N${totalEndRow})`,
        result: 0
      },
      { // 退票费（第15列）求和
        formula: `=SUM(O${totalStartRow}:O${totalEndRow})`,
        result: 0
      },
      { // 销售总价（第16列）求和
        formula: `=SUM(P${totalStartRow}:P${totalEndRow})`,
        result: 0
      },
      { // 企业支付（第17列）求和
        formula: `=SUM(Q${totalStartRow}:Q${totalEndRow})`,
        result: 0
      },
      { // 服务费（第18列）求和
        formula: `=SUM(R${totalStartRow}:R${totalEndRow})`,
        result: 0
      },
      { // 应还款总金额（第19列）对小计行求和
        formula: `=SUMIF(A${totalStartRow}:A${totalEndRow},"小计",S${totalStartRow}:S${totalEndRow})`,
        result: 0
      },
      '' // 签字确认留空
    ]);
    worksheet.getRow(totalRow.number).height = 30;

    // 设置总计行样式
    totalRow.font = { bold: true };
    totalRow.getCell(5).alignment = { horizontal: 'right' }; // 总计文字右对齐

    // 合并总计行的第1列到第9列
    worksheet.mergeCells(totalRow.number, 1, totalRow.number, 9);

    // 设置总计行的数字格式
    totalRow.getCell(10).numFmt = '#,##0.00'; // 票销售价
    totalRow.getCell(11).numFmt = '#,##0.00'; // 机建费(国内)
    totalRow.getCell(12).numFmt = '#,##0.00'; // 燃油费(国内)
    totalRow.getCell(13).numFmt = '#,##0.00'; // 改签费
    totalRow.getCell(14).numFmt = '#,##0.00'; // 升舱费
    totalRow.getCell(15).numFmt = '#,##0.00'; // 退票费
    totalRow.getCell(16).numFmt = '#,##0.00'; // 销售总价
    totalRow.getCell(17).numFmt = '#,##0.00'; // 企业支付
    totalRow.getCell(18).numFmt = '#,##0.00'; // 服务费
    totalRow.getCell(19).numFmt = '#,##0.00'; // 应还款总金额

    // 添加签名行（最后一行的下一行）
    const signatureRow = worksheet.addRow([
      '', '经办人：', '', '', '审核人：', '', '', '日期：', '', '', '部门负责人审批：', '', '', ''
    ]);

    // 设置签名行样式
    signatureRow.alignment = { vertical: 'middle' }; // 签名行垂直居中
    worksheet.getRow(signatureRow.number).height = 66; // 签名行高度设置为66磅

    // 合并单元格用于签名信息
    worksheet.mergeCells(signatureRow.number, 2, signatureRow.number, 4); // 合并B-D列（第2-4列）经办人
    worksheet.mergeCells(signatureRow.number, 5, signatureRow.number, 7); // 合并E-G列（第5-7列）审核人
    worksheet.mergeCells(signatureRow.number, 8, signatureRow.number, 10); // 合并H-J列（第8-10列）日期
    worksheet.mergeCells(signatureRow.number, 11, signatureRow.number, 13); // 合并K-M列（第11-13列）部门负责人审批

    // 设置第四行到总计行的样式（跳过第3行表头）
    for (let i = 4; i <= totalRow.number; i++) {
      const row = worksheet.getRow(i);

      // 检查是否是小计行，如果是则设置行高为30磅
      const cell = row.getCell(1); // 第1列是小计标识列
      if (cell.value && cell.value.toString() === '小计') {
        row.height = 30;
      }

      // 设置每个单元格的样式
      for (let j = 1; j <= 20; j++) {
        const cell = row.getCell(j);

        // 如果是小计行的第一个单元格，保持原有的右对齐设置
        const isSubtotalRow = cell.value && cell.value.toString() === '小计';

        if (!isSubtotalRow) {
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        }

        cell.font = { size: 10 };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      }
    }

    // 收集所有数据行用于智能列宽计算
    const allFlightData = processedAllData.map(item => item.processedRow);
    setSmartColumnWidths(worksheet, headers, allFlightData, 'flight');

    // 设置金额列的数字格式（千分号和两位小数）
    worksheet.getColumn(10).numFmt = '#,##0.00'; // 票销售价
    worksheet.getColumn(11).numFmt = '#,##0.00'; // 机建费(国内)
    worksheet.getColumn(12).numFmt = '#,##0.00'; // 燃油费(国内)
    worksheet.getColumn(13).numFmt = '#,##0.00'; // 改签费
    worksheet.getColumn(15).numFmt = '#,##0.00'; // 退票费
    worksheet.getColumn(16).numFmt = '#,##0.00'; // 销售总价
    worksheet.getColumn(17).numFmt = '#,##0.00'; // 企业支付
    worksheet.getColumn(18).numFmt = '#,##0.00'; // 服务费
    worksheet.getColumn(19).numFmt = '#,##0.00'; // 应还款总金额

    // 计算总数据量
    const totalRows = flightData.reduce((sum, { data }) => sum + data.length, 0);

    // 生成文件
    const fileName = `${worksheetName}.xlsx`;
    generatedFiles.value.push({
      fileName,
      departmentName: worksheetName,
      rowCount: totalRows,
      workbook: newWorkbook
    });

    console.log(`已准备生成机票文件: ${fileName}, 包含 ${totalRows} 条数据`);

  } catch (error) {
    console.error(`生成 ${fullDepartmentName} 机票部门报告失败:`, error);
  }
};

// 生成酒店部门报告
const generateHotelDepartmentReport = async (departmentName: string, hotelData: any[], columnMapping: any, headers: any[] = []) => {
  const fullDepartmentName = `商务-酒店-${departmentName}`;

  try {
    console.log(`=== 生成 ${fullDepartmentName} 酒店部门报告 ===`);

    // 创建新的工作簿
    const newWorkbook = new ExcelJS.Workbook();
    const worksheet = newWorkbook.addWorksheet(fullDepartmentName);

    // 计算上个月日期
    const now = new Date();
    const lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    const year = lastMonth.getFullYear();
    const month = lastMonth.getMonth() + 1;
    const monthStr = month.toString().padStart(2, '0');

    // 第一行：标题
    const titleRow = worksheet.addRow([`华安保险${year}年${month}月酒店对账单`]);
    titleRow.font = { bold: true, size: 22, name: '微软雅黑' };
    titleRow.alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.mergeCells(1, 1, 1, 16);
    worksheet.getRow(1).height = 53;

    // 第二行：部门信息
    const deptRow = worksheet.addRow([`部门：${departmentName}`]);
    deptRow.font = { bold: true, size: 12 };
    deptRow.alignment = { vertical: 'middle' };
    worksheet.mergeCells(2, 1, 2, 16);
    worksheet.getRow(2).height = 30;

    // 第三行：酒店表头
    const hotelHeaders = [
      '预订/退款日期', '订单状态', '预订人', '旅客姓名', '旅客直属部门', '入住日期', '离店日期', '入住城市',
      '酒店名称', '间夜数', '平均客房单价', '销售总价', '企业支付', '服务费', '应还款总金额', '酒店开票类型'
    ];
    const headerRow = worksheet.addRow(hotelHeaders);
    headerRow.font = { bold: true, color: { argb: 'FFFFFF' } }; // 白色字体
    headerRow.alignment = { horizontal: 'center', vertical: 'middle' }; // 居中对齐
    worksheet.getRow(3).height = 30;

    // 设置表头行的填充颜色和边框
    for (let i = 1; i <= 16; i++) {
      const cell = headerRow.getCell(i);
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFF8E22' } // #FF8E22 橙色
      };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    }

    // 处理酒店数据
    const processedHotelData = hotelData.map((originalRow, rowIndex) => {
      // 计算服务费 = 系统使用费 + 酒店托管费 + 代购费
      const systemUsageFee = parseFloat(originalRow[columnMapping['系统使用费']]) || 0;
      const hotelManagementFee = parseFloat(originalRow[columnMapping['酒店托管费']]) || 0;
      const purchaseFee = parseFloat(originalRow[columnMapping['代购费']]) || 0;
      const totalServiceFee = systemUsageFee + hotelManagementFee + purchaseFee;

      // 获取订房费用、退订费用和夜数，处理可能是公式对象的情况
      let bookingFeeRaw = originalRow[columnMapping['订房费用']];
      let nightCountRaw = originalRow[columnMapping['夜数']];

      // 查找退订费用列
      const cancelFeeIndex = headers.findIndex((h: any) =>
        h && (h.toString().includes('退订费用') || h.toString().includes('退订费') || h.toString().includes('退房费用'))
      );
      let cancelFeeRaw = cancelFeeIndex !== -1 ? originalRow[cancelFeeIndex] : 0;

      // 调试日志：检查数据值（对所有行都输出）
      console.log(`=== 酒店数据处理调试信息 - 第${rowIndex + 1}行 ===`);
      console.log('原始行数据:', originalRow);
      console.log('列映射:', columnMapping);
      console.log('订房费用列索引:', columnMapping['订房费用']);
      console.log('夜数列索引:', columnMapping['夜数']);
      console.log('退订费用列索引:', cancelFeeIndex);
      console.log('订房费用原始值:', bookingFeeRaw, '类型:', typeof bookingFeeRaw);
      console.log('夜数原始值:', nightCountRaw, '类型:', typeof nightCountRaw);
      console.log('退订费用原始值:', cancelFeeRaw, '类型:', typeof cancelFeeRaw);

      // 处理可能是公式对象的情况
      let bookingFee = 0;
      let nightCount = 0;
      let cancelFee = 0;

      if (bookingFeeRaw !== null && bookingFeeRaw !== undefined) {
        if (typeof bookingFeeRaw === 'object' && bookingFeeRaw.result !== undefined) {
          // 如果是公式对象，使用计算结果
          bookingFee = parseFloat(bookingFeeRaw.result) || 0;
          console.log('订房费用是公式对象，使用结果值:', bookingFee);
        } else {
          // 如果是普通值，直接解析
          bookingFee = parseFloat(bookingFeeRaw) || 0;
          console.log('订房费用是普通值，解析结果:', bookingFee);
        }
      }

      if (nightCountRaw !== null && nightCountRaw !== undefined) {
        if (typeof nightCountRaw === 'object' && nightCountRaw.result !== undefined) {
          // 如果是公式对象，使用计算结果
          nightCount = parseFloat(nightCountRaw.result) || 0;
          console.log('夜数是公式对象，使用结果值:', nightCount);
        } else {
          // 如果是普通值，直接解析
          nightCount = parseFloat(nightCountRaw) || 0;
          console.log('夜数是普通值，解析结果:', nightCount);
        }
      }

      if (cancelFeeRaw !== null && cancelFeeRaw !== undefined) {
        if (typeof cancelFeeRaw === 'object' && cancelFeeRaw.result !== undefined) {
          // 如果是公式对象，使用计算结果
          cancelFee = parseFloat(cancelFeeRaw.result) || 0;
          console.log('退订费用是公式对象，使用结果值:', cancelFee);
        } else {
          // 如果是普通值，直接解析
          cancelFee = parseFloat(cancelFeeRaw) || 0;
          console.log('退订费用是普通值，解析结果:', cancelFee);
        }
      }

      // 计算平均客房单价 = (订房费用 + 退订费用) / 夜数
      const totalFee = bookingFee + cancelFee;
      const averageRoomPrice = nightCount > 0 ? totalFee / nightCount : 0;
      console.log('计算的平均客房单价:', averageRoomPrice, '(订房费用:', bookingFee, ', 退订费用:', cancelFee, ', 总费用:', totalFee, ', 夜数:', nightCount, ')');
      console.log('========================================\n');

      return {
        processedRow: [
          originalRow[columnMapping['记账日期']] || '', // 预订/退款日期 -> 记账日期
          originalRow[columnMapping['订单状态']] || '', // 订单状态
          originalRow[columnMapping['预订人']] || '', // 预订人
          originalRow[columnMapping['入住人']] || '', // 旅客姓名 -> 入住人
          originalRow[columnMapping['费用归属']] || '', // 旅客直属部门 -> 费用归属
          originalRow[columnMapping['入住日期']] || '', // 入住日期
          originalRow[columnMapping['离店日期']] || '', // 离店日期
          originalRow[columnMapping['酒店城市']] || '', // 入住城市 -> 酒店城市
          originalRow[columnMapping['酒店名称']] || '', // 酒店名称
          originalRow[columnMapping['夜数']] || 0, // 间夜数 -> 夜数
          averageRoomPrice, // 平均客房单价 = 订房费用 / 夜数
          undefined, // 销售总价 - 稍后设置公式
          '', // 企业支付 - 留空
          totalServiceFee, // 服务费 = 系统使用费 + 酒店托管费 + 代购费
          '', // 应还款总金额 - 留空
          '专票' // 酒店开票类型 - 写死专票
        ],
        passengerName: originalRow[columnMapping['入住人']] || '',
        checkInDate: originalRow[columnMapping['入住日期']] || '',
        originalRow: originalRow,
        serviceFee: totalServiceFee,
        averageRoomPrice: averageRoomPrice
      };
    });

    // 按旅客姓名分组和排序
    const passengerNames = Array.from(new Set(processedHotelData.map(item => item.passengerName || '未知乘客')));
    passengerNames.sort((a, b) => a.localeCompare(b, 'zh-CN'));

    const groupedByPassenger: { [key: string]: typeof processedHotelData } = {};
    passengerNames.forEach(passengerName => {
      const passengerData = processedHotelData.filter(item =>
        (item.passengerName || '未知乘客') === passengerName
      );
      groupedByPassenger[passengerName] = passengerData;
    });

    // 对每个分组按入住日期排序
    Object.keys(groupedByPassenger).forEach(passengerName => {
      groupedByPassenger[passengerName].sort((a, b) => {
        const dateA = new Date(a.checkInDate || '').getTime();
        const dateB = new Date(b.checkInDate || '').getTime();
        return dateA - dateB;
      });
    });

    // 添加数据行和小计
    let currentRowNumber = 4; // 数据从第4行开始
    Object.keys(groupedByPassenger).forEach(passengerName => {
      const passengerData = groupedByPassenger[passengerName];

      // 添加该乘客的所有数据行
      passengerData.forEach(item => {
        const newRow = worksheet.addRow(item.processedRow);
        worksheet.getRow(newRow.number).height = 30;

        // 设置销售总价公式（第12列）= 平均客房单价 × 间夜数
        const salesTotalCell = newRow.getCell(12);
        salesTotalCell.value = {
          formula: `=K${currentRowNumber}*J${currentRowNumber}`,
          result: 0
        };

        // 设置企业支付公式（第13列），直接引用销售总价
        const enterprisePaymentCell = newRow.getCell(13);
        enterprisePaymentCell.value = {
          formula: `=L${currentRowNumber}`,
          result: 0
        };

        // 设置应还款总金额公式（第15列）= 销售总价 + 服务费
        const totalAmountCell = newRow.getCell(15);
        totalAmountCell.value = {
          formula: `=L${currentRowNumber}+N${currentRowNumber}`,
          result: 0
        };

        currentRowNumber++;
      });

      // 添加小计行
      if (passengerData.length > 0) {
        const startRow = currentRowNumber - passengerData.length;
        const endRow = currentRowNumber - 1;

        const subtotalRow = worksheet.addRow([
          '', '', '', '', '', '', '', '', '', '', '', '', '', '', // 前14列全部留空，将合并显示"小计"
          { // 应还款总金额（第15列）设置求和公式
            formula: `=SUM(O${startRow}:O${endRow})`,
            result: 0
          },
          '' // 酒店开票类型留空（第16列）
        ]);

        // 合并小计行的第1列到第14列，用于显示"小计"文字
        worksheet.mergeCells(subtotalRow.number, 1, subtotalRow.number, 14);

        // 设置小计行样式
        subtotalRow.font = { bold: true };
        subtotalRow.getCell(1).value = '小计'; // 在合并后的第一个单元格中设置"小计"文字
        subtotalRow.getCell(1).alignment = { horizontal: 'right', vertical: 'middle' }; // 小计文字右对齐垂直居中
        subtotalRow.getCell(15).numFmt = '#,##0.00'; // 应还款总金额数字格式

        currentRowNumber++;
      }
    });

    // 添加总计行
    const totalStartRow = 4; // 数据开始行
    const totalEndRow = currentRowNumber - 1; // 最后一行数据（包括小计行）

    const totalRow = worksheet.addRow([
      '', '', '', '', // 预订/退款日期、订单状态、预订人、旅客姓名留空
      '', // 旅客直属部门留空
      '', '', '', '', // 入住日期、离店日期、入住城市、酒店名称留空
      { // 间夜数（第10列）求和
        formula: `=SUM(J${totalStartRow}:J${totalEndRow})`,
        result: 0
      },
      { // 平均客房单价（第11列）计算加权平均
        formula: `=IF(SUM(J${totalStartRow}:J${totalEndRow})=0,0,SUMPRODUCT(K${totalStartRow}:K${totalEndRow},J${totalStartRow}:J${totalEndRow})/SUM(J${totalStartRow}:J${totalEndRow}))`,
        result: 0
      },
      { // 销售总价（第12列）求和
        formula: `=SUM(L${totalStartRow}:L${totalEndRow})`,
        result: 0
      },
      { // 企业支付（第13列）求和
        formula: `=SUM(M${totalStartRow}:M${totalEndRow})`,
        result: 0
      },
      { // 服务费（第14列）求和
        formula: `=SUM(N${totalStartRow}:N${totalEndRow})`,
        result: 0
      },
      { // 应还款总金额（第15列）对小计行求和
        formula: `=SUMIF(A${totalStartRow}:A${totalEndRow},"小计",O${totalStartRow}:O${totalEndRow})`,
        result: 0
      },
      '' // 酒店开票类型留空
    ]);
    worksheet.getRow(totalRow.number).height = 30;

    // 设置总计行样式
    totalRow.font = { bold: true };
    totalRow.getCell(5).alignment = { horizontal: 'right' }; // 总计文字右对齐

    // 合并总计行的第1列到第9列
    worksheet.mergeCells(totalRow.number, 1, totalRow.number, 9);

    // 设置总计行的数字格式
    totalRow.getCell(10).numFmt = '#,##0.00'; // 间夜数
    totalRow.getCell(11).numFmt = '#,##0.00'; // 平均客房单价
    totalRow.getCell(12).numFmt = '#,##0.00'; // 销售总价
    totalRow.getCell(13).numFmt = '#,##0.00'; // 企业支付
    totalRow.getCell(14).numFmt = '#,##0.00'; // 服务费
    totalRow.getCell(15).numFmt = '#,##0.00'; // 应还款总金额

    // 添加签名行（最后一行的下一行）
    const signatureRow = worksheet.addRow([
      '', '经办人：', '', '', '审核人：', '', '', '日期：', '', '', '部门负责人审批：', '', '', ''
    ]);

    // 设置签名行样式
    signatureRow.alignment = { vertical: 'middle' }; // 签名行垂直居中
    worksheet.getRow(signatureRow.number).height = 66; // 签名行高度设置为66磅

    // 合并单元格用于签名信息
    worksheet.mergeCells(signatureRow.number, 2, signatureRow.number, 4); // 合并B-D列（第2-4列）经办人
    worksheet.mergeCells(signatureRow.number, 5, signatureRow.number, 7); // 合并E-G列（第5-7列）审核人
    worksheet.mergeCells(signatureRow.number, 8, signatureRow.number, 10); // 合并H-J列（第8-10列）日期
    worksheet.mergeCells(signatureRow.number, 11, signatureRow.number, 13); // 合并K-M列（第11-13列）部门负责人审批

    // 设置第四行到总计行的样式（跳过第3行表头）
    for (let i = 4; i <= totalRow.number; i++) {
      const row = worksheet.getRow(i);

      // 检查是否是小计行，如果是则设置行高为30磅
      const cell = row.getCell(1); // 第1列是小计标识列
      if (cell.value && cell.value.toString() === '小计') {
        row.height = 30;
      }

      // 设置每个单元格的样式
      for (let j = 1; j <= 16; j++) {
        const cell = row.getCell(j);

        // 如果是小计行的第一个单元格，保持原有的右对齐设置
        const isSubtotalRow = cell.value && cell.value.toString() === '小计';

        if (!isSubtotalRow) {
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        }

        cell.font = { size: 10 };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      }
    }

    // 收集所有数据行用于智能列宽计算
    const allHotelData = hotelData.map(row => [
      row[columnMapping['预订/退款日期']] || '',
      row[columnMapping['订单状态']] || '',
      row[columnMapping['预订人']] || '',
      row[columnMapping['旅客姓名']] || '',
      row[columnMapping['旅客直属部门']] || '',
      row[columnMapping['入住日期']] || '',
      row[columnMapping['离店日期']] || '',
      row[columnMapping['入住城市']] || '',
      row[columnMapping['酒店名称']] || '',
      row[columnMapping['间夜数']] || 0,
      row[columnMapping['平均客房单价']] || 0,
      row[columnMapping['销售总价']] || 0,
      row[columnMapping['企业支付']] || 0,
      row[columnMapping['服务费']] || 0,
      row[columnMapping['应还款总金额']] || 0,
      row[columnMapping['酒店开票类型']] || ''
    ]);

    setSmartColumnWidths(worksheet, hotelHeaders, allHotelData, 'hotel');

    // 设置金额列的数字格式
    worksheet.getColumn(10).numFmt = '#,##0.00'; // 间夜数
    worksheet.getColumn(11).numFmt = '#,##0.00'; // 平均客房单价
    worksheet.getColumn(12).numFmt = '#,##0.00'; // 销售总价
    worksheet.getColumn(13).numFmt = '#,##0.00'; // 企业支付
    worksheet.getColumn(14).numFmt = '#,##0.00'; // 服务费
    worksheet.getColumn(15).numFmt = '#,##0.00'; // 应还款总金额

    // 生成文件
    const fileName = `${fullDepartmentName}.xlsx`;
    generatedFiles.value.push({
      fileName,
      departmentName: fullDepartmentName,
      rowCount: hotelData.length,
      workbook: newWorkbook
    });

    console.log(`已准备生成酒店文件: ${fileName}, 包含 ${hotelData.length} 条数据`);

  } catch (error) {
    console.error(`生成 ${fullDepartmentName} 酒店部门报告失败:`, error);
  }
};

const generateExcelFiles = async () => {
  if (!originalWorkbook.value || allSheetData.value.length === 0) {
    ElMessage.error("请先上传并处理Excel文件");
    return;
  }

  generating.value = true;

  try {
    console.log('开始生成ZIP文件...');

    // 创建ZIP文件
    const zip = new JSZip();
    let totalFiles = 0;

    // 生成部门拆分文件并添加到ZIP
    if (generatedFiles.value.length > 0) {
      console.log(`生成 ${generatedFiles.value.length} 个部门拆分文件`);

      for (const fileData of generatedFiles.value) {
        const excelBuffer = await fileData.workbook.xlsx.writeBuffer();
        zip.file(fileData.fileName, excelBuffer);
        totalFiles++;
        console.log(`已添加到ZIP: ${fileData.fileName}, 包含 ${fileData.rowCount} 条数据`);
      }
    }


    // 生成ZIP文件
    const zipContent = await zip.generateAsync({ type: "blob" });
    const zipFileName = `华安保险处理结果_${new Date().toISOString().slice(0, 10)}.zip`;

    saveAs(zipContent, zipFileName);

    console.log(`ZIP文件生成完成: ${zipFileName}`);
    console.log(`部门拆分文件数: ${generatedFiles.value.length}`);
    console.log(`总文件数: ${totalFiles}`);

    ElMessage.success(`成功生成ZIP包！包含 ${totalFiles} 个部门拆分文件`);

  } catch (error) {
    console.error("生成ZIP文件失败:", error);
    ElMessage.error("生成ZIP文件失败");
  } finally {
    generating.value = false;
  }
};
</script>

<style scoped>
.bill-split-container {
  padding: 20px;
}

.upload-section {
  margin-bottom: 30px;
}

.upload-dragger {
  width: 100%;
}

.data-section {
  background: white;
  border-radius: 8px;
  box-shadow: 0 2px 12px rgba(0, 0, 0, 0.1);
  padding: 20px;
}

.data-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 20px;
}

.data-header h3 {
  margin: 0;
  color: #303133;
}

.data-summary {
  margin-bottom: 20px;
}

.data-table {
  margin-top: 20px;
}

.header-buttons {
  display: flex;
  gap: 10px;
}

.department-results {
  margin-top: 30px;
}

.department-results h4 {
  margin: 0 0 15px 0;
  color: #303133;
  font-size: 16px;
}
</style>
