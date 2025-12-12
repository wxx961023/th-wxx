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
          title="数据概览"
          type="info"
          :description="`已读取 ${allSheetData.length} 个工作表,共 ${getTotalRows()} 行数据`"
          show-icon
        />

        <el-alert
          v-if="matchedRows.length > 0"
          title="通用产品处理结果"
          type="success"
          :description="`已匹配 ${matchedRows.length} 条通用产品记录，金额已合并`"
          show-icon
          style="margin-top: 10px"
        />

        <el-alert
          v-if="generatedFiles.length > 0"
          title="账单拆分结果"
          type="info"
          :description="`已按部门拆分为 ${generatedFiles.length} 个文件，共处理数据 ${getTotalFlightRows()} 条`"
          show-icon
          style="margin-top: 10px"
        />
      </div>

      <div class="data-table">
        <el-table :data="allSheetData" border style="width: 100%">
          <el-table-column prop="name" label="工作表名称" width="200" />
          <el-table-column prop="rowCount" label="数据行数" width="120" />
          <el-table-column prop="columnCount" label="列数" width="120" />
          <el-table-column label="预览">
            <template #default="scope">
              <el-button size="small" @click="previewSheet(scope.row)">
                查看数据
              </el-button>
            </template>
          </el-table-column>
        </el-table>
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
              <el-tag :type="scope.row.departmentName.includes('火车') ? 'warning' : 'success'" size="small">
                {{ scope.row.departmentName.includes('火车') ? '火车票' : '机票' }}
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
const processedWorkbook = ref<any>(null);
const matchedRows = ref<any[]>([]); // 记录匹配的行
const generatedFiles = ref<any[]>([]); // 记录生成的文件

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

          // 处理通用产品数据
          processUniversalProducts(workbook);

          // 异步处理机票和火车票数据
          const processAllData = async () => {
            // 处理国内和国际机票按部门拆分
            const domesticResult = processDomesticFlights(workbook);
            const internationalResult = processInternationalFlights(workbook);
            const trainResult = processDomesticTrains(workbook);

            // 合并机票和火车票的数据
            const mergedDepartmentData: { [key: string]: Array<{data: any[], type: string}> } = {};
            const allDepartments = new Set([
              ...Object.keys(domesticResult.departmentData),
              ...Object.keys(internationalResult.departmentData),
              ...Object.keys(trainResult.departmentData)
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
            });

            // 生成合并后的部门报告
            const columnMappings = {
              domestic: domesticResult.columnMapping,
              international: internationalResult.columnMapping,
              train: trainResult.columnMapping
            };

            // 分别处理火车票和机票数据
            const departments = Object.keys(mergedDepartmentData);
            for (const dept of departments) {
              const deptData = mergedDepartmentData[dept];
              if (deptData.length > 0) {
                // 分离火车票和机票数据
                const trainData = deptData.filter(item => item.type === 'train');
                const flightData = deptData.filter(item => item.type === 'domestic-flight' || item.type === 'international-flight');

                // 生成火车票报告
                if (trainData.length > 0) {
                  const trainRows = trainData.map(item => item.data);
                  await generateTrainDepartmentReport(dept, trainRows, columnMappings.train);
                }

                // 生成机票报告
                if (flightData.length > 0) {
                  await generateFlightDepartmentReport(dept, flightData, columnMappings);
                }
              }
            }

            const totalProcessedRows = domesticResult.processedRows + internationalResult.processedRows + trainResult.processedRows;
            ElMessage.success(`处理完成！共处理 ${totalProcessedRows} 行数据（国内机票 ${domesticResult.processedRows} 行，国际机票 ${internationalResult.processedRows} 行，国内火车 ${trainResult.processedRows} 行），分成 ${departments.length} 个部门`);
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

const getTotalFlightRows = () => {
  return generatedFiles.value.reduce((sum, file) => sum + file.rowCount, 0);
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
    'XIANGONG/LIU': '刘现功',
    'RONGGUANG/LI': '李荣光'
  };

  const upperName = englishName.toString().trim().toUpperCase();
  return nameMap[upperName] || englishName;
};


// 处理通用产品数据
const processUniversalProducts = (workbook: any) => {
  console.log('=== 开始处理通用产品数据 ===');

  // 获取通用产品工作表
  const universalSheet = workbook.getWorksheet('通用产品');
  if (!universalSheet) {
    console.log('未找到"通用产品"工作表');
    return;
  }

  console.log('找到"通用产品"工作表');

  // 读取通用产品数据（包含空单元格）
  const universalData: any[][] = [];
  universalSheet.eachRow((row: any, rowNumber: number) => {
    const rowData: any[] = [];
    // 使用 includeEmpty 参数确保包含所有单元格（包括空的）
    row.eachCell({ includeEmpty: true }, (cell: any, colNumber: number) => {
      rowData.push(cell.value);
    });
    universalData.push(rowData);
  });

  console.log(`通用产品数据行数: ${universalData.length}`);

  // 假设第一行是表头
  const headers = universalData[0];
  console.log('通用产品表头:', headers);

  // 查找列索引
  const productTypeIndex = headers.findIndex((h: any) =>
    h && h.toString().includes('产品类型')
  );
  const remarkIndex = headers.findIndex((h: any) =>
    h && h.toString().includes('产品备注')
  );
  const totalAmountIndex = headers.findIndex((h: any) =>
    h && h.toString().includes('总金额')
  );

  console.log(`列索引 - 产品类型: ${productTypeIndex}, 产品备注: ${remarkIndex}, 总金额: ${totalAmountIndex}`);

  if (productTypeIndex === -1 || remarkIndex === -1 || totalAmountIndex === -1) {
    console.error('未找到必要的列: 产品类型、产品备注或总金额');
    ElMessage.warning('通用产品工作表格式不正确，缺少必要的列');
    return;
  }

  // 处理每一行数据（从第二行开始，跳过表头）
  let processedCount = 0;
  let matchedCount = 0;

  for (let i = 1; i < universalData.length; i++) {
    const row = universalData[i];
    const productType = row[productTypeIndex]?.toString().trim();
    const remark = row[remarkIndex]?.toString().trim();
    const totalAmount = parseFloat(row[totalAmountIndex]) || 0;

    if (!productType || !remark) {
      continue;
    }

    console.log(`\n处理第 ${i + 1} 行: 产品类型="${productType}", 备注="${remark}", 总金额=${totalAmount}`);

    // 提取订单号（前9位数字）
    const orderNumberMatch = remark.match(/^(\d{9})/);
    if (!orderNumberMatch) {
      console.log(`  无法提取订单号（需要9位数字开头）`);
      continue;
    }

    const orderNumber = orderNumberMatch[1];
    console.log(`  提取到订单号: ${orderNumber}`);

    // 根据产品类型找到对应的工作表
    const targetSheet = workbook.getWorksheet(productType);
    if (!targetSheet) {
      console.log(`  未找到工作表: ${productType}`);
      continue;
    }

    console.log(`  找到目标工作表: ${productType}`);

    // 在目标工作表中查找订单号列并匹配（包含空单元格）
    const targetData: any[][] = [];
    targetSheet.eachRow((row: any, rowNumber: number) => {
      const rowData: any[] = [];
      // 使用 includeEmpty 参数确保包含所有单元格（包括空的）
      row.eachCell({ includeEmpty: true }, (cell: any, colNumber: number) => {
        rowData.push(cell.value);
      });
      targetData.push(rowData);
    });

    // 查找订单号列
    const targetHeaders = targetData[0];
    console.log(`  目标工作表表头:`, targetHeaders);

    const orderNumberColIndex = targetHeaders.findIndex((h: any) =>
      h && h.toString().includes('订单号')
    );
    const targetTotalAmountIndex = targetHeaders.findIndex((h: any) =>
      h && h.toString().includes('总金额')
    );

    console.log(`  目标工作表列索引 - 订单号: ${orderNumberColIndex}, 总金额: ${targetTotalAmountIndex}`);

    if (orderNumberColIndex === -1 || targetTotalAmountIndex === -1) {
      console.log(`  目标工作表缺少必要的列`);
      continue;
    }

    // 查找匹配的行
    let found = false;
    console.log(`  开始在 ${targetData.length - 1} 行数据中查找订单号: ${orderNumber}`);

    for (let j = 1; j < targetData.length; j++) {
      const targetRow = targetData[j];
      const targetOrderNumberRaw = targetRow[orderNumberColIndex];
      const targetOrderNumber = targetOrderNumberRaw?.toString().trim();

      // 调试：打印前5行的订单号进行对比
      if (j <= 5) {
        console.log(`    行 ${j + 1}: 订单号="${targetOrderNumber}" (原始值: ${targetOrderNumberRaw}, 类型: ${typeof targetOrderNumberRaw})`);
      }

      if (targetOrderNumber === orderNumber) {
        console.log(`  ✅ 找到匹配行: 第 ${j + 1} 行, 订单号: ${targetOrderNumber}`);

        const originalAmount = parseFloat(targetRow[targetTotalAmountIndex]) || 0;
        const newAmount = originalAmount + totalAmount;

        console.log(`    原金额: ${originalAmount}, 通用产品金额: ${totalAmount}, 新金额: ${newAmount}`);

        // 更新金额 - 使用公式形式显示
        const excelRowNumber = j + 1;
        const excelColNumber = targetTotalAmountIndex + 1;
        const targetCell = targetSheet.getRow(excelRowNumber).getCell(excelColNumber);

        // 设置为公式：如果通用产品金额为正数用加法，为负数用减法
        let formula: string;
        if (totalAmount >= 0) {
          formula = `=${originalAmount}+${totalAmount}`;
        } else {
          // 负数转为正数显示为减法
          formula = `=${originalAmount}${totalAmount}`; // totalAmount本身带负号
        }
        console.log(`    更新单元格: 行${excelRowNumber}, 列${excelColNumber}, 公式: ${formula}`);
        targetCell.value = { formula: formula };

        matchedCount++;
        matchedRows.value.push({
          universalRow: i + 1,
          targetSheet: productType,
          targetRow: excelRowNumber,
          orderNumber,
          originalAmount,
          addedAmount: totalAmount,
          newAmount
        });

        found = true;
        break; // 找到匹配后跳出循环
      }
    }

    if (!found) {
      console.log(`  ❌ 未找到匹配的订单号: "${orderNumber}"`);
      console.log(`  提示: 请检查目标工作表中是否存在该订单号，注意检查数据格式和空格`);
    }

    processedCount++;
  }

  // 所有数据处理完成
  console.log(`\n=== 处理完成 ===`);
  console.log(`处理行数: ${processedCount}`);
  console.log(`匹配成功: ${matchedCount}`);
  console.log(`匹配详情:`, matchedRows.value);

  processedWorkbook.value = workbook;

  ElMessage.success(`数据处理完成！匹配 ${matchedCount} 条记录`);
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
    titleRow.font = { bold: true, size: 16 };
    titleRow.alignment = { horizontal: 'center' };
    worksheet.mergeCells(1, 1, 1, 17);

    // 第二行：部门信息
    const deptRow = worksheet.addRow([`部门：${departmentName}`]);
    deptRow.font = { bold: true };
    worksheet.mergeCells(2, 1, 2, 17);

    // 第三行：火车票表头
    const trainHeaders = [
      '预订/退款日期', '订单状态', '预订人', '旅客姓名', '旅客直属部门', '行程', '车次', '出发时间',
      '坐席', '座位编号', '车票单价', '改签费', '退票费', '销售总价', '企业支付', '服务费', '应还款总金额'
    ];
    const headerRow = worksheet.addRow(trainHeaders);
    headerRow.font = { bold: true };

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

        // 火车票的销售总价和企业支付留空，不设置公式
        currentRowNumber++;
      });

      // 添加小计行
      if (passengerData.length > 0) {
        const startRow = currentRowNumber - passengerData.length;
        const endRow = currentRowNumber - 1;

        const subtotalRow = worksheet.addRow([
          '', '', '', '', // 预订/退款日期、订单状态、预订人、旅客姓名留空
          '小计', // 旅客直属部门列显示"小计"
          '', '', '', '', // 行程、车次、出发时间、坐席留空
          '', '', '', '', '', '', // 座位编号、车票单价、改签费、退票费、销售总价、企业支付留空
          '', // 服务费留空
          { // 应还款总金额（第17列）设置求和公式
            formula: `=SUM(Q${startRow}:Q${endRow})`,
            result: 0
          }
        ]);

        // 设置小计行样式
        subtotalRow.font = { bold: true };
        subtotalRow.getCell(5).alignment = { horizontal: 'right' }; // 小计文字右对齐
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
        formula: `=SUMIF(E${totalStartRow}:E${totalEndRow},"小计",Q${totalStartRow}:Q${totalEndRow})`,
        result: 0
      }
    ]);

    // 设置总计行样式
    totalRow.font = { bold: true };
    totalRow.getCell(5).alignment = { horizontal: 'right' }; // 总计文字右对齐
    totalRow.getCell(17).numFmt = '#,##0.00'; // 应还款总金额数字格式

    // 设置列宽
    for (let i = 1; i <= 17; i++) {
      worksheet.getColumn(i).width = 15;
    }

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
  const fullDepartmentName = `商务-机票-${departmentName}`;

  try {
    console.log(`=== 生成 ${fullDepartmentName} 机票部门报告 ===`);

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
    const titleRow = worksheet.addRow([`华安保险${year}年${month}月机票对账单`]);
    titleRow.font = { bold: true, size: 16 };
    titleRow.alignment = { horizontal: 'center' };
    worksheet.mergeCells(1, 1, 1, 20);

    // 第二行：部门信息
    const deptRow = worksheet.addRow([`部门：${departmentName}`]);
    deptRow.font = { bold: true };
    worksheet.mergeCells(2, 1, 2, 20);

    // 第三行：表头
    const headers = [
      '动支号', '票号', '机票状态', '预订人', '旅客姓名', '旅客直属部门', '行程', '航班号',
      '起飞时间', '票销售价', '机建费(国内)', '燃油费(国内)', '改签费', '升舱费', '退票费', '销售总价',
      '企业支付', '服务费', '应还款总金额', '签字确认'
    ];
    const headerRow = worksheet.addRow(headers);
    headerRow.font = { bold: true };

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

      return {
        processedRow: [
          '', // 动支号 - 留空
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
          '', '', '', '', // 动支号、票号、机票状态、预订人留空
          '小计', // 旅客姓名列只显示"小计"
          '', '', '', '', // 旅客直属部门、行程、航班号、起飞时间留空
          '', '', '', '', '', '', // 票销售价到退票费留空
          '', // 销售总价留空
          '', // 企业支付留空
          '', // 服务费留空
          { // 应还款总金额（第19列）设置求和公式
            formula: `=SUM(S${startRow}:S${endRow})`,
            result: 0
          },
          '' // 签字确认留空
        ]);

        // 设置小计行样式
        subtotalRow.font = { bold: true };
        subtotalRow.getCell(5).alignment = { horizontal: 'right' }; // 小计文字右对齐
        subtotalRow.getCell(19).numFmt = '#,##0.00'; // 设置应还款总金额的数字格式

        currentRowNumber++;
      }
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
        formula: `=SUMIF(E${totalStartRow}:E${totalEndRow},"小计",S${totalStartRow}:S${totalEndRow})`,
        result: 0
      },
      '' // 签字确认留空
    ]);

    // 设置总计行样式
    totalRow.font = { bold: true };
    totalRow.getCell(5).alignment = { horizontal: 'right' }; // 总计文字右对齐

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

    // 设置列宽
    for (let i = 1; i <= 20; i++) {
      worksheet.getColumn(i).width = 15;
    }

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
    const fileName = `${fullDepartmentName}.xlsx`;
    generatedFiles.value.push({
      fileName,
      departmentName: fullDepartmentName,
      rowCount: totalRows,
      workbook: newWorkbook
    });

    console.log(`已准备生成机票文件: ${fileName}, 包含 ${totalRows} 条数据`);

  } catch (error) {
    console.error(`生成 ${fullDepartmentName} 机票部门报告失败:`, error);
  }
};

const generateExcelFiles = async () => {
  if (!processedWorkbook.value || allSheetData.value.length === 0) {
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

    // 生成通用产品处理文件并添加到ZIP
    const workbook = processedWorkbook.value;
    const excelBuffer = await workbook.xlsx.writeBuffer();
    const fileName = `华安保险_通用产品处理结果_${new Date().toISOString().slice(0, 10)}.xlsx`;
    zip.file(fileName, excelBuffer);
    totalFiles++;

    console.log(`已添加到ZIP: ${fileName}`);

    // 生成ZIP文件
    const zipContent = await zip.generateAsync({ type: "blob" });
    const zipFileName = `华安保险处理结果_${new Date().toISOString().slice(0, 10)}.zip`;

    saveAs(zipContent, zipFileName);

    console.log(`ZIP文件生成完成: ${zipFileName}`);
    console.log(`匹配记录数: ${matchedRows.value.length}`);
    console.log(`部门拆分文件数: ${generatedFiles.value.length}`);
    console.log(`总文件数: ${totalFiles}`);

    if (matchedRows.value.length > 0) {
      console.log('匹配详情:', matchedRows.value);
    }

    ElMessage.success(`成功生成ZIP包！包含 ${totalFiles} 个文件 (${generatedFiles.value.length} 个部门文件 + 1 个通用产品处理文件)`);

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
