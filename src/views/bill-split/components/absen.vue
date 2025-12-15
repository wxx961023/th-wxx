<template>
  <div class="absen-bill-split-container">
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
          将艾比森账单Excel文件拖到此处,或<em>点击上传</em>
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
        <h3>艾比森账单数据</h3>
        <div class="header-buttons">
          <el-button
            type="primary"
            :loading="processing"
            @click="processAccountantInfo"
            :disabled="!showData"
          >
            {{ processing ? "处理中..." : "填充对账人信息并下载" }}
          </el-button>
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
          <el-table-column prop="contactName" label="对账人" width="120" />
          <el-table-column prop="contactEmail" label="邮箱" width="200" />
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
import { getAbsenContactByDepartment } from '../absenDepartmentContactConfig';

defineOptions({
  name: "AbsenBillSplit"
});

const uploadedFile = ref<File | null>(null);
const allSheetData = ref<any[]>([]);
const loading = ref(false);
const showData = ref(false);
const processing = ref(false);
const generating = ref(false);
const originalWorkbook = ref<any>(null);
const processedWorkbook = ref<any>(null); // 处理后的工作簿
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
          console.log('=== 艾比森Excel文件加载成功 ===');
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

          // 重置处理后的工作簿
          processedWorkbook.value = null;
          generatedFiles.value = [];

          ElMessage.success(
            `成功读取 ${sheetInfoArray.length} 个工作表！可以开始处理对账人信息。`
          );

          showData.value = true;
          loading.value = false;
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

// 根据部门名称获取对账人信息，支持从费用归属列备用查询
const getContactInfoWithFallback = (
  costBelongFullValue: any,
  costBelongValue: any,
  sheetName: string,
  rowNumber: number
) => {
  // 首先尝试从费用归属（全路径）获取部门信息
  if (costBelongFullValue) {
    const costBelongFullStr = costBelongFullValue.toString().trim();
    const parts = costBelongFullStr.split('-');

    if (parts.length >= 2) {
      let departmentName = parts[1].trim(); // 默认取第二个值

      // 如果第二个值是"国际服务运营部"，则取第三个值
      if (parts[1].trim() === '国际服务运营部' && parts.length > 2) {
        departmentName = parts[2].trim();
        console.log(`工作表 ${sheetName} 第 ${rowNumber} 行: 检测到国际服务运营部，改用第三个值"${departmentName}"`);
      }

      const contactInfo = getAbsenContactByDepartment(departmentName);
      if (contactInfo && contactInfo.accountant) {
        console.log(`工作表 ${sheetName} 第 ${rowNumber} 行: 费用归属（全路径）="${costBelongFullStr}" -> 部门="${departmentName}" -> 对账人="${contactInfo.accountant}"`);
        return { contactInfo, source: '费用归属（全路径）', departmentName };
      }
    }
  }

  // 如果费用归属（全路径）为空或没有找到匹配的值，尝试从费用归属列直接匹配
  if (costBelongValue && costBelongValue.toString().trim()) {
    const costBelongStr = costBelongValue.toString().trim();
    const contactInfo = getAbsenContactByDepartment(costBelongStr);

    if (contactInfo && contactInfo.accountant) {
      console.log(`工作表 ${sheetName} 第 ${rowNumber} 行: 费用归属="${costBelongStr}" -> 对账人="${contactInfo.accountant}" (备用查询成功)`);
      return { contactInfo, source: '费用归属', departmentName: costBelongStr };
    } else {
      console.log(`工作表 ${sheetName} 第 ${rowNumber} 行: 费用归属="${costBelongStr}" -> 未找到对账人 (备用查询失败)`);
    }
  }

  // 都没有找到匹配的值
  console.log(`工作表 ${sheetName} 第 ${rowNumber} 行: 两种查询方式都未找到对账人`);
  return { contactInfo: null, source: 'none', departmentName: '' };
};

// 处理对账人信息
const processAccountantInfo = async () => {
  if (!originalWorkbook.value) {
    ElMessage.error("请先上传Excel文件");
    return;
  }

  processing.value = true;

  try {
    console.log('=== 开始处理对账人信息 ===');

    // 创建工作簿的副本进行修改
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(await originalWorkbook.value.xlsx.writeBuffer());

    const targetSheets = ['国内机票', '国际机票', '国内酒店', '国际酒店', '通用产品'];
    let totalProcessed = 0;
    let totalYellowCells = 0;

    // 处理每个目标工作表
    for (const sheetName of targetSheets) {
      const worksheet = workbook.getWorksheet(sheetName);
      if (!worksheet) {
        console.log(`未找到工作表: ${sheetName}`);
        continue;
      }

      console.log(`处理工作表: ${sheetName}`);

      // 获取表头
      const headers: any[] = [];
      const headerRow = worksheet.getRow(1);
      headerRow.eachCell({ includeEmpty: true }, (cell: any) => {
        headers.push(cell.value);
      });

      console.log(`工作表 ${sheetName} 表头:`, headers.map((h, i) => `${i}: "${h}"`));

      // 查找"费用归属（全路径）"列的索引 - 增加模糊匹配
      let costBelongFullIndex = headers.findIndex((h: any) =>
        h && h.toString() === '费用归属（全路径）'
      );

      // 如果精确匹配没找到，尝试模糊匹配
      if (costBelongFullIndex === -1) {
        costBelongFullIndex = headers.findIndex((h: any) =>
          h && h.toString().includes('费用归属')
        );
        if (costBelongFullIndex !== -1) {
          console.log(`工作表 ${sheetName} 使用模糊匹配找到费用归属列: "${headers[costBelongFullIndex]}"`);
        }
      }

      if (costBelongFullIndex === -1) {
        console.log(`工作表 ${sheetName} 未找到"费用归属（全路径）列，跳过`);
        console.log(`可用列名:`, headers.map((h, i) => `${i}: "${h}"`));
        continue;
      }

      console.log(`工作表 ${sheetName} "费用归属（全路径）"列索引: ${costBelongFullIndex}`);

      // 查找"费用归属"列的索引（作为备用查询）
      let costBelongIndex = headers.findIndex((h: any) =>
        h && h.toString() === '费用归属'
      );

      console.log(`工作表 ${sheetName} "费用归属"列索引: ${costBelongIndex}`);

      // 查找"预订人"或"预定人"列的索引
      const bookingPersonIndex = headers.findIndex((h: any) =>
        h && (h.toString().includes('预订人') || h.toString().includes('预定人'))
      );

      console.log(`工作表 ${sheetName} "预订人/预定人"列索引: ${bookingPersonIndex}`);
      if (bookingPersonIndex !== -1) {
        console.log(`找到预订人列: "${headers[bookingPersonIndex]}"`);
      }

      // 查找"对账人"列是否存在，如果不存在则添加
      let accountantIndex = headers.findIndex((h: any) =>
        h && h.toString().includes('对账人')
      );

      if (accountantIndex === -1) {
        if (bookingPersonIndex === -1) {
          console.log(`工作表 ${sheetName} 未找到"预订人"列，在最后一列添加"对账人"列`);
          // 如果没找到预订人列，就在最后一列添加
          accountantIndex = headers.length;
        } else {
          const foundColumnName = headers[bookingPersonIndex];
          console.log(`工作表 ${sheetName} 在"${foundColumnName}"列前添加"对账人"列，列索引: ${bookingPersonIndex}`);
          // 在预订人/预定人列前面插入对账人列
          accountantIndex = bookingPersonIndex;
        }

        // 在指定位置插入列
        for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber++) {
          const row = worksheet.getRow(rowNumber);
          // 移动该列及其之后的所有列向右一格
          for (let col = worksheet.columnCount; col >= accountantIndex + 1; col--) {
            const sourceCell = row.getCell(col);
            const targetCell = row.getCell(col + 1);
            targetCell.value = sourceCell.value;
            // 复制样式时排除填充样式，避免颜色污染
            if (sourceCell.style) {
              const styleCopy = { ...sourceCell.style };
              // 清除填充样式，避免颜色错乱
              if (styleCopy.fill) {
                delete styleCopy.fill;
              }
              targetCell.style = styleCopy;
            }
          }
          // 清空原来位置的单元格并设置默认样式
          const newCell = row.getCell(accountantIndex + 1);
          newCell.value = null;
          newCell.font = { bold: false }; // 确保新插入的列单元格没有加粗
          // 确保新单元格没有填充样式
          newCell.fill = {
            type: 'pattern',
            pattern: 'none'
          };
        }

        // 设置表头
        const headerCell = headerRow.getCell(accountantIndex + 1);
        headerCell.value = '对账人';
        headerCell.font = { bold: true }; // 表头加粗
        // 确保表头没有填充样式
        headerCell.fill = {
          type: 'pattern',
          pattern: 'none'
        };
      }

      // 由于插入了对账人列，需要调整费用归属列的索引
      let adjustedCostBelongFullIndex = costBelongFullIndex;
      let adjustedCostBelongIndex = costBelongIndex;
      if (accountantIndex !== -1 && accountantIndex <= costBelongFullIndex) {
        adjustedCostBelongFullIndex = costBelongFullIndex + 1;
        console.log(`工作表 ${sheetName} 对账人列插入在费用归属（全路径）列前，调整后费用归属（全路径）列索引: ${adjustedCostBelongFullIndex}`);
      }
      if (costBelongIndex !== -1 && accountantIndex !== -1 && accountantIndex <= costBelongIndex) {
        adjustedCostBelongIndex = costBelongIndex + 1;
        console.log(`工作表 ${sheetName} 对账人列插入在费用归属列前，调整后费用归属列索引: ${adjustedCostBelongIndex}`);
      }

      // 重新获取表头，因为可能已经插入了对账人列
      const updatedHeaders: any[] = [];
      headerRow.eachCell({ includeEmpty: true }, (cell: any) => {
        updatedHeaders.push(cell.value);
      });

      // 重新获取对账人列索引
      const updatedAccountantIndex = updatedHeaders.findIndex((h: any) =>
        h && h.toString() === '对账人'
      );

      // 处理数据行
      const rowCount = worksheet.rowCount;
      let sheetProcessed = 0;
      let sheetYellowCells = 0;

      for (let rowNumber = 2; rowNumber <= rowCount; rowNumber++) {
        const row = worksheet.getRow(rowNumber);
        const costBelongFullValue = row.getCell(adjustedCostBelongFullIndex + 1).value;

        // 获取费用归属列的值（如果存在该列）
        let costBelongValue = null;
        if (adjustedCostBelongIndex !== -1) {
          costBelongValue = row.getCell(adjustedCostBelongIndex + 1).value;
        }

        const accountantCell = row.getCell(updatedAccountantIndex + 1);

        // 使用新的备用查询逻辑
        const { contactInfo, source, departmentName } = getContactInfoWithFallback(
          costBelongFullValue,
          costBelongValue,
          sheetName,
          rowNumber
        );

        // 记录前几行的匹配情况
        if (rowNumber <= 5) {
          console.log(`工作表 ${sheetName} 第 ${rowNumber} 行: 查询来源="${source}", 部门="${departmentName}", 对账人=${contactInfo?.accountant || '未找到'}`);
        }

        if (contactInfo && contactInfo.accountant) {
          // 填写对账人姓名
          accountantCell.value = contactInfo.accountant;
          accountantCell.font = { bold: false }; // 数据行不加粗
          console.log(`工作表 ${sheetName} 第 ${rowNumber} 行: 对账人="${contactInfo.accountant}" (来源: ${source})`);
          sheetProcessed++;
        } else {
          // 如果两种查询方式都找不到对账人，设置"未找到对账人"并设置红色
          accountantCell.value = '未找到对账人';

          // 设置背景为红色
          accountantCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFF0000' } // 红色背景
          };

          // 强制重新应用样式
          accountantCell.style = {
            ...accountantCell.style,
            font: { bold: true },
            fill: accountantCell.fill
          };

          console.log(`工作表 ${sheetName} 第 ${rowNumber} 行: 两种查询方式都未找到对账人，设置红色标识`);
          sheetYellowCells++;
          totalYellowCells++;
        }
      }

      console.log(`工作表 ${sheetName} 处理完成: 处理 ${sheetProcessed} 行，${sheetYellowCells} 行未找到对账人`);
      totalProcessed += sheetProcessed;

      // 样式已经在数据填充时设置，无需单独处理
    }

    // 直接下载处理后的Excel文件
    const buffer = await workbook.xlsx.writeBuffer();
    const fileName = `艾比森账单_对账人已填充_${new Date().toISOString().slice(0, 10)}.xlsx`;
    saveAs(new Blob([buffer]), fileName);

    // 保存处理后的工作簿（可选，用于后续的ZIP生成）
    processedWorkbook.value = workbook;

    console.log(`对账人信息处理完成并下载: 总共处理 ${totalProcessed} 行对账人信息，${totalYellowCells} 行未找到对账人`);
    ElMessage.success(`处理完成并已下载！总共处理 ${totalProcessed} 行对账人信息，${totalYellowCells} 行未找到对账人（已显示"未找到对账人"）`);

  } catch (error) {
    console.error("处理对账人信息失败:", error);
    ElMessage.error("处理对账人信息失败");
  } finally {
    processing.value = false;
  }
};

const generateExcelFiles = async () => {
  if (!originalWorkbook.value || allSheetData.value.length === 0) {
    ElMessage.error("请先上传并处理Excel文件");
    return;
  }

  generating.value = true;

  try {
    console.log('开始生成艾比森账单ZIP文件...');

    // 创建ZIP文件
    const zip = new JSZip();
    let totalFiles = 0;

    // TODO: 根据拆分规则生成具体的部门Excel文件
    // 现在先生成一些示例文件来测试功能
    const sampleDepartments = [
      { name: '北美Live大客户部', contact: getAbsenContactByDepartment('北美Live大客户部') },
      { name: '欧洲业务一区', contact: getAbsenContactByDepartment('欧洲业务一区') },
      { name: '亚太二地区部管理', contact: getAbsenContactByDepartment('亚太二地区部管理') }
    ];

    for (const dept of sampleDepartments) {
      // 创建新的工作簿
      const newWorkbook = new ExcelJS.Workbook();
      const worksheet = newWorkbook.addWorksheet(dept.name);

      // 添加示例标题
      const titleRow = worksheet.addRow([`艾比森${dept.name}账单`]);
      titleRow.font = { size: 16 };
      titleRow.alignment = { horizontal: 'center' };
      worksheet.mergeCells(1, 1, 1, 5);

      // 添加对账人信息
      if (dept.contact) {
        const contactRow = worksheet.addRow([
          `对账人: ${dept.contact.accountant}`,
          `邮箱: ${dept.contact.email}`
        ]);
        contactRow.font = { size: 12 };
        worksheet.mergeCells(2, 1, 2, 5);
      }

      // 添加表头
      const headers = ['日期', '描述', '金额', '状态', '备注'];
      const headerRow = worksheet.addRow(headers);

      // 设置列宽
      worksheet.columns.forEach(column => {
        column.width = 15;
      });

      // 生成文件并添加到ZIP
      const excelBuffer = await newWorkbook.xlsx.writeBuffer();
      const fileName = `${dept.name}.xlsx`;
      zip.file(fileName, excelBuffer);

      // 记录生成的文件信息
      generatedFiles.value.push({
        fileName,
        departmentName: dept.name,
        rowCount: 0, // 示例文件暂无数据
        contactName: dept.contact?.accountant || '未配置',
        contactEmail: dept.contact?.email || '未配置'
      });

      totalFiles++;
      console.log(`已生成示例文件: ${fileName}`);
    }

    // 生成ZIP文件
    const zipContent = await zip.generateAsync({ type: "blob" });
    const zipFileName = `艾比森账单处理结果_${new Date().toISOString().slice(0, 10)}.zip`;

    saveAs(zipContent, zipFileName);

    console.log(`艾比森账单ZIP文件生成完成: ${zipFileName}`);
    console.log(`生成的部门拆分文件数: ${totalFiles}`);

    ElMessage.success(`成功生成艾比森账单ZIP包！包含 ${totalFiles} 个部门拆分文件`);

  } catch (error) {
    console.error("生成艾比森账单ZIP文件失败:", error);
    ElMessage.error("生成ZIP文件失败");
  } finally {
    generating.value = false;
  }
};
</script>

<style scoped>
.absen-bill-split-container {
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
