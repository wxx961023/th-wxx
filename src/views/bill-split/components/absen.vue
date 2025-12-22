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
          <div class="el-upload__tip">只能上传 xlsx/xls 文件,且不超过 10MB</div>
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
            @click="downloadAdjustedData"
            :disabled="!showData || generatedFiles.length === 0"
          >
            {{ generating ? "下载中..." : "下载调整完的数据" }}
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
          <el-table-column type="index" label="序号" width="60" />
          <el-table-column prop="departmentName" label="部门名称" width="200" />
          <el-table-column prop="rowCount" label="数据行数" width="120" />
          <el-table-column prop="fileName" label="生成文件名">
            <template #default="{ row }">
              <el-input
                v-model="row.fileName"
                @change="updateFileName(row)"
                placeholder="请输入文件名"
                size="small"
              />
            </template>
          </el-table-column>
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
import {
  getAbsenContactByDepartment,
  ABSEN_DEPARTMENT_TO_CONTACT_MAP
} from "../absenDepartmentContactConfig";

// 导入 absen账单概览.xlsx 文件
import overviewFileUrl from "../absen/absen账单概览.xlsx?url";

defineOptions({
  name: "AbsenBillSplit"
});

const uploadedFile = ref<File | null>(null);
const allSheetData = ref<any[]>([]);
const loading = ref(false);
const showData = ref(false);
const generating = ref(false);
const processing = ref(false); // 处理对账人信息的loading状态
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
          console.log("=== 艾比森Excel文件加载成功 ===");
          console.log(
            "所有工作表:",
            workbook.worksheets.map(ws => ws.name)
          );

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

            console.log(
              `工作表 "${worksheet.name}": ${jsonData.length} 行, ${worksheet.columnCount} 列`
            );
          });

          allSheetData.value = sheetInfoArray;
          originalWorkbook.value = workbook;

          // 重置处理后的工作簿
          processedWorkbook.value = null;
          generatedFiles.value = [];

          ElMessage.success(`成功读取 ${sheetInfoArray.length} 个工作表！`);

          showData.value = true;

          // 自动开始处理对账人信息并预生成拆分结果
          setTimeout(() => {
            console.log("自动开始处理对账人信息...");
            processAccountantInfoAndGenerateResults();
          }, 500);
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
  const isExcel =
    file.type ===
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" ||
    file.type === "application/vnd.ms-excel" ||
    file.name.endsWith(".xlsx") ||
    file.name.endsWith(".xls");

  if (!isExcel) {
    ElMessage.error("只能上传Excel文件！");
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
    const parts = costBelongFullStr.split("-");

    if (parts.length >= 2) {
      let departmentName = parts[1].trim(); // 默认取第二个值

      // 如果第二个值是"国际服务运营部"，则取第三个值
      if (parts[1].trim() === "国际服务运营部" && parts.length > 2) {
        departmentName = parts[2].trim();
        console.log(
          `工作表 ${sheetName} 第 ${rowNumber} 行: 检测到国际服务运营部，改用第三个值"${departmentName}"`
        );
      }

      const contactInfo = getAbsenContactByDepartment(departmentName);
      if (contactInfo && contactInfo.accountant) {
        console.log(
          `工作表 ${sheetName} 第 ${rowNumber} 行: 费用归属（全路径）="${costBelongFullStr}" -> 部门="${departmentName}" -> 对账人="${contactInfo.accountant}"`
        );
        return { contactInfo, source: "费用归属（全路径）", departmentName };
      }
    }
  }

  // 如果费用归属（全路径）为空或没有找到匹配的值，尝试从费用归属列直接匹配
  if (costBelongValue && costBelongValue.toString().trim()) {
    const costBelongStr = costBelongValue.toString().trim();
    const contactInfo = getAbsenContactByDepartment(costBelongStr);

    if (contactInfo && contactInfo.accountant) {
      console.log(
        `工作表 ${sheetName} 第 ${rowNumber} 行: 费用归属="${costBelongStr}" -> 对账人="${contactInfo.accountant}" (备用查询成功)`
      );
      return { contactInfo, source: "费用归属", departmentName: costBelongStr };
    } else {
      console.log(
        `工作表 ${sheetName} 第 ${rowNumber} 行: 费用归属="${costBelongStr}" -> 未找到对账人 (备用查询失败)`
      );
    }
  }

  // 都没有找到匹配的值
  console.log(
    `工作表 ${sheetName} 第 ${rowNumber} 行: 两种查询方式都未找到对账人`
  );
  return { contactInfo: null, source: "none", departmentName: "" };
};

// 处理对账人信息
const processAccountantInfo = async () => {
  if (!originalWorkbook.value) {
    ElMessage.error("请先上传Excel文件");
    return;
  }

  processing.value = true;

  try {
    console.log("=== 开始处理对账人信息 ===");

    // 创建工作簿的副本进行修改
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(await originalWorkbook.value.xlsx.writeBuffer());

    const targetSheets = [
      "国内机票",
      "国际机票",
      "国内酒店",
      "国际酒店",
      "通用产品"
    ];
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

      console.log(
        `工作表 ${sheetName} 表头:`,
        headers.map((h, i) => `${i}: "${h}"`)
      );

      // 查找"费用归属（全路径）"列的索引 - 增加模糊匹配
      let costBelongFullIndex = headers.findIndex(
        (h: any) => h && h.toString() === "费用归属（全路径）"
      );

      // 如果精确匹配没找到，尝试模糊匹配
      if (costBelongFullIndex === -1) {
        costBelongFullIndex = headers.findIndex(
          (h: any) => h && h.toString().includes("费用归属")
        );
        if (costBelongFullIndex !== -1) {
          console.log(
            `工作表 ${sheetName} 使用模糊匹配找到费用归属列: "${headers[costBelongFullIndex]}"`
          );
        }
      }

      if (costBelongFullIndex === -1) {
        console.log(`工作表 ${sheetName} 未找到"费用归属（全路径）列，跳过`);
        console.log(
          `可用列名:`,
          headers.map((h, i) => `${i}: "${h}"`)
        );
        continue;
      }

      console.log(
        `工作表 ${sheetName} "费用归属（全路径）"列索引: ${costBelongFullIndex}`
      );

      // 查找"费用归属"列的索引（作为备用查询）
      let costBelongIndex = headers.findIndex(
        (h: any) => h && h.toString() === "费用归属"
      );

      console.log(`工作表 ${sheetName} "费用归属"列索引: ${costBelongIndex}`);

      // 查找"预订人"或"预定人"列的索引
      const bookingPersonIndex = headers.findIndex(
        (h: any) =>
          h &&
          (h.toString().includes("预订人") || h.toString().includes("预定人"))
      );

      console.log(
        `工作表 ${sheetName} "预订人/预定人"列索引: ${bookingPersonIndex}`
      );
      if (bookingPersonIndex !== -1) {
        console.log(`找到预订人列: "${headers[bookingPersonIndex]}"`);
      }

      // 查找"对账人"列是否存在，如果不存在则添加
      let accountantIndex = headers.findIndex(
        (h: any) => h && h.toString().includes("对账人")
      );

      if (accountantIndex === -1) {
        if (bookingPersonIndex === -1) {
          console.log(
            `工作表 ${sheetName} 未找到"预订人"列，在最后一列添加"对账人"列`
          );
          // 如果没找到预订人列，就在最后一列添加
          accountantIndex = headers.length;
        } else {
          const foundColumnName = headers[bookingPersonIndex];
          console.log(
            `工作表 ${sheetName} 在"${foundColumnName}"列前添加"对账人"列，列索引: ${bookingPersonIndex}`
          );
          // 在预订人/预定人列前面插入对账人列
          accountantIndex = bookingPersonIndex;
        }

        // 在指定位置插入列
        for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber++) {
          const row = worksheet.getRow(rowNumber);
          // 移动该列及其之后的所有列向右一格
          for (
            let col = worksheet.columnCount;
            col >= accountantIndex + 1;
            col--
          ) {
            const sourceCell = row.getCell(col);
            const targetCell = row.getCell(col + 1);
            targetCell.value = sourceCell.value;

            // 区分数据行和表头行的样式处理
            if (rowNumber > 1) {
              // 数据行：复制完整样式但排除填充样式
              if (sourceCell.style) {
                const styleCopy = { ...sourceCell.style };
                // 清除填充样式，避免颜色错乱
                if (styleCopy.fill) {
                  delete styleCopy.fill;
                }
                targetCell.style = styleCopy;
              }
            } else {
              // 表头行：不复制任何样式，保持默认状态，让Excel自动处理
              // 这样可以避免样式污染问题
            }
          }
          // 清空原来位置的单元格并设置默认样式
          const newCell = row.getCell(accountantIndex + 1);
          newCell.value = null;
          if (rowNumber === 1) {
            // 表头行：不设置加粗，稍后统一设置
            newCell.font = { bold: false };
          } else {
            // 数据行：确保不加粗
            newCell.font = { bold: false };
          }
          // 确保新单元格没有填充样式
          newCell.fill = {
            type: "pattern",
            pattern: "none"
          };
        }

        // 重新设置表头样式，确保只有对账人列表头加粗
        headerRow.eachCell(
          { includeEmpty: true },
          (cell: any, colNumber: number) => {
            // 只对对账人列表头设置加粗，其他表头保持原样或默认不加粗
            if (colNumber === accountantIndex + 1) {
              cell.font = { bold: true }; // 对账人表头加粗
            } else {
              // 其他表头不强制设置样式，保持Excel默认状态
              cell.font = { bold: false }; // 确保其他表头不加粗
            }
            // 确保表头没有填充样式
            cell.fill = {
              type: "pattern",
              pattern: "none"
            };
          }
        );

        // 单独设置对账人列表头的值和样式
        const headerCell = headerRow.getCell(accountantIndex + 1);
        headerCell.value = "对账人";
        headerCell.font = { bold: true }; // 表头加粗
        headerCell.fill = {
          type: "pattern",
          pattern: "none"
        };
      }

      // 由于插入了对账人列，需要调整费用归属列的索引
      let adjustedCostBelongFullIndex = costBelongFullIndex;
      let adjustedCostBelongIndex = costBelongIndex;
      if (accountantIndex !== -1 && accountantIndex <= costBelongFullIndex) {
        adjustedCostBelongFullIndex = costBelongFullIndex + 1;
        console.log(
          `工作表 ${sheetName} 对账人列插入在费用归属（全路径）列前，调整后费用归属（全路径）列索引: ${adjustedCostBelongFullIndex}`
        );
      }
      if (
        costBelongIndex !== -1 &&
        accountantIndex !== -1 &&
        accountantIndex <= costBelongIndex
      ) {
        adjustedCostBelongIndex = costBelongIndex + 1;
        console.log(
          `工作表 ${sheetName} 对账人列插入在费用归属列前，调整后费用归属列索引: ${adjustedCostBelongIndex}`
        );
      }

      // 重新获取表头，因为可能已经插入了对账人列
      const updatedHeaders: any[] = [];
      headerRow.eachCell({ includeEmpty: true }, (cell: any) => {
        updatedHeaders.push(cell.value);
      });

      // 重新获取对账人列索引
      const updatedAccountantIndex = updatedHeaders.findIndex(
        (h: any) => h && h.toString() === "对账人"
      );

      // 处理数据行
      const rowCount = worksheet.rowCount;
      let sheetProcessed = 0;
      let sheetYellowCells = 0;

      for (let rowNumber = 2; rowNumber <= rowCount; rowNumber++) {
        const row = worksheet.getRow(rowNumber);
        const costBelongFullValue = row.getCell(
          adjustedCostBelongFullIndex + 1
        ).value;

        // 获取费用归属列的值（如果存在该列）
        let costBelongValue = null;
        if (adjustedCostBelongIndex !== -1) {
          costBelongValue = row.getCell(adjustedCostBelongIndex + 1).value;
        }

        // 检查是否为空行（所有列都是空的）
        let isEmptyRow = true;
        row.eachCell({ includeEmpty: false }, (cell: any) => {
          if (
            cell.value !== null &&
            cell.value !== undefined &&
            cell.value !== ""
          ) {
            isEmptyRow = false;
          }
        });

        // 如果是空行，跳过处理
        if (isEmptyRow) {
          console.log(`工作表 ${sheetName} 第 ${rowNumber} 行: 空行，跳过处理`);
          continue;
        }

        const accountantCell = row.getCell(updatedAccountantIndex + 1);

        // 使用新的备用查询逻辑
        const { contactInfo, source, departmentName } =
          getContactInfoWithFallback(
            costBelongFullValue,
            costBelongValue,
            sheetName,
            rowNumber
          );

        // 记录前几行的匹配情况
        if (rowNumber <= 5) {
          console.log(
            `工作表 ${sheetName} 第 ${rowNumber} 行: 查询来源="${source}", 部门="${departmentName}", 对账人=${contactInfo?.accountant || "未找到"}`
          );
        }

        if (contactInfo && contactInfo.accountant) {
          // 填写对账人姓名
          accountantCell.value = contactInfo.accountant;
          accountantCell.font = { bold: false }; // 数据行不加粗
          // 清除任何现有的填充样式
          accountantCell.fill = {
            type: "pattern",
            pattern: "none"
          };
          console.log(
            `工作表 ${sheetName} 第 ${rowNumber} 行: 对账人="${contactInfo.accountant}" (来源: ${source})`
          );
          sheetProcessed++;
        } else {
          // 如果两种查询方式都找不到对账人，设置"未找到对账人"并设置红色
          accountantCell.value = "未找到对账人";

          // 设置背景为红色
          accountantCell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFF0000" } // 红色背景
          };

          // 强制重新应用样式
          accountantCell.style = {
            ...accountantCell.style,
            font: { bold: true },
            fill: accountantCell.fill
          };

          console.log(
            `工作表 ${sheetName} 第 ${rowNumber} 行: 两种查询方式都未找到对账人，设置红色标识`
          );
          sheetYellowCells++;
          totalYellowCells++;
        }
      }

      console.log(
        `工作表 ${sheetName} 处理完成: 处理 ${sheetProcessed} 行，${sheetYellowCells} 行未找到对账人`
      );
      totalProcessed += sheetProcessed;

      // 设置该工作表的字体和行高样式
      // 设置表头样式
      const styleHeaderRow = worksheet.getRow(1);
      styleHeaderRow.height = 22;
      styleHeaderRow.eachCell(cell => {
        if (cell.value !== null && cell.value !== undefined) {
          cell.font = { size: 10, bold: true, name: "宋体" };
          cell.alignment = { vertical: "middle", horizontal: "center" };
          // 添加细边框
          cell.border = {
            top: { style: "thin", color: { argb: "FF000000" } },
            left: { style: "thin", color: { argb: "FF000000" } },
            bottom: { style: "thin", color: { argb: "FF000000" } },
            right: { style: "thin", color: { argb: "FF000000" } }
          };
        }
      });

      // 设置数据行样式
      for (let rowNum = 2; rowNum <= worksheet.rowCount; rowNum++) {
        const styleDataRow = worksheet.getRow(rowNum);
        styleDataRow.height = 22;
        styleDataRow.eachCell(cell => {
          if (cell.value !== null && cell.value !== undefined) {
            cell.font = { size: 10, name: "宋体" };
            cell.alignment = { vertical: "middle" };
            // 添加细边框
            cell.border = {
              top: { style: "thin", color: { argb: "FF000000" } },
              left: { style: "thin", color: { argb: "FF000000" } },
              bottom: { style: "thin", color: { argb: "FF000000" } },
              right: { style: "thin", color: { argb: "FF000000" } }
            };
          }
        });
      }
    }

    // 直接下载处理后的Excel文件
    const buffer = await workbook.xlsx.writeBuffer();
    const fileName = `艾比森账单_对账人已填充_${new Date().toISOString().slice(0, 10)}.xlsx`;
    saveAs(new Blob([buffer]), fileName);

    // 保存处理后的工作簿（可选，用于后续的ZIP生成）
    processedWorkbook.value = workbook;

    console.log(
      `对账人信息处理完成并下载: 总共处理 ${totalProcessed} 行对账人信息，${totalYellowCells} 行未找到对账人`
    );
    ElMessage.success(
      `处理完成并已下载！总共处理 ${totalProcessed} 行对账人信息，${totalYellowCells} 行未找到对账人（已显示"未找到对账人"）`
    );
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

  // 加载 absen账单概览.xlsx 文件
  const overviewWorkbook = await loadOverviewWorkbook();

  // 检查是否已经处理了对账人信息，如果没有则先处理
  if (!processedWorkbook.value) {
    console.log("检测到未处理对账人信息，自动开始处理...");
    try {
      // 创建工作簿的副本进行修改，但不下载
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(await originalWorkbook.value.xlsx.writeBuffer());

      const targetSheets = [
        "国内机票",
        "国际机票",
        "国内酒店",
        "国际酒店",
        "通用产品"
      ];
      let totalProcessed = 0;
      let totalYellowCells = 0;

      // 处理每个目标工作表 - 复制对账人信息填充逻辑，但不下载文件
      for (const sheetName of targetSheets) {
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) continue;

        // 获取表头
        const headers: any[] = [];
        const headerRow = worksheet.getRow(1);
        headerRow.eachCell({ includeEmpty: true }, (cell: any) => {
          headers.push(cell.value);
        });

        // 查找"费用归属（全路径）"列的索引
        let costBelongFullIndex = headers.findIndex(
          (h: any) => h && h.toString() === "费用归属（全路径）"
        );

        // 如果精确匹配没找到，尝试模糊匹配
        if (costBelongFullIndex === -1) {
          costBelongFullIndex = headers.findIndex(
            (h: any) => h && h.toString().includes("费用归属")
          );
        }

        if (costBelongFullIndex === -1) continue;

        // 查找"费用归属"列的索引（作为备用查询）
        let costBelongIndex = headers.findIndex(
          (h: any) => h && h.toString() === "费用归属"
        );

        // 查找"预订人"或"预定人"列的索引
        const bookingPersonIndex = headers.findIndex(
          (h: any) =>
            h &&
            (h.toString().includes("预订人") || h.toString().includes("预定人"))
        );

        // 查找"对账人"列是否存在，如果不存在则添加
        let accountantIndex = headers.findIndex(
          (h: any) => h && h.toString().includes("对账人")
        );

        if (accountantIndex === -1) {
          if (bookingPersonIndex === -1) {
            accountantIndex = headers.length;
          } else {
            accountantIndex = bookingPersonIndex;
          }

          // 在指定位置插入列
          for (
            let rowNumber = 1;
            rowNumber <= worksheet.rowCount;
            rowNumber++
          ) {
            const row = worksheet.getRow(rowNumber);
            // 移动该列及其之后的所有列向右一格
            for (
              let col = worksheet.columnCount;
              col >= accountantIndex + 1;
              col--
            ) {
              const sourceCell = row.getCell(col);
              const targetCell = row.getCell(col + 1);
              targetCell.value = sourceCell.value;

              // 区分数据行和表头行的样式处理
              if (rowNumber > 1) {
                if (sourceCell.style) {
                  const styleCopy = { ...sourceCell.style };
                  if (styleCopy.fill) {
                    delete styleCopy.fill;
                  }
                  targetCell.style = styleCopy;
                }
              }
            }
            // 清空原来位置的单元格并设置默认样式
            const newCell = row.getCell(accountantIndex + 1);
            newCell.value = null;
            newCell.font = { bold: false };
            newCell.fill = {
              type: "pattern",
              pattern: "none"
            };
          }

          // 设置表头
          const headerCell = headerRow.getCell(accountantIndex + 1);
          headerCell.value = "对账人";
          headerCell.font = { bold: true };
          headerCell.fill = {
            type: "pattern",
            pattern: "none"
          };
        }

        // 由于插入了对账人列，需要调整费用归属列的索引
        let adjustedCostBelongFullIndex = costBelongFullIndex;
        let adjustedCostBelongIndex = costBelongIndex;
        if (accountantIndex !== -1 && accountantIndex <= costBelongFullIndex) {
          adjustedCostBelongFullIndex = costBelongFullIndex + 1;
        }
        if (
          costBelongIndex !== -1 &&
          accountantIndex !== -1 &&
          accountantIndex <= costBelongIndex
        ) {
          adjustedCostBelongIndex = costBelongIndex + 1;
        }

        // 重新获取表头，因为可能已经插入了对账人列
        const updatedHeaders: any[] = [];
        headerRow.eachCell({ includeEmpty: true }, (cell: any) => {
          updatedHeaders.push(cell.value);
        });

        // 重新获取对账人列索引
        const updatedAccountantIndex = updatedHeaders.findIndex(
          (h: any) => h && h.toString() === "对账人"
        );

        // 处理数据行
        const rowCount = worksheet.rowCount;
        for (let rowNumber = 2; rowNumber <= rowCount; rowNumber++) {
          const row = worksheet.getRow(rowNumber);
          const costBelongFullValue = row.getCell(
            adjustedCostBelongFullIndex + 1
          ).value;
          let costBelongValue = null;
          if (adjustedCostBelongIndex !== -1) {
            costBelongValue = row.getCell(adjustedCostBelongIndex + 1).value;
          }

          const accountantCell = row.getCell(updatedAccountantIndex + 1);

          // 检查是否为空行（所有列都是空的）
          let isEmptyRow = true;
          row.eachCell({ includeEmpty: false }, (cell: any) => {
            if (
              cell.value !== null &&
              cell.value !== undefined &&
              cell.value !== ""
            ) {
              isEmptyRow = false;
            }
          });

          // 如果是空行，跳过处理
          if (isEmptyRow) {
            continue;
          }

          // 使用备用查询逻辑
          const { contactInfo } = getContactInfoWithFallback(
            costBelongFullValue,
            costBelongValue,
            sheetName,
            rowNumber
          );

          if (contactInfo && contactInfo.accountant) {
            accountantCell.value = contactInfo.accountant;
            accountantCell.font = { bold: false };
            // 清除任何现有的填充样式
            accountantCell.fill = {
              type: "pattern",
              pattern: "none"
            };
            totalProcessed++;
          } else {
            accountantCell.value = "未找到对账人";
            accountantCell.fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: "FFFF0000" }
            };
            accountantCell.style = {
              ...accountantCell.style,
              font: { bold: true },
              fill: accountantCell.fill
            };
            totalYellowCells++;
          }
        }
      }

      // 保存处理后的工作簿，但不下载文件
      processedWorkbook.value = workbook;
      console.log(
        `对账人信息处理完成: 总共处理 ${totalProcessed} 行对账人信息，${totalYellowCells} 行未找到对账人`
      );
    } catch (error) {
      console.error("自动处理对账人信息失败:", error);
      ElMessage.error("自动处理对账人信息失败");
      generating.value = false;
      return;
    }
  }

  try {
    console.log("=== 开始按对账人拆分艾比森账单 ===");

    // 创建ZIP文件
    const zip = new JSZip();
    let totalFiles = 0;

    // 使用已处理的工作簿，确保对账人信息已经填充
    const sourceWorkbook = processedWorkbook.value || originalWorkbook.value;

    // 收集所有对账人的数据
    const accountantDataMap = new Map<
      string,
      {
        accountant: string;
        email: string;
        data: Array<{
          sheetName: string;
          rows: any[];
        }>;
      }
    >();

    const targetSheets = [
      "国内机票",
      "国际机票",
      "国内酒店",
      "国际酒店",
      "通用产品"
    ];

    // 遍历每个工作表收集数据
    for (const sheetName of targetSheets) {
      const worksheet = sourceWorkbook.getWorksheet(sheetName);
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

      // 找到对账人列索引
      const accountantIndex = headers.findIndex(
        (h: any) => h && h.toString().includes("对账人")
      );

      if (accountantIndex === -1) {
        console.log(`工作表 ${sheetName} 未找到对账人列，跳过`);
        continue;
      }

      console.log(`工作表 ${sheetName} 对账人列索引: ${accountantIndex}`);

      // 收集数据行，按对账人分组
      const rowCount = worksheet.rowCount;
      for (let rowNumber = 2; rowNumber <= rowCount; rowNumber++) {
        const row = worksheet.getRow(rowNumber);
        const accountantCell = row.getCell(accountantIndex + 1);
        const accountantName = accountantCell.value?.toString().trim();

        if (!accountantName || accountantName === "未找到对账人") {
          continue; // 跳过空行和"未找到对账人"的数据
        }

        // 如果该对账人还未在Map中，先初始化
        if (!accountantDataMap.has(accountantName)) {
          accountantDataMap.set(accountantName, {
            accountant: accountantName,
            email: "", // 稍后填充
            data: []
          });
        }

        // 获取整行数据
        const rowData: any[] = [];
        row.eachCell({ includeEmpty: true }, (cell: any) => {
          rowData.push(cell.value);
        });

        // 添加到对应对账人的数据中
        const accountantData = accountantDataMap.get(accountantName)!;
        if (!accountantData.data.find(d => d.sheetName === sheetName)) {
          accountantData.data.push({
            sheetName,
            rows: []
          });
        }
        const sheetData = accountantData.data.find(
          d => d.sheetName === sheetName
        )!;
        sheetData.rows.push(rowData);
      }
    }

    console.log(`收集到 ${accountantDataMap.size} 个对账人的数据`);

    // 为每个对账人生成Excel文件
    for (const [accountantName, accountantInfo] of accountantDataMap) {
      // 查找对账人邮箱
      let email = "";
      for (const contact of Object.values(ABSEN_DEPARTMENT_TO_CONTACT_MAP)) {
        if (contact.accountant === accountantName) {
          email = contact.email;
          break;
        }
      }

      // 创建新的工作簿
      const newWorkbook = new ExcelJS.Workbook();

      let totalRows = 0;

      // 首先添加账单概览工作表（如果存在）
      let overviewSheet = null;
      if (overviewWorkbook && overviewWorkbook.worksheets.length > 0) {
        for (const sourceOverviewSheet of overviewWorkbook.worksheets) {
          // 使用增强版复制函数复制概览工作表
          overviewSheet = newWorkbook.addWorksheet(sourceOverviewSheet.name);
          copyWorksheetWithFormat(sourceOverviewSheet, overviewSheet);
        }
      }

      // 为每个工作表创建数据表
      for (const sheetData of accountantInfo.data) {
        const newWorksheet = newWorkbook.addWorksheet(sheetData.sheetName);

        // 复制原工作表的表头
        const sourceWorksheet = sourceWorkbook.getWorksheet(
          sheetData.sheetName
        );
        const sourceHeaderRow = sourceWorksheet!.getRow(1);

        // 创建表头行并设置样式
        const newHeaderRow = newWorksheet.getRow(1);
        newHeaderRow.height = 22;

        // 先复制表头数据，然后立即设置样式
        const headerData: any[] = [];
        sourceHeaderRow.eachCell({ includeEmpty: true }, (cell: any) => {
          headerData.push(cell.value);
        });

        // 设置表头单元格并应用样式
        for (let col = 0; col < headerData.length; col++) {
          const cell = newHeaderRow.getCell(col + 1);
          cell.value = headerData[col];
          cell.font = { size: 10, bold: true, name: "宋体" }; // 直接设置
          cell.alignment = { vertical: "middle", horizontal: "center" };
        }

        // 复制数据行并应用样式
        for (let i = 0; i < sheetData.rows.length; i++) {
          const rowData = sheetData.rows[i];
          const newRow = newWorksheet.getRow(i + 2);
          newRow.height = 22;

          // 设置数据并应用样式
          for (let j = 0; j < rowData.length; j++) {
            const cell = newRow.getCell(j + 1);
            cell.value = rowData[j];
            cell.font = { size: 10, name: "宋体" }; // 直接设置
            cell.alignment = { vertical: "middle" };
          }
        }

        // 设置列宽
        newWorksheet.columns.forEach(column => {
          column.width = 15;
        });

        // 强制提交工作表更改
        (newWorksheet.model as any).rows.forEach((rowModel: any) => {
          if (rowModel) {
            if (!rowModel.ht || rowModel.ht !== 22) {
              rowModel.ht = 22; // 行高
              rowModel.customHeight = true; // 自定义行高标志
            }
          }
        });

        // 统计总行数
        totalRows += sheetData.rows.length;
      }

      // 尝试最后的解决方案 - 强制重置所有样式
      newWorkbook.eachSheet(worksheet => {
        console.log(
          `工作表 ${worksheet.name}: ${worksheet.rowCount} 行, ${worksheet.columnCount} 列`
        );

        // 清除所有行的默认样式
        worksheet.eachRow((row, rowNumber) => {
          // 重置行高
          row.height = 22;

          // 清除所有单元格样式
          row.eachCell(cell => {
            if (cell.value !== null && cell.value !== undefined) {
              // 强制设置字体，确保覆盖任何默认设置
              cell.font = {
                name: "宋体",
                size: 10,
                bold: rowNumber === 1,
                italic: false,
                underline: false,
                color: { argb: "FF000000" }
              };

              // 设置对齐
              cell.alignment = {
                vertical: "middle",
                horizontal: rowNumber === 1 ? "center" : "left"
              };

              // 添加细边框
              cell.border = {
                top: { style: "thin", color: { argb: "FF000000" } },
                left: { style: "thin", color: { argb: "FF000000" } },
                bottom: { style: "thin", color: { argb: "FF000000" } },
                right: { style: "thin", color: { argb: "FF000000" } }
              };

              // 调试输出
              if (rowNumber <= 3) {
                console.log(
                  `行${rowNumber}列${cell.col}: 字体大小=${cell.font.size}, 值=${cell.value}`
                );
              }
            }
          });
        });
      });

      console.log("样式设置完成，准备添加合计行...");

      // 为每个工作表添加合计行
      newWorkbook.eachSheet(worksheet => {
        console.log(`为工作表 ${worksheet.name} 添加合计行`);

        // 获取表头，找出金额列
        const headers: any[] = [];
        const headerRow = worksheet.getRow(1);
        headerRow.eachCell((cell: any) => {
          headers.push(cell.value);
        });

        // 找出总金额列 - 精确匹配
        let totalAmountColumnIndex = -1;
        headers.forEach((header, index) => {
          if (header && header.toString() === "总金额") {
            totalAmountColumnIndex = index + 1; // Excel列索引从1开始
          }
        });

        console.log(
          `工作表 ${worksheet.name} 找到总金额列:`,
          totalAmountColumnIndex !== -1
            ? headers[totalAmountColumnIndex - 1]
            : "未找到"
        );

        if (totalAmountColumnIndex !== -1 && worksheet.rowCount > 1) {
          // 添加合计行
          const totalRow = worksheet.addRow([]);
          const totalRowNumber = worksheet.rowCount;

          // 设置行高
          worksheet.getRow(totalRowNumber).height = 22;

          // 第一列显示"合计"
          const totalFirstCell = totalRow.getCell(1);
          totalFirstCell.value = "合计";
          totalFirstCell.font = { size: 10, bold: true, name: "宋体" };
          totalFirstCell.alignment = {
            vertical: "middle",
            horizontal: "center"
          };
          totalFirstCell.border = {
            top: { style: "thin", color: { argb: "FF000000" } },
            left: { style: "thin", color: { argb: "FF000000" } },
            bottom: { style: "thin", color: { argb: "FF000000" } },
            right: { style: "thin", color: { argb: "FF000000" } }
          };

          // 为总金额列添加求和公式
          const sumCell = totalRow.getCell(totalAmountColumnIndex);

          // 创建求和公式：从第2行到倒数第二行的总金额列
          const startRow = 2;
          const endRow = totalRowNumber - 1;
          const columnLetter = String.fromCharCode(64 + totalAmountColumnIndex); // 列号转字母

          sumCell.value = {
            formula: `SUM(${columnLetter}${startRow}:${columnLetter}${endRow})`,
            result: 0
          };

          sumCell.font = { size: 10, bold: true, name: "宋体" };
          sumCell.alignment = { vertical: "middle", horizontal: "right" };
          sumCell.border = {
            top: { style: "thin", color: { argb: "FF000000" } },
            left: { style: "thin", color: { argb: "FF000000" } },
            bottom: { style: "thin", color: { argb: "FF000000" } },
            right: { style: "thin", color: { argb: "FF000000" } }
          };
          sumCell.numFmt = "#,##0.00";

          // 为其他列设置边框
          for (let col = 2; col <= worksheet.columnCount; col++) {
            if (col !== totalAmountColumnIndex) {
              const cell = totalRow.getCell(col);
              cell.font = { size: 10, name: "宋体" };
              cell.border = {
                top: { style: "thin", color: { argb: "FF000000" } },
                left: { style: "thin", color: { argb: "FF000000" } },
                bottom: { style: "thin", color: { argb: "FF000000" } },
                right: { style: "thin", color: { argb: "FF000000" } }
              };
            }
          }

          console.log(
            `工作表 ${worksheet.name} 合计行添加完成，第 ${totalRowNumber} 行`
          );
        } else {
          console.log(
            `工作表 ${worksheet.name} 没有找到金额列或数据行为空，跳过合计行`
          );
        }
      });

      // 在合计行添加完成后，填充账单概览数据
      if (overviewSheet) {
        fillOverviewData(overviewSheet, accountantInfo.data, newWorkbook);
      }

      // 生成文件名
      const fileName = generateFileName(accountantName);
      if (!fileName) {
        continue; // 跳过"未找到对账人"的文件
      }

      // 生成文件并添加到ZIP
      const excelBuffer = await newWorkbook.xlsx.writeBuffer();
      zip.file(fileName, excelBuffer);

      // 记录生成的文件信息
      const displayName =
        accountantName === "未找到对账人" ? "未匹配部门的账单" : accountantName;
      generatedFiles.value.push({
        fileName,
        departmentName: displayName,
        rowCount: totalRows,
        contactName: displayName,
        contactEmail: email || "未配置"
      });

      totalFiles++;
      console.log(
        `已生成对账人文件: ${fileName}, 包含 ${accountantInfo.data.length} 个工作表, 共 ${totalRows} 行数据`
      );
    }

    // 生成ZIP文件
    const zipContent = await zip.generateAsync({ type: "blob" });
    const zipFileName = `艾比森账单按对账人拆分_${new Date().toISOString().slice(0, 10)}.zip`;

    saveAs(zipContent, zipFileName);

    console.log(`艾比森账单按对账人拆分完成: ${zipFileName}`);
    console.log(`生成的对账人文件数: ${totalFiles}`);

    ElMessage.success(
      `成功生成艾比森账单拆分ZIP包！包含 ${totalFiles} 个对账人文件`
    );
  } catch (error) {
    console.error("生成艾比森账单拆分文件失败:", error);
    ElMessage.error("生成拆分文件失败");
  } finally {
    generating.value = false;
  }
};

// 获取日期范围（上个月25号到当月24号）
const getDateRange = () => {
  const now = new Date();
  const currentMonth = now.getMonth() + 1;
  const currentYear = now.getFullYear();

  // 上个月的25号
  const lastMonthDate = new Date(currentYear, currentMonth - 2, 25);
  const lastMonthStr = `${lastMonthDate.getMonth() + 1}.${String(lastMonthDate.getDate()).padStart(2, "0")}`;

  // 当月的24号
  const currentMonthDate = new Date(currentYear, currentMonth - 1, 24);
  const currentMonthStr = `${currentMonthDate.getMonth() + 1}.${String(currentMonthDate.getDate()).padStart(2, "0")}`;

  return `${lastMonthStr}-${currentMonthStr}`;
};

// 生成文件名
const generateFileName = (accountantName: string) => {
  // 如果是"未找到对账人"，返回null（不生成文件）
  if (accountantName === "未找到对账人") {
    return null;
  }

  // 从processedWorkbook中查找该对账人对应的实际数据行
  if (processedWorkbook.value && processedWorkbook.value.worksheets) {
    const targetSheets = [
      "国内机票",
      "国际机票",
      "国内酒店",
      "国际酒店",
      "通用产品"
    ];

    for (const sheetName of targetSheets) {
      const worksheet = processedWorkbook.value.getWorksheet(sheetName);
      if (worksheet && worksheet.rowCount > 1) {
        // 获取表头
        const headers: any[] = [];
        const headerRow = worksheet.getRow(1);
        headerRow.eachCell({ includeEmpty: true }, (cell: any) => {
          headers.push(cell.value);
        });

        // 找到对账人和费用归属（全路径）列的索引
        const accountantIndex = headers.findIndex(
          (h: any) => h && h.toString().includes("对账人")
        );
        const costFullIndex = headers.findIndex(
          (h: any) => h && h.toString().includes("费用归属（全路径）")
        );

        if (accountantIndex !== -1 && costFullIndex !== -1) {
          // 遍历数据行，找到匹配该对账人的行
          for (let row = 2; row <= worksheet.rowCount; row++) {
            const dataRow = worksheet.getRow(row);
            const rowAccountant = dataRow.getCell(accountantIndex + 1).value;

            if (
              rowAccountant &&
              rowAccountant.toString().trim() === accountantName
            ) {
              const costValue = dataRow.getCell(costFullIndex + 1).value;
              if (costValue && typeof costValue === "string") {
                // 提取部门名称
                if (costValue.includes("-")) {
                  const parts = costValue.split("-");
                  let deptName = parts[0].trim();

                  // 特殊处理：如果是"服务体系"或"国际服务运营部"，取最后一个值
                  if (
                    parts.length > 1 &&
                    (deptName === "服务体系" || deptName === "国际服务运营部")
                  ) {
                    deptName = parts[parts.length - 1].trim();
                  }

                  const dateRange = getDateRange();
                  return `特航账单-${deptName}${dateRange}.xlsx`;
                } else {
                  // 如果没有"-"分隔符，直接使用原数据
                  const deptName = costValue.trim();
                  const dateRange = getDateRange();
                  return `特航账单-${deptName}${dateRange}.xlsx`;
                }
              }
            }
          }
        }
      }
    }
  }

  // 如果没有找到费用归属数据，使用对账人姓名作为部门名
  const dateRange = getDateRange();
  return `特航账单-${accountantName}${dateRange}.xlsx`;
};

// 处理对账人信息并生成拆分结果（不下载）
const processAccountantInfoAndGenerateResults = async () => {
  if (!originalWorkbook.value) {
    ElMessage.error("请先上传Excel文件");
    return;
  }

  try {
    console.log("=== 开始处理对账人信息并生成拆分结果 ===");

    // 创建工作簿的副本进行修改
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(await originalWorkbook.value.xlsx.writeBuffer());

    const targetSheets = [
      "国内机票",
      "国际机票",
      "国内酒店",
      "国际酒店",
      "通用产品"
    ];
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

      // 查找"费用归属（全路径）"列的索引 - 增加模糊匹配
      let costBelongFullIndex = headers.findIndex(
        (h: any) => h && h.toString() === "费用归属（全路径）"
      );

      // 如果精确匹配没找到，尝试模糊匹配
      if (costBelongFullIndex === -1) {
        costBelongFullIndex = headers.findIndex(
          (h: any) => h && h.toString().includes("费用归属")
        );
        if (costBelongFullIndex !== -1) {
          console.log(
            `工作表 ${sheetName} 使用模糊匹配找到费用归属列: "${headers[costBelongFullIndex]}"`
          );
        }
      }

      if (costBelongFullIndex === -1) {
        console.log(`工作表 ${sheetName} 未找到"费用归属（全路径）列，跳过`);
        continue;
      }

      // 查找"费用归属"列的索引（作为备用查询）
      let costBelongIndex = headers.findIndex(
        (h: any) => h && h.toString() === "费用归属"
      );

      // 查找"预订人"或"预定人"列的索引
      const bookingPersonIndex = headers.findIndex(
        (h: any) =>
          h &&
          (h.toString().includes("预订人") || h.toString().includes("预定人"))
      );

      // 查找"对账人"列是否存在，如果不存在则添加
      let accountantIndex = headers.findIndex(
        (h: any) => h && h.toString().includes("对账人")
      );

      if (accountantIndex === -1) {
        if (bookingPersonIndex === -1) {
          // 如果没找到预订人列，就在最后一列添加
          accountantIndex = headers.length;
        } else {
          // 在预订人/预定人列前面插入对账人列
          accountantIndex = bookingPersonIndex;
        }

        // 在指定位置插入列
        for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber++) {
          const row = worksheet.getRow(rowNumber);
          // 移动该列及其之后的所有列向右一格
          for (
            let col = worksheet.columnCount;
            col >= accountantIndex + 1;
            col--
          ) {
            const sourceCell = row.getCell(col);
            const targetCell = row.getCell(col + 1);
            targetCell.value = sourceCell.value;

            // 区分数据行和表头行的样式处理
            if (rowNumber > 1) {
              // 数据行：复制完整样式但排除填充样式
              if (sourceCell.style) {
                const styleCopy = { ...sourceCell.style };
                // 清除填充样式，避免颜色错乱
                if (styleCopy.fill) {
                  delete styleCopy.fill;
                }
                targetCell.style = styleCopy;
              }
            }
          }
          // 清空原来位置的单元格并设置默认样式
          const newCell = row.getCell(accountantIndex + 1);
          newCell.value = null;
          newCell.font = { bold: false };
          newCell.fill = {
            type: "pattern",
            pattern: "none"
          };
        }

        // 单独设置对账人列表头的值和样式
        const headerCell = headerRow.getCell(accountantIndex + 1);
        headerCell.value = "对账人";
        headerCell.font = { bold: true };
        headerCell.fill = {
          type: "pattern",
          pattern: "none"
        };
      }

      // 由于插入了对账人列，需要调整费用归属列的索引
      let adjustedCostBelongFullIndex = costBelongFullIndex;
      let adjustedCostBelongIndex = costBelongIndex;
      if (accountantIndex !== -1 && accountantIndex <= costBelongFullIndex) {
        adjustedCostBelongFullIndex = costBelongFullIndex + 1;
      }
      if (
        costBelongIndex !== -1 &&
        accountantIndex !== -1 &&
        accountantIndex <= costBelongIndex
      ) {
        adjustedCostBelongIndex = costBelongIndex + 1;
      }

      // 重新获取表头，因为可能已经插入了对账人列
      const updatedHeaders: any[] = [];
      headerRow.eachCell({ includeEmpty: true }, (cell: any) => {
        updatedHeaders.push(cell.value);
      });

      // 重新获取对账人列索引
      const updatedAccountantIndex = updatedHeaders.findIndex(
        (h: any) => h && h.toString() === "对账人"
      );

      // 处理数据行
      const rowCount = worksheet.rowCount;
      let sheetProcessed = 0;
      let sheetYellowCells = 0;

      for (let rowNumber = 2; rowNumber <= rowCount; rowNumber++) {
        const row = worksheet.getRow(rowNumber);
        const costBelongFullValue = row.getCell(
          adjustedCostBelongFullIndex + 1
        ).value;

        // 获取费用归属列的值（如果存在该列）
        let costBelongValue = null;
        if (adjustedCostBelongIndex !== -1) {
          costBelongValue = row.getCell(adjustedCostBelongIndex + 1).value;
        }

        // 检查是否为空行（所有列都是空的）
        let isEmptyRow = true;
        row.eachCell({ includeEmpty: false }, (cell: any) => {
          if (
            cell.value !== null &&
            cell.value !== undefined &&
            cell.value !== ""
          ) {
            isEmptyRow = false;
          }
        });

        // 如果是空行，跳过处理
        if (isEmptyRow) {
          continue;
        }

        const accountantCell = row.getCell(updatedAccountantIndex + 1);

        // 使用新的备用查询逻辑
        const { contactInfo, source, departmentName } =
          getContactInfoWithFallback(
            costBelongFullValue,
            costBelongValue,
            sheetName,
            rowNumber
          );

        if (contactInfo && contactInfo.accountant) {
          // 填写对账人姓名
          accountantCell.value = contactInfo.accountant;
          accountantCell.font = { bold: false };
          // 清除任何现有的填充样式
          accountantCell.fill = {
            type: "pattern",
            pattern: "none"
          };
          console.log(
            `工作表 ${sheetName} 第 ${rowNumber} 行: 对账人="${contactInfo.accountant}" (来源: ${source})`
          );
          sheetProcessed++;
        } else {
          // 如果两种查询方式都找不到对账人，设置"未找到对账人"并设置红色
          accountantCell.value = "未找到对账人";

          // 设置背景为红色
          accountantCell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFF0000" }
          };

          // 强制重新应用样式
          accountantCell.style = {
            ...accountantCell.style,
            font: { bold: true },
            fill: accountantCell.fill
          };

          console.log(
            `工作表 ${sheetName} 第 ${rowNumber} 行: 两种查询方式都未找到对账人，设置红色标识`
          );
          sheetYellowCells++;
          totalYellowCells++;
        }
      }

      console.log(
        `工作表 ${sheetName} 处理完成: 处理 ${sheetProcessed} 行，${sheetYellowCells} 行未找到对账人`
      );
      totalProcessed += sheetProcessed;
    }

    // 保存处理后的工作簿
    processedWorkbook.value = workbook;

    console.log(
      `对账人信息处理完成: 总共处理 ${totalProcessed} 行对账人信息，${totalYellowCells} 行未找到对账人`
    );

    // 生成拆分结果预览
    await generateSplitResults();

    ElMessage.success(
      `处理完成！总共处理 ${totalProcessed} 行对账人信息，${totalYellowCells} 行未找到对账人`
    );
  } catch (error) {
    console.error("处理对账人信息失败:", error);
    ElMessage.error("处理对账人信息失败");
  }
};

// 生成拆分结果预览（不下载）
const generateSplitResults = async () => {
  if (!processedWorkbook.value) {
    ElMessage.error("请先处理对账人信息");
    return;
  }

  try {
    console.log("=== 生成拆分结果预览 ===");

    // 清空之前的结果
    generatedFiles.value = [];

    // 使用已处理的工作簿
    const sourceWorkbook = processedWorkbook.value;

    // 收集所有对账人的数据
    const accountantDataMap = new Map<
      string,
      {
        accountant: string;
        email: string;
        data: Array<{
          sheetName: string;
          rows: any[];
        }>;
        firstRowData?: any[];
      }
    >();

    const targetSheets = [
      "国内机票",
      "国际机票",
      "国内酒店",
      "国际酒店",
      "通用产品"
    ];

    // 遍历每个工作表收集数据
    for (const sheetName of targetSheets) {
      const worksheet = sourceWorkbook.getWorksheet(sheetName);
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

      // 找到对账人列索引
      const accountantIndex = headers.findIndex(
        (h: any) => h && h.toString().includes("对账人")
      );

      if (accountantIndex === -1) {
        console.log(`工作表 ${sheetName} 未找到对账人列，跳过`);
        continue;
      }

      // 收集数据行，按对账人分组
      const rowCount = worksheet.rowCount;
      for (let rowNumber = 2; rowNumber <= rowCount; rowNumber++) {
        const row = worksheet.getRow(rowNumber);
        const accountantCell = row.getCell(accountantIndex + 1);
        const accountantName = accountantCell.value?.toString().trim();

        if (!accountantName || accountantName === "未找到对账人") {
          continue; // 跳过空行和"未找到对账人"的数据
        }

        // 如果该对账人还未在Map中，先初始化
        if (!accountantDataMap.has(accountantName)) {
          accountantDataMap.set(accountantName, {
            accountant: accountantName,
            email: "", // 稍后填充
            data: []
          });
        }

        // 获取整行数据
        const rowData: any[] = [];
        row.eachCell({ includeEmpty: true }, (cell: any) => {
          rowData.push(cell.value);
        });

        // 添加到对应对账人的数据中
        const accountantData = accountantDataMap.get(accountantName)!;
        if (!accountantData.data.find(d => d.sheetName === sheetName)) {
          accountantData.data.push({
            sheetName,
            rows: []
          });
        }
        const sheetData = accountantData.data.find(
          d => d.sheetName === sheetName
        )!;
        sheetData.rows.push(rowData);

        // 保存第一条数据（用于文件名生成）
        if (!accountantData.firstRowData && rowData.length > 0) {
          accountantData.firstRowData = rowData;
        }
      }
    }

    console.log(`收集到 ${accountantDataMap.size} 个对账人的数据`);

    // 输出详细的调试信息
    console.log("=== 对账人数据详情 ===");
    for (const [accountantName, accountantInfo] of accountantDataMap) {
      console.log(`对账人: ${accountantName}`);
      for (const sheetData of accountantInfo.data) {
        console.log(
          `  - 工作表: ${sheetData.sheetName}, 行数: ${sheetData.rows.length}`
        );
      }
    }

    // 为每个对账人生成文件信息
    for (const [accountantName, accountantInfo] of accountantDataMap) {
      // 查找对账人邮箱
      let email = "";
      for (const contact of Object.values(ABSEN_DEPARTMENT_TO_CONTACT_MAP)) {
        if (contact.accountant === accountantName) {
          email = contact.email;
          break;
        }
      }

      // 统计总行数
      let totalRows = 0;
      for (const sheetData of accountantInfo.data) {
        totalRows += sheetData.rows.length;
      }

      // 生成文件名
      const fileName = generateFileName(accountantName);
      if (!fileName) {
        continue; // 跳过"未找到对账人"的文件
      }

      // 记录生成的文件信息
      const displayName =
        accountantName === "未找到对账人" ? "未匹配部门的账单" : accountantName;
      generatedFiles.value.push({
        fileName,
        departmentName: displayName,
        rowCount: totalRows,
        contactName: displayName,
        contactEmail: email || "未配置"
      });

      console.log(
        `预览文件: ${fileName}, 包含 ${accountantInfo.data.length} 个工作表, 共 ${totalRows} 行数据`
      );
    }

    console.log(
      `拆分结果预览生成完成，共 ${generatedFiles.value.length} 个文件`
    );
  } catch (error) {
    console.error("生成拆分结果预览失败:", error);
    ElMessage.error("生成拆分结果预览失败");
  }
};

// 下载调整后的数据
const downloadAdjustedData = async () => {
  if (!processedWorkbook.value || generatedFiles.value.length === 0) {
    ElMessage.error("请先上传文件并生成拆分结果");
    return;
  }

  try {
    console.log("=== 开始下载调整后的数据 ===");

    // 创建ZIP文件
    const zip = new JSZip();
    let totalFiles = 0;

    // 加载 absen账单概览.xlsx 文件
    const overviewWorkbook = await loadOverviewWorkbook();

    // 使用已处理的工作簿
    const sourceWorkbook = processedWorkbook.value;

    // 收集所有对账人的数据（重新收集以确保使用最新数据）
    const accountantDataMap = new Map<
      string,
      {
        accountant: string;
        email: string;
        data: Array<{
          sheetName: string;
          rows: any[];
        }>;
        firstRowData?: any[];
      }
    >();

    const targetSheets = [
      "国内机票",
      "国际机票",
      "国内酒店",
      "国际酒店",
      "通用产品"
    ];

    // 遍历每个工作表收集数据
    for (const sheetName of targetSheets) {
      const worksheet = sourceWorkbook.getWorksheet(sheetName);
      if (!worksheet) continue;

      // 获取表头
      const headers: any[] = [];
      const headerRow = worksheet.getRow(1);
      headerRow.eachCell({ includeEmpty: true }, (cell: any) => {
        headers.push(cell.value);
      });

      // 找到对账人列索引
      const accountantIndex = headers.findIndex(
        (h: any) => h && h.toString().includes("对账人")
      );

      if (accountantIndex === -1) continue;

      // 收集数据行，按对账人分组
      const rowCount = worksheet.rowCount;
      for (let rowNumber = 2; rowNumber <= rowCount; rowNumber++) {
        const row = worksheet.getRow(rowNumber);
        const accountantCell = row.getCell(accountantIndex + 1);
        const accountantName = accountantCell.value?.toString().trim();

        if (!accountantName || accountantName === "未找到对账人") {
          continue; // 跳过空行和"未找到对账人"的数据
        }

        // 如果该对账人还未在Map中，先初始化
        if (!accountantDataMap.has(accountantName)) {
          accountantDataMap.set(accountantName, {
            accountant: accountantName,
            email: "",
            data: []
          });
        }

        // 获取整行数据
        const rowData: any[] = [];
        row.eachCell({ includeEmpty: true }, (cell: any) => {
          rowData.push(cell.value);
        });

        // 添加到对应对账人的数据中
        const accountantData = accountantDataMap.get(accountantName)!;
        if (!accountantData.data.find(d => d.sheetName === sheetName)) {
          accountantData.data.push({
            sheetName,
            rows: []
          });
        }
        const sheetData = accountantData.data.find(
          d => d.sheetName === sheetName
        )!;
        sheetData.rows.push(rowData);

        // 保存第一条数据（用于文件名生成）
        if (!accountantData.firstRowData && rowData.length > 0) {
          accountantData.firstRowData = rowData;
        }
      }
    }

    // 为每个对账人生成Excel文件
    for (const [accountantName, accountantInfo] of accountantDataMap) {
      // 查找生成的文件信息中的文件名
      const generatedFileInfo = generatedFiles.value.find(
        f => f.contactName === accountantName
      );

      // 如果没有预生成的文件信息，生成文件名
      let fileName = generatedFileInfo?.fileName;
      if (!fileName) {
        fileName = generateFileName(accountantName);
        if (!fileName) {
          continue; // 跳过"未找到对账人"的文件
        }
      }

      // 创建新的工作簿
      const newWorkbook = new ExcelJS.Workbook();

      // 首先添加账单概览工作表（如果存在）
      let overviewSheet = null;
      if (overviewWorkbook && overviewWorkbook.worksheets.length > 0) {
        for (const sourceOverviewSheet of overviewWorkbook.worksheets) {
          // 使用增强版复制函数复制概览工作表
          overviewSheet = newWorkbook.addWorksheet(sourceOverviewSheet.name);
          copyWorksheetWithFormat(sourceOverviewSheet, overviewSheet);
        }
      }

      // 为每个工作表创建数据表
      for (const sheetData of accountantInfo.data) {
        const newWorksheet = newWorkbook.addWorksheet(sheetData.sheetName);

        // 复制原工作表的表头
        const sourceWorksheet = sourceWorkbook.getWorksheet(
          sheetData.sheetName
        );
        const sourceHeaderRow = sourceWorksheet!.getRow(1);

        // 创建表头行并设置样式
        const newHeaderRow = newWorksheet.getRow(1);
        newHeaderRow.height = 22;

        // 复制表头数据并设置样式
        const headerData: any[] = [];
        sourceHeaderRow.eachCell({ includeEmpty: true }, (cell: any) => {
          headerData.push(cell.value);
        });

        // 设置表头单元格并应用样式
        for (let col = 0; col < headerData.length; col++) {
          const cell = newHeaderRow.getCell(col + 1);
          cell.value = headerData[col];
          cell.font = { size: 10, bold: true, name: "宋体" };
          cell.alignment = { vertical: "middle", horizontal: "center" };
          cell.border = {
            top: { style: "thin", color: { argb: "FF000000" } },
            left: { style: "thin", color: { argb: "FF000000" } },
            bottom: { style: "thin", color: { argb: "FF000000" } },
            right: { style: "thin", color: { argb: "FF000000" } }
          };
        }

        // 复制数据行并应用样式
        for (let i = 0; i < sheetData.rows.length; i++) {
          const rowData = sheetData.rows[i];
          const newRow = newWorksheet.getRow(i + 2);
          newRow.height = 22;

          for (let j = 0; j < rowData.length; j++) {
            const cell = newRow.getCell(j + 1);
            cell.value = rowData[j];
            cell.font = { size: 10, name: "宋体" };
            cell.alignment = { vertical: "middle" };
            cell.border = {
              top: { style: "thin", color: { argb: "FF000000" } },
              left: { style: "thin", color: { argb: "FF000000" } },
              bottom: { style: "thin", color: { argb: "FF000000" } },
              right: { style: "thin", color: { argb: "FF000000" } }
            };
          }
        }

        // 设置列宽
        newWorksheet.columns.forEach(column => {
          column.width = 15;
        });
      }

      // 为每个工作表添加合计行
      newWorkbook.eachSheet(worksheet => {
        // 获取表头，找出总金额列
        const headers: any[] = [];
        const headerRow = worksheet.getRow(1);
        headerRow.eachCell((cell: any) => {
          headers.push(cell.value);
        });

        // 找出总金额列 - 精确匹配
        let totalAmountColumnIndex = -1;
        headers.forEach((header, index) => {
          if (header && header.toString() === "总金额") {
            totalAmountColumnIndex = index + 1;
          }
        });

        if (totalAmountColumnIndex !== -1 && worksheet.rowCount > 1) {
          // 添加合计行
          const totalRow = worksheet.addRow([]);
          const totalRowNumber = worksheet.rowCount;

          // 设置行高
          worksheet.getRow(totalRowNumber).height = 22;

          // 第一列显示"合计"
          const totalFirstCell = totalRow.getCell(1);
          totalFirstCell.value = "合计";
          totalFirstCell.font = { size: 10, bold: true, name: "宋体" };
          totalFirstCell.alignment = {
            vertical: "middle",
            horizontal: "center"
          };
          totalFirstCell.border = {
            top: { style: "thin", color: { argb: "FF000000" } },
            left: { style: "thin", color: { argb: "FF000000" } },
            bottom: { style: "thin", color: { argb: "FF000000" } },
            right: { style: "thin", color: { argb: "FF000000" } }
          };

          // 为总金额列添加求和公式
          const sumCell = totalRow.getCell(totalAmountColumnIndex);
          const startRow = 2;
          const endRow = totalRowNumber - 1;
          const columnLetter = String.fromCharCode(64 + totalAmountColumnIndex);

          sumCell.value = {
            formula: `SUM(${columnLetter}${startRow}:${columnLetter}${endRow})`,
            result: 0
          };

          sumCell.font = { size: 10, bold: true, name: "宋体" };
          sumCell.alignment = { vertical: "middle", horizontal: "right" };
          sumCell.border = {
            top: { style: "thin", color: { argb: "FF000000" } },
            left: { style: "thin", color: { argb: "FF000000" } },
            bottom: { style: "thin", color: { argb: "FF000000" } },
            right: { style: "thin", color: { argb: "FF000000" } }
          };
          sumCell.numFmt = "#,##0.00";

          // 为其他列设置边框
          for (let col = 2; col <= worksheet.columnCount; col++) {
            if (col !== totalAmountColumnIndex) {
              const cell = totalRow.getCell(col);
              cell.font = { size: 10, name: "宋体" };
              cell.border = {
                top: { style: "thin", color: { argb: "FF000000" } },
                left: { style: "thin", color: { argb: "FF000000" } },
                bottom: { style: "thin", color: { argb: "FF000000" } },
                right: { style: "thin", color: { argb: "FF000000" } }
              };
            }
          }
        }
      });

      // 在合计行添加完成后，填充账单概览数据
      if (overviewSheet) {
        fillOverviewData(overviewSheet, accountantInfo.data, newWorkbook);
      }

      // 生成文件并添加到ZIP
      const excelBuffer = await newWorkbook.xlsx.writeBuffer();
      zip.file(fileName, excelBuffer);

      totalFiles++;
      console.log(`已生成对账人文件: ${fileName}`);
    }

    // 生成ZIP文件
    const zipContent = await zip.generateAsync({ type: "blob" });
    const zipFileName = `艾比森账单按对账人拆分_${new Date().toISOString().slice(0, 10)}.zip`;

    saveAs(zipContent, zipFileName);

    console.log(`艾比森账单按对账人拆分完成: ${zipFileName}`);
    console.log(`生成的对账人文件数: ${totalFiles}`);

    ElMessage.success(
      `成功下载艾比森账单拆分ZIP包！包含 ${totalFiles} 个对账人文件`
    );
  } catch (error) {
    console.error("下载拆分文件失败:", error);
    ElMessage.error("下载拆分文件失败");
  }
};

// 更新文件名
// 加载账单概览工作簿
const loadOverviewWorkbook = async () => {
  try {
    const response = await fetch(overviewFileUrl);
    if (response.ok) {
      const buffer = await response.arrayBuffer();
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);
      console.log("成功加载 absen账单概览.xlsx 文件");
      return workbook;
    }
  } catch (error) {
    console.warn("未能加载 absen账单概览.xlsx 文件:", error);
  }
  return null;
};

// 增强的工作表复制函数，保留所有格式
const copyWorksheetWithFormat = (
  sourceWorksheet: any,
  targetWorksheet: any
) => {
  console.log(`开始复制工作表: ${sourceWorksheet.name}`);

  // 复制所有行数据和样式
  sourceWorksheet.eachRow((sourceRow: any, rowNumber: number) => {
    const targetRow = targetWorksheet.getRow(rowNumber);

    // 复制行高
    if (sourceRow.height) {
      targetRow.height = sourceRow.height;
    }

    // 复制每个单元格
    sourceRow.eachCell((sourceCell: any, colNumber: number) => {
      const targetCell = targetRow.getCell(colNumber);

      // 复制值
      targetCell.value = sourceCell.value;

      // 复制完整的样式对象
      if (sourceCell.style) {
        targetCell.style = JSON.parse(JSON.stringify(sourceCell.style));
      }

      // 复制数字格式
      if (sourceCell.numFmt) {
        targetCell.numFmt = sourceCell.numFmt;
      }
    });
  });

  // 复制列宽
  sourceWorksheet.columns.forEach((column: any, index: number) => {
    if (column.width) {
      targetWorksheet.getColumn(index + 1).width = column.width;
    }
  });

  // 复制合并单元格
  if (sourceWorksheet.model && sourceWorksheet.model.merges) {
    sourceWorksheet.model.merges.forEach((merge: any) => {
      try {
        targetWorksheet.mergeCells(merge);
        console.log(`复制合并单元格: ${merge}`);
      } catch (e) {
        console.warn(`忽略合并单元格错误: ${merge}`, e);
      }
    });
  }

  // 复制工作表属性
  if (sourceWorksheet.properties) {
    targetWorksheet.properties = { ...sourceWorksheet.properties };
  }

  // 复制工作表视图
  if (sourceWorksheet.views) {
    targetWorksheet.views = sourceWorksheet.views.map((view: any) => ({
      ...view
    }));
  }

  console.log(`工作表 ${sourceWorksheet.name} 复制完成`);
};

// 辅助函数：查找列索引
const findColumnIndex = (worksheet: any, columnName: string): number => {
  if (!worksheet) return -1;

  const headers: any[] = [];
  const headerRow = worksheet.getRow(1);
  headerRow.eachCell((cell: any) => {
    headers.push(cell.value);
  });

  const index = headers.findIndex(
    (h: any) => h && h.toString().trim() === columnName.trim()
  );

  return index !== -1 ? index + 1 : -1; // Excel列索引从1开始
};

// 填充账单概览数据
const fillOverviewData = (
  overviewWorksheet: any,
  accountantData: any,
  newWorkbook: any
) => {
  console.log("开始填充账单概览数据");

  // 首先检查概览表的结构
  console.log("=== 检查概览表结构 ===");
  overviewWorksheet.eachRow((row: any, rowNumber: number) => {
    const rowData = [];
    // 查看所有列
    for (let i = 1; i <= 20; i++) { // 看前20列
      const cell = row.getCell(i);
      const cellValue = cell.value;
      // 只记录有值的单元格
      if (cellValue !== null && cellValue !== undefined && cellValue !== '') {
        rowData.push(`第${i}列: ${cellValue}`);
      }
    }
    if (rowData.length > 0) {
      console.log(`概览表第${rowNumber}行:`, rowData);
    }
  });

  // 计算各列的总和
  let domesticFlightTotal = 0;
  let systemFeeTotal = 0;
  let insuranceTotal = 0;
  let changeFeeTotal = 0;
  let refundFeeTotal = 0;

  // 从拆分后的国内机票工作表中获取数据
  const domesticFlightSheet = newWorkbook.getWorksheet("国内机票");
  let flightFee = 0; // 声明在外部，确保后续能访问

  if (!domesticFlightSheet) {
    console.warn("未找到国内机票工作表，跳过国内机票数据处理");
  } else {

  console.log(
    `国内机票工作表信息: 行数=${domesticFlightSheet.rowCount}, 列数=${domesticFlightSheet.columnCount}`
  );

  // 获取表头以找到各列的索引
  const headers: any[] = [];
  const headerRow = domesticFlightSheet.getRow(1);
  headerRow.eachCell((cell: any) => {
    headers.push(cell.value);
  });

  console.log("国内机票表头:", headers);

  // 查找各列的索引
  const totalAmountIndex = findColumnIndex(domesticFlightSheet, "总金额");
  const systemFeeIndex = findColumnIndex(domesticFlightSheet, "系统使用费");
  const insuranceIndex = findColumnIndex(domesticFlightSheet, "保险费");
  const changeFeeIndex =
    findColumnIndex(domesticFlightSheet, "改签费") ||
    findColumnIndex(domesticFlightSheet, "改签手续费");
  const refundFeeIndex =
    findColumnIndex(domesticFlightSheet, "退票费") ||
    findColumnIndex(domesticFlightSheet, "退票手续费");

  console.log(
    `国内机票列索引: 总金额=${totalAmountIndex}, 系统使用费=${systemFeeIndex}, 保险费=${insuranceIndex}, 改签费=${changeFeeIndex}, 退票费=${refundFeeIndex}`
  );

  domesticFlightSheet.eachRow((row: any, rowNumber: number) => {
    if (rowNumber > 1) {
      // 跳过表头
      // 检查是否是合计行（通常第一列会显示"合计"）
      const firstCell = row.getCell(1);
      const isTotalRow = firstCell.value === "合计";

      // 记录前几行的数据
      if (rowNumber <= 3) {
        console.log(`第${rowNumber}行数据:`, {
          第一列: firstCell.value,
          总金额:
            totalAmountIndex !== -1
              ? row.getCell(totalAmountIndex).value
              : "未找到列",
          系统使用费:
            systemFeeIndex !== -1
              ? row.getCell(systemFeeIndex).value
              : "未找到列",
          保险费:
            insuranceIndex !== -1
              ? row.getCell(insuranceIndex).value
              : "未找到列",
          是合计行: isTotalRow
        });
      }

      if (!isTotalRow) {
        // 计算总金额
        if (totalAmountIndex !== -1) {
          const totalAmountCell = row.getCell(totalAmountIndex);
          if (
            totalAmountCell.value &&
            typeof totalAmountCell.value === "number"
          ) {
            domesticFlightTotal += totalAmountCell.value;
          }
        }

        // 计算系统使用费
        if (systemFeeIndex !== -1) {
          const systemFeeCell = row.getCell(systemFeeIndex);
          if (systemFeeCell.value && typeof systemFeeCell.value === "number") {
            systemFeeTotal += systemFeeCell.value;
          }
        }

        // 计算保险费
        if (insuranceIndex !== -1) {
          const insuranceCell = row.getCell(insuranceIndex);
          if (insuranceCell.value && typeof insuranceCell.value === "number") {
            insuranceTotal += insuranceCell.value;
          }
        }

        // 计算改签手续费
        if (changeFeeIndex !== -1) {
          const changeFeeCell = row.getCell(changeFeeIndex);
          if (changeFeeCell.value && typeof changeFeeCell.value === "number") {
            changeFeeTotal += changeFeeCell.value;
          }
        }

        // 计算退票手续费
        if (refundFeeIndex !== -1) {
          const refundFeeCell = row.getCell(refundFeeIndex);
          if (refundFeeCell.value && typeof refundFeeCell.value === "number") {
            refundFeeTotal += refundFeeCell.value;
          }
        }
      }
    }
  });

  console.log("计算结果:", {
    合计: domesticFlightTotal,
    系统使用费: systemFeeTotal,
    保险费: insuranceTotal,
    改签手续费: changeFeeTotal,
    退票手续费: refundFeeTotal
  });

  // 计算机票费
  flightFee =
    domesticFlightTotal -
    systemFeeTotal -
    insuranceTotal -
    changeFeeTotal -
    refundFeeTotal;
  }

  // 处理国内酒店费用
  let domesticHotelTotal = 0;
  let hotelSystemFeeTotal = 0;
  let hotelCancelFeeTotal = 0;

  // 从拆分后的国内酒店工作表中获取数据
  const domesticHotelSheet = newWorkbook.getWorksheet("国内酒店");
  if (domesticHotelSheet) {
    // 获取表头以找到各列的索引
    const hotelHeaders: any[] = [];
    const hotelHeaderRow = domesticHotelSheet.getRow(1);
    hotelHeaderRow.eachCell((cell: any) => {
      hotelHeaders.push(cell.value);
    });

    console.log("国内酒店表头:", hotelHeaders);

    // 查找各列的索引
    const hotelTotalAmountIndex = findColumnIndex(domesticHotelSheet, "总金额");
    const hotelSystemFeeIndex = findColumnIndex(
      domesticHotelSheet,
      "系统使用费"
    );
    const hotelCancelFeeIndex =
      findColumnIndex(domesticHotelSheet, "退订费") ||
      findColumnIndex(domesticHotelSheet, "退订手续费");

    console.log(
      `国内酒店列索引: 总金额=${hotelTotalAmountIndex}, 系统使用费=${hotelSystemFeeIndex}, 退订费=${hotelCancelFeeIndex}`
    );

    domesticHotelSheet.eachRow((row: any, rowNumber: number) => {
      if (rowNumber > 1) {
        // 跳过表头
        // 检查是否是合计行
        const firstCell = row.getCell(1);
        const isTotalRow = firstCell.value === "合计";

        // 记录前几行的数据
        if (rowNumber <= 3) {
          console.log(`国内酒店第${rowNumber}行数据:`, {
            第一列: firstCell.value,
            总金额:
              hotelTotalAmountIndex !== -1
                ? row.getCell(hotelTotalAmountIndex).value
                : "未找到列",
            系统使用费:
              hotelSystemFeeIndex !== -1
                ? row.getCell(hotelSystemFeeIndex).value
                : "未找到列",
            退订费:
              hotelCancelFeeIndex !== -1
                ? row.getCell(hotelCancelFeeIndex).value
                : "未找到列",
            是合计行: isTotalRow
          });
        }

        if (!isTotalRow) {
          // 计算总金额
          if (hotelTotalAmountIndex !== -1) {
            const totalAmountCell = row.getCell(hotelTotalAmountIndex);
            if (
              totalAmountCell.value &&
              typeof totalAmountCell.value === "number"
            ) {
              domesticHotelTotal += totalAmountCell.value;
            }
          }

          // 计算系统使用费
          if (hotelSystemFeeIndex !== -1) {
            const systemFeeCell = row.getCell(hotelSystemFeeIndex);
            if (
              systemFeeCell.value &&
              typeof systemFeeCell.value === "number"
            ) {
              hotelSystemFeeTotal += systemFeeCell.value;
            }
          }

          // 计算退订手续费
          if (hotelCancelFeeIndex !== -1) {
            const cancelFeeCell = row.getCell(hotelCancelFeeIndex);
            if (
              cancelFeeCell.value &&
              typeof cancelFeeCell.value === "number"
            ) {
              hotelCancelFeeTotal += cancelFeeCell.value;
            }
          }
        }
      }
    });

    console.log("国内酒店计算结果:", {
      合计: domesticHotelTotal,
      系统使用费: hotelSystemFeeTotal,
      退订手续费: hotelCancelFeeTotal
    });
  }

  // 计算酒店费
  const hotelFee =
    domesticHotelTotal - hotelSystemFeeTotal - hotelCancelFeeTotal;

  // 处理国际酒店费用
  let internationalHotelTotal = 0;
  let intlHotelSystemFeeTotal = 0;
  let intlHotelCancelFeeTotal = 0;

  // 从拆分后的国际酒店工作表中获取数据
  const internationalHotelSheet = newWorkbook.getWorksheet("国际酒店");
  if (internationalHotelSheet) {
    // 获取表头以找到各列的索引
    const intlHotelHeaders: any[] = [];
    const intlHotelHeaderRow = internationalHotelSheet.getRow(1);
    intlHotelHeaderRow.eachCell((cell: any) => {
      intlHotelHeaders.push(cell.value);
    });

    console.log("国际酒店表头:", intlHotelHeaders);

    // 查找各列的索引
    const intlHotelTotalAmountIndex = findColumnIndex(
      internationalHotelSheet,
      "总金额"
    );
    const intlHotelSystemFeeIndex = findColumnIndex(
      internationalHotelSheet,
      "系统使用费"
    );
    const intlHotelCancelFeeIndex =
      findColumnIndex(internationalHotelSheet, "退订手续费") ||
      findColumnIndex(internationalHotelSheet, "退订费");

    console.log(
      `国际酒店列索引: 总金额=${intlHotelTotalAmountIndex}, 系统使用费=${intlHotelSystemFeeIndex}, 退订手续费=${intlHotelCancelFeeIndex}`
    );

    internationalHotelSheet.eachRow((row: any, rowNumber: number) => {
      if (rowNumber > 1) {
        // 跳过表头
        // 检查是否是合计行
        const firstCell = row.getCell(1);
        const isTotalRow = firstCell.value === "合计";

        // 记录前几行的数据
        if (rowNumber <= 3) {
          console.log(`国际酒店第${rowNumber}行数据:`, {
            第一列: firstCell.value,
            总金额:
              intlHotelTotalAmountIndex !== -1
                ? row.getCell(intlHotelTotalAmountIndex).value
                : "未找到列",
            系统使用费:
              intlHotelSystemFeeIndex !== -1
                ? row.getCell(intlHotelSystemFeeIndex).value
                : "未找到列",
            退订手续费:
              intlHotelCancelFeeIndex !== -1
                ? row.getCell(intlHotelCancelFeeIndex).value
                : "未找到列",
            是合计行: isTotalRow
          });
        }

        if (!isTotalRow) {
          // 计算总金额
          if (intlHotelTotalAmountIndex !== -1) {
            const totalAmountCell = row.getCell(intlHotelTotalAmountIndex);
            if (
              totalAmountCell.value &&
              typeof totalAmountCell.value === "number"
            ) {
              internationalHotelTotal += totalAmountCell.value;
            }
          }

          // 计算系统使用费
          if (intlHotelSystemFeeIndex !== -1) {
            const systemFeeCell = row.getCell(intlHotelSystemFeeIndex);
            if (
              systemFeeCell.value &&
              typeof systemFeeCell.value === "number"
            ) {
              intlHotelSystemFeeTotal += systemFeeCell.value;
            }
          }

          // 计算退订手续费
          if (intlHotelCancelFeeIndex !== -1) {
            const cancelFeeCell = row.getCell(intlHotelCancelFeeIndex);
            if (
              cancelFeeCell.value &&
              typeof cancelFeeCell.value === "number"
            ) {
              intlHotelCancelFeeTotal += cancelFeeCell.value;
            }
          }
        }
      }
    });

    console.log("国际酒店计算结果:", {
      合计: internationalHotelTotal,
      系统使用费: intlHotelSystemFeeTotal,
      退订手续费: intlHotelCancelFeeTotal
    });
  }

  // 计算国际酒店费
  const intlHotelFee =
    internationalHotelTotal - intlHotelSystemFeeTotal - intlHotelCancelFeeTotal;

  // 处理国际机票费用
  let internationalFlightTotal = 0;
  let intlSystemFeeTotal = 0;
  let intlInsuranceTotal = 0;
  let intlChangeFeeTotal = 0;
  let intlRefundFeeTotal = 0;
  let intlFlightFee = 0; // 在外部声明 intlFlightFee

  // 从拆分后的国际机票工作表中获取数据
  const internationalFlightSheet = newWorkbook.getWorksheet("国际机票");
  if (internationalFlightSheet) {
    // 获取表头以找到各列的索引
    const intlHeaders: any[] = [];
    const intlHeaderRow = internationalFlightSheet.getRow(1);
    intlHeaderRow.eachCell((cell: any) => {
      intlHeaders.push(cell.value);
    });

    console.log("国际机票表头:", intlHeaders);

    // 查找各列的索引 - 按照用户要求精确匹配
    const intlTotalAmountIndex = findColumnIndex(
      internationalFlightSheet,
      "总金额"
    );
    const intlSystemFeeIndex = findColumnIndex(
      internationalFlightSheet,
      "系统使用费"
    );
    const intlInsuranceIndex = findColumnIndex(
      internationalFlightSheet,
      "保险费"
    );
    const intlChangeFeeIndex = findColumnIndex(
      internationalFlightSheet,
      "改签费"
    );
    const intlRefundFeeIndex = findColumnIndex(
      internationalFlightSheet,
      "退票费"
    );

    console.log("国际机票所有可用列:", intlHeaders);
    console.log(
      `国际机票列索引: 总金额=${intlTotalAmountIndex}, 系统使用费=${intlSystemFeeIndex}, 保险费=${intlInsuranceIndex}, 改签费=${intlChangeFeeIndex}, 退票费=${intlRefundFeeIndex}`
    );

    // 检查是否所有必需的列都找到了
    if (intlTotalAmountIndex === -1) {
      console.error("未找到'总金额'列！");
    }
    if (intlSystemFeeIndex === -1) {
      console.error("未找到'系统使用费'列！");
    }
    if (intlInsuranceIndex === -1) {
      console.error("未找到'保险费'列！");
    }
    if (intlChangeFeeIndex === -1) {
      console.error("未找到'改签费'列！");
    }
    if (intlRefundFeeIndex === -1) {
      console.error("未找到'退票费'列！");
    }

    internationalFlightSheet.eachRow((row: any, rowNumber: number) => {
      if (rowNumber > 1) {
        // 跳过表头
        // 检查是否是合计行
        const firstCell = row.getCell(1);
        const isTotalRow = firstCell.value === "合计";

        // 记录前几行的数据
        if (rowNumber <= 3) {
          console.log(`国际机票第${rowNumber}行数据:`, {
            第一列: firstCell.value,
            总金额:
              intlTotalAmountIndex !== -1
                ? row.getCell(intlTotalAmountIndex).value
                : "未找到列",
            系统使用费:
              intlSystemFeeIndex !== -1
                ? row.getCell(intlSystemFeeIndex).value
                : "未找到列",
            保险费:
              intlInsuranceIndex !== -1
                ? row.getCell(intlInsuranceIndex).value
                : "未找到列",
            是合计行: isTotalRow
          });
        }

        if (!isTotalRow) {
          // 计算总金额
          if (intlTotalAmountIndex !== -1) {
            const totalAmountCell = row.getCell(intlTotalAmountIndex);
            if (
              totalAmountCell.value &&
              typeof totalAmountCell.value === "number"
            ) {
              internationalFlightTotal += totalAmountCell.value;
            }
          }

          // 计算系统使用费
          if (intlSystemFeeIndex !== -1) {
            const systemFeeCell = row.getCell(intlSystemFeeIndex);
            if (
              systemFeeCell.value &&
              typeof systemFeeCell.value === "number"
            ) {
              intlSystemFeeTotal += systemFeeCell.value;
            }
          }

          // 计算保险费
          if (intlInsuranceIndex !== -1) {
            const insuranceCell = row.getCell(intlInsuranceIndex);
            if (
              insuranceCell.value &&
              typeof insuranceCell.value === "number"
            ) {
              intlInsuranceTotal += insuranceCell.value;
            }
          }

          // 计算改签手续费
          if (intlChangeFeeIndex !== -1) {
            const changeFeeCell = row.getCell(intlChangeFeeIndex);
            if (
              changeFeeCell.value &&
              typeof changeFeeCell.value === "number"
            ) {
              intlChangeFeeTotal += changeFeeCell.value;
            }
          }

          // 计算退票手续费
          if (intlRefundFeeIndex !== -1) {
            const refundFeeCell = row.getCell(intlRefundFeeIndex);
            if (
              refundFeeCell.value &&
              typeof refundFeeCell.value === "number"
            ) {
              intlRefundFeeTotal += refundFeeCell.value;
            }
          }
        }
      }
    });

    console.log("国际机票计算结果:", {
      合计: internationalFlightTotal,
      系统使用费: intlSystemFeeTotal,
      保险费: intlInsuranceTotal,
      改签手续费: intlChangeFeeTotal,
      退票手续费: intlRefundFeeTotal
    });

    // 详细分析数据
    console.log("=== 国际机票数据详细分析 ===");
    internationalFlightSheet.eachRow((row: any, rowNumber: number) => {
      if (rowNumber > 1 && rowNumber <= 10) {
        // 只显示前10行数据
        const firstCell = row.getCell(1);
        const totalAmount =
          intlTotalAmountIndex !== -1
            ? row.getCell(intlTotalAmountIndex).value
            : 0;
        const systemFee =
          intlSystemFeeIndex !== -1 ? row.getCell(intlSystemFeeIndex).value : 0;
        const insurance =
          intlInsuranceIndex !== -1 ? row.getCell(intlInsuranceIndex).value : 0;
        const changeFee =
          intlChangeFeeIndex !== -1 ? row.getCell(intlChangeFeeIndex).value : 0;
        const refundFee =
          intlRefundFeeIndex !== -1 ? row.getCell(intlRefundFeeIndex).value : 0;

        console.log(`第${rowNumber}行 (${firstCell.value}):`, {
          总金额: totalAmount,
          系统使用费: systemFee,
          保险费: insurance,
          改签手续费: changeFee,
          退票手续费: refundFee,
          是否为负数退票: typeof refundFee === "number" && refundFee < 0
        });
      }
    });

    // 计算国际机票费
    intlFlightFee =
      internationalFlightTotal -
      intlSystemFeeTotal -
      intlInsuranceTotal -
      intlChangeFeeTotal -
      intlRefundFeeTotal;

    console.log("国际机票费计算详情:", {
      合计金额: internationalFlightTotal,
      减去系统使用费: intlSystemFeeTotal,
      减去保险费: intlInsuranceTotal,
      减去改签手续费: intlChangeFeeTotal,
      减去退票手续费: intlRefundFeeTotal,
      计算出的机票费: intlFlightFee,
      是否为负数: intlFlightFee < 0,
      警告: intlFlightFee < 0 ? "机票费为负数，可能数据结构有问题！" : ""
    });
  }

  // 更新账单标题中的日期
  overviewWorksheet.eachRow((row: any) => {
    const firstCell = row.getCell(1);
    if (firstCell && firstCell.value) {
      const cellText = firstCell.value.toString();
      if (cellText.includes("深圳市艾比森光电股份有限公司账单")) {
        // 获取当前年月
        const now = new Date();
        const year = now.getFullYear();
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const yearMonth = `${year}${month}`;

        // 替换标题中的日期部分
        const newTitle = cellText.replace(/\d{6}/, yearMonth);
        firstCell.value = newTitle;
        console.log(`账单标题已更新: ${newTitle}`);
      }
    }
  });

  // 更新付款日期为下个月的14号
  overviewWorksheet.eachRow((row: any) => {
    // 检查每个单元格
    for (let col = 1; col <= 20; col++) {
      const cell = row.getCell(col);
      if (cell && cell.value) {
        const cellText = cell.value.toString();
        if (cellText.includes("请在YYYY-MM-DD前，将本期应还金额付款到以下账户：")) {
          // 获取下个月的日期
          const now = new Date();
          const nextMonth = new Date(now.getFullYear(), now.getMonth() + 1, 14);
          const year = nextMonth.getFullYear();
          const month = String(nextMonth.getMonth() + 1).padStart(2, '0');
          const day = String(nextMonth.getDate()).padStart(2, '0');
          const formattedDate = `${year}-${month}-${day}`;

          // 替换日期部分
          const newText = cellText.replace("YYYY-MM-DD", formattedDate);
          cell.value = newText;
          console.log(`付款日期已更新: ${newText}`);
        }
      }
    }
  });

  // 在概览工作表中查找国内机票费用行并填充数据
  overviewWorksheet.eachRow((row: any) => {
    const firstCell = row.getCell(1);
    if (
      firstCell.value &&
      firstCell.value.toString().includes("国内机票费用")
    ) {
      // 填充合计
      const totalCell = row.getCell(2); // 合计列
      totalCell.value = domesticFlightTotal;
      totalCell.numFmt = "#,##0.00";

      // 填充系统使用费
      const systemFeeCell = row.getCell(3);
      systemFeeCell.value = systemFeeTotal;
      systemFeeCell.numFmt = "#,##0.00";

      // 填充保险费
      const insuranceCell = row.getCell(4);
      insuranceCell.value = insuranceTotal;
      insuranceCell.numFmt = "#,##0.00";

      // 填充改签手续费
      const changeFeeCell = row.getCell(5);
      changeFeeCell.value = changeFeeTotal;
      changeFeeCell.numFmt = "#,##0.00";

      // 填充退票手续费
      const refundFeeCell = row.getCell(6);
      refundFeeCell.value = refundFeeTotal;
      refundFeeCell.numFmt = "#,##0.00";

      // 填充机票费
      const flightFeeCell = row.getCell(7);
      flightFeeCell.value = flightFee;
      flightFeeCell.numFmt = "#,##0.00";

      console.log(
        `国内机票费用数据已填充: 合计=${domesticFlightTotal}, 系统使用费=${systemFeeTotal}, 保险费=${insuranceTotal}, 改签手续费=${changeFeeTotal}, 退票手续费=${refundFeeTotal}, 机票费=${flightFee}`
      );
    }
  });

  // 在概览工作表中查找国内酒店费用行并填充数据
  overviewWorksheet.eachRow((row: any) => {
    const firstCell = row.getCell(1);
    if (
      firstCell.value &&
      firstCell.value.toString().includes("国内酒店费用")
    ) {
      // 填充合计
      const totalCell = row.getCell(2); // 合计列
      totalCell.value = domesticHotelTotal;
      totalCell.numFmt = "#,##0.00";

      // 填充托管费 - 固定为0.00
      const custodyFeeCell = row.getCell(3);
      custodyFeeCell.value = 0;
      custodyFeeCell.numFmt = "#,##0.00";

      // 填充代购费 - 固定为0.00
      const purchaseFeeCell = row.getCell(4);
      purchaseFeeCell.value = 0;
      purchaseFeeCell.numFmt = "#,##0.00";

      // 填充系统使用费
      const systemFeeCell = row.getCell(5);
      systemFeeCell.value = hotelSystemFeeTotal;
      systemFeeCell.numFmt = "#,##0.00";

      // 填充退订手续费
      const cancelFeeCell = row.getCell(6);
      cancelFeeCell.value = hotelCancelFeeTotal;
      cancelFeeCell.numFmt = "#,##0.00";

      // 填充酒店费
      const hotelFeeCell = row.getCell(7);
      hotelFeeCell.value = hotelFee;
      hotelFeeCell.numFmt = "#,##0.00";

      console.log(
        `国内酒店费用数据已填充: 合计=${domesticHotelTotal}, 系统使用费=${hotelSystemFeeTotal}, 退订手续费=${hotelCancelFeeTotal}, 酒店费=${hotelFee}`
      );
    }
  });

  // 在概览工作表中查找国际酒店费用行并填充数据
  overviewWorksheet.eachRow((row: any) => {
    const firstCell = row.getCell(1);
    if (
      firstCell.value &&
      firstCell.value.toString().includes("国际酒店费用")
    ) {
      console.log(`找到国际酒店费用行: 第${row.number}行, 第一列内容: ${firstCell.value}`);

      // 根据用户说明的列位置填充数据
      // B列(2) → 酒店费, C列(3) → 退订手续费, D列(4) → 系统使用费, E列(5) → 合计
      const fillData = {
        2: intlHotelFee,            // B列 - 酒店费
        3: intlHotelCancelFeeTotal, // C列 - 退订手续费
        4: intlHotelSystemFeeTotal, // D列 - 系统使用费
        5: internationalHotelTotal // E列 - 合计
      };

      console.log(`准备填充国际酒店数据:`, fillData);

      // 按正确的列顺序填充
      for (let col = 2; col <= 5; col++) {
        const cell = row.getCell(col);
        cell.value = fillData[col];
        cell.numFmt = "#,##0.00";
        console.log(`${getCellLabel(col)}列(第${col}列)已填充: ${fillData[col]}`);
      }

      // 验证填充结果
      console.log(`国际酒店费用行验证:`);
      for (let i = 1; i <= 7; i++) {
        const cell = row.getCell(i);
        console.log(`  第${i}列 (${getCellLabel(i)}): ${cell.value}`);
      }

      // 强制提交更改
      row.commit();

      console.log(
        `国际酒店费用数据已填充: 合计=${internationalHotelTotal}, 系统使用费=${intlHotelSystemFeeTotal}, 退订手续费=${intlHotelCancelFeeTotal}, 酒店费=${intlHotelFee}`
      );
    }
  });

  // 在概览工作表中查找国际机票费用行并填充数据
  overviewWorksheet.eachRow((row: any) => {
    const firstCell = row.getCell(1);
    if (firstCell.value && firstCell.value.toString().includes("国际机票费用")) {
      console.log(`找到国际机票费用行: 第${row.number}行, 第一列内容: ${firstCell.value}`);

      // 根据用户说明的列位置填充数据
      // B列(2) → 机票费, C列(3) → 退票手续费, D列(4) → 改签手续费, E列(5) → 保险费, F列(6) → 系统使用费, G列(7) → 合计
      const fillData = {
        2: intlFlightFee,          // B列 - 机票费
        3: intlRefundFeeTotal,     // C列 - 退票手续费
        4: intlChangeFeeTotal,     // D列 - 改签手续费
        5: intlInsuranceTotal,     // E列 - 保险费
        6: intlSystemFeeTotal,     // F列 - 系统使用费
        7: internationalFlightTotal // G列 - 合计
      };

      console.log(`准备填充的数据:`, fillData);

      // 按正确的列顺序填充
      for (let col = 2; col <= 7; col++) {
        const cell = row.getCell(col);
        cell.value = fillData[col];
        cell.numFmt = "#,##0.00";
        console.log(`${getCellLabel(col)}列(第${col}列)已填充: ${fillData[col]}`);
      }

      // 验证填充结果
      console.log(`国际机票费用行验证:`);
      for (let i = 1; i <= 7; i++) {
        const cell = row.getCell(i);
        console.log(`  第${i}列 (${getCellLabel(i)}): ${cell.value}`);
      }

      // 强制提交更改
      row.commit();
    }
  });

  // 辅助函数：将列号转换为Excel列标签
  function getCellLabel(colNumber: number): string {
    let label = '';
    while (colNumber > 0) {
      colNumber--;
      label = String.fromCharCode(65 + (colNumber % 26)) + label;
      colNumber = Math.floor(colNumber / 26);
    }
    return label;
  }

  // 直接在D2和G2单元格设置公式
  console.log("开始设置D2和G2单元格的公式");

  // 先找到各费用行的行号
  let domesticFlightRow = 0;
  let domesticHotelRow = 0;
  let internationalFlightRow = 0;
  let internationalHotelRow = 0;

  overviewWorksheet.eachRow((searchRow: any) => {
    const searchFirstCell = searchRow.getCell(1);
    if (searchFirstCell && searchFirstCell.value) {
      const cellText = searchFirstCell.value.toString();
      if (cellText.includes("国内机票费用")) {
        domesticFlightRow = searchRow.number;
      } else if (cellText.includes("国内酒店费用")) {
        domesticHotelRow = searchRow.number;
      } else if (cellText.includes("国际机票费用")) {
        internationalFlightRow = searchRow.number;
      } else if (cellText.includes("国际酒店费用")) {
        internationalHotelRow = searchRow.number;
      }
    }
  });

  console.log(`找到各费用行: 国内机票=${domesticFlightRow}, 国内酒店=${domesticHotelRow}, 国际机票=${internationalFlightRow}, 国际酒店=${internationalHotelRow}`);

  // 在D3单元格设置公式（待还款金额）
  // 待还款金额 = 本期账单总金额
  const d3Cell = overviewWorksheet.getCell(3, 4); // 第3行第4列（D3）
  if (domesticFlightRow > 0 && domesticHotelRow > 0 && internationalFlightRow > 0 && internationalHotelRow > 0) {
    // 先设置G3的公式（本期账单总金额）
    const g3Cell = overviewWorksheet.getCell(3, 7); // 第3行第7列（G3）
    const totalFormula = `=G${domesticFlightRow}+G${domesticHotelRow}+G${internationalFlightRow}+E${internationalHotelRow}`;
    g3Cell.value = { formula: totalFormula };
    g3Cell.numFmt = "#,##0.00";
    console.log(`G3单元格公式已设置: ${totalFormula}`);

    // 再设置D3的公式（待还款金额 = 本期账单总金额）
    const repaymentFormula = `=G3`;
    d3Cell.value = { formula: repaymentFormula };
    d3Cell.numFmt = "#,##0.00";
    console.log(`D3单元格公式已设置: ${repaymentFormula}`);
  }

  console.log("账单概览数据填充完成");

  // 为H列左侧添加细边框
  addLeftBorderToColumnH(overviewWorksheet);

  // 隐藏合计为0的费用行
  hideZeroTotalRows(overviewWorksheet);
};

// 为H列左侧添加细边框
const addLeftBorderToColumnH = (overviewWorksheet: any) => {
  console.log("开始为H列左侧添加细边框");

  // 遍历所有行，为H列（第8列）的左侧添加边框
  overviewWorksheet.eachRow((row: any) => {
    // 获取H列单元格
    const cell = row.getCell(8); // H列是第8列

    // 如果没有样式对象，创建一个
    if (!cell.style) {
      cell.style = {};
    }

    // 保留现有的边框样式
    const existingBorders = cell.style.border || {};

    // 添加左侧细边框
    cell.style.border = {
      ...existingBorders,
      left: {
        style: 'thin',
        color: { argb: 'FF000000' } // 黑色
      }
    };
  });

  console.log("H列左侧细边框添加完成");
};

// 隐藏合计为0的费用行及其上面两行
const hideZeroTotalRows = (overviewWorksheet: any) => {
  console.log("开始处理合计为0的费用行隐藏");

  // 定义需要检查的费用行
  const feeTypes = [
    "国内机票费用",
    "国内酒店费用",
    "国际机票费用",
    "国际酒店费用"
  ];

  // 收集需要隐藏的行号
  const rowsToHide: number[] = [];

  overviewWorksheet.eachRow((row: any) => {
    const firstCell = row.getCell(1);
    if (firstCell && firstCell.value) {
      const cellText = firstCell.value.toString();

      // 检查是否是需要处理的费用行
      for (const feeType of feeTypes) {
        if (cellText.includes(feeType)) {
          // 获取合计值（第2列）
          const totalCell = row.getCell(2);
          const totalValue = totalCell.value;

          console.log(`检查 ${feeType} 行（第${row.number}行），合计值: ${totalValue}`);

          // 如果合计为0或null/undefined，标记需要隐藏的行
          if (!totalValue || totalValue === 0) {
            console.log(`${feeType} 合计为0，需要隐藏第${row.number}行及其上面两行`);

            // 添加当前行和上面两行到隐藏列表
            rowsToHide.push(row.number); // 费用行本身
            rowsToHide.push(row.number - 1); // 上面一行
            rowsToHide.push(row.number - 2); // 上面两行
          }
        }
      }
    }
  });

  // 去重并排序需要隐藏的行号
  const uniqueRowsToHide = [...new Set(rowsToHide)].sort((a, b) => a - b);

  console.log(`需要隐藏的行号: ${uniqueRowsToHide.join(', ')}`);

  // 隐藏指定的行 - 使用Excel的hidden属性
  uniqueRowsToHide.forEach(rowNumber => {
    if (rowNumber > 0) { // 确保行号有效
      const row = overviewWorksheet.getRow(rowNumber);

      // 设置行隐藏属性
      row.hidden = true;

      // 可选：同时也设置行高为0（双重保险）
      row.height = 0;

      // 设置所有单元格为空值和透明样式（确保在取消隐藏时不会显示内容）
      row.eachCell((cell: any) => {
        cell.value = null;
        cell.fill = {
          type: 'pattern',
          pattern: 'none',
          fgColor: { argb: 'FFFFFFFF' }, // 白色背景
          bgColor: { argb: 'FFFFFFFF' }
        };
        cell.font = {
          color: { argb: 'FFFFFFFF' } // 白色字体
        };
      });

      // 提交行的更改（如果ExcelJS支持的话）
      // 注意：不是所有版本的ExcelJS都有commit方法
      if (typeof row.commit === 'function') {
        row.commit();
      }
    }
  });

  console.log("行隐藏处理完成 - 使用hidden属性");
};

const updateFileName = (row: any) => {
  console.log(`更新文件名: ${row.fileName}`);
  // 可以在这里添加验证逻辑
  if (!row.fileName) {
    ElMessage.warning("文件名不能为空");
    // 重新生成默认文件名
    const dateRange = getDateRange();
    row.fileName = `特航账单-${row.departmentName}${dateRange}.xlsx`;
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
