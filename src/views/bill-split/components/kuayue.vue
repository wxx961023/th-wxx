<script setup lang="ts">
import { ref } from "vue";
import { ElMessage } from "element-plus";
import { UploadFilled } from "@element-plus/icons-vue";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import * as XLSX from "xlsx";

defineOptions({
  name: "KuayueBillSplit"
});

// 机票新表头定义
const FLIGHT_NEW_HEADERS = [
  "序号",
  "成本中心",
  "业务类型",
  "订单号",
  "预订日期(yyyy-MM-dd)",
  "出发时间(yyyy-MM-dd HH:mm)",
  "行程",
  "航班号",
  "登机人",
  "登机人工号",
  "状态",
  "票号",
  "消费金额",
  "员工自付",
  "机票价格",
  "机建费",
  "燃油费",
  "退票费",
  "改签费",
  "系统使用费",
  "供应商",
  "审批单号",
  "工资区域",
  "行程单金额"
];

// 机票旧表头名称
const FLIGHT_OLD_HEADERS = [
  "业务类型",
  "订单号（重复项标亮）",
  "预订日期",
  "出发时间",
  "行程",
  "航班号",
  "登机人",
  "登机人工号",
  "状态",
  "票号",
  "个人支付金额",
  "折扣价",
  "机建费",
  "燃油费",
  "退票费",
  "改签费",
  "系统使用费",
  "特航商旅",
  "审批单号"
];

// 机票表头映射：新表头索引 -> 旧表头名称
const FLIGHT_HEADER_MAPPING: Record<number, string> = {
  2: "业务类型",
  3: "订单号（重复项标亮）",
  4: "预订日期",
  5: "出发时间",
  6: "行程",
  7: "航班号",
  8: "登机人",
  9: "登机人工号",
  10: "状态",
  11: "票号",
  13: "个人支付金额",
  14: "折扣价",
  15: "机建费",
  16: "燃油费",
  17: "退票费",
  18: "改签费",
  19: "系统使用费",
  20: "特航商旅",
  21: "审批单号"
};

// 机票文件相关
const flightFile = ref<File | null>(null);
const flightData = ref<{ headers: any[]; data: any[][]; transformedData: any[][] } | null>(null);
const flightLoading = ref(false);

// 酒店文件相关
const hotelFile = ref<File | null>(null);
const hotelData = ref<{ headers: any[]; data: any[][] } | null>(null);
const hotelLoading = ref(false);

// 汇总相关
const summarizing = ref(false);
const summaryData = ref<{ headers: any[]; data: any[][] } | null>(null);
const showSummary = ref(false);
const generating = ref(false);

// 文件上传校验
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

// 读取Excel文件（支持 .xls 和 .xlsx）
const readExcelFile = async (file: File): Promise<{ headers: any[]; data: any[][] }> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const buffer = e.target?.result as ArrayBuffer;
        const isXls = file.name.toLowerCase().endsWith(".xls");

        let rows: any[][] = [];

        if (isXls) {
          // 使用 xlsx 库读取 .xls 文件
          const workbook = XLSX.read(buffer, { type: "array" });
          const firstSheetName = workbook.SheetNames[0];
          if (!firstSheetName) {
            reject(new Error("未找到工作表"));
            return;
          }
          const worksheet = workbook.Sheets[firstSheetName];
          rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
        } else {
          // 使用 ExcelJS 读取 .xlsx 文件
          const workbook = new ExcelJS.Workbook();
          await workbook.xlsx.load(buffer);

          const worksheet = workbook.worksheets[0];
          if (!worksheet) {
            reject(new Error("未找到工作表"));
            return;
          }

          worksheet.eachRow((row) => {
            const rowData: any[] = [];
            row.eachCell({ includeEmpty: true }, (cell) => {
              rowData.push(cell.value);
            });
            rows.push(rowData);
          });
        }

        if (rows.length === 0) {
          reject(new Error("文件无数据"));
          return;
        }

        resolve({
          headers: rows[0],
          data: rows.slice(1)
        });
      } catch (error: any) {
        console.error("读取Excel失败:", error);
        reject(new Error(error.message || "读取Excel文件失败"));
      }
    };
    reader.onerror = () => reject(new Error("文件读取失败"));
    reader.readAsArrayBuffer(file);
  });
};

// 构建表头索引映射（支持模糊匹配）
const buildHeaderIndexMap = (headers: any[]): Map<string, number> => {
  const map = new Map<string, number>();
  headers.forEach((h, i) => {
    if (h) {
      map.set(h.toString().trim(), i);
    }
  });

  // 打印实际表头，方便调试
  console.log("实际表头:", headers);

  return map;
};

// 根据表头名称查找索引（支持模糊匹配）
const findHeaderIndex = (
  headerIndexMap: Map<string, number>,
  headerName: string
): number | undefined => {
  // 先尝试精确匹配
  if (headerIndexMap.has(headerName)) {
    return headerIndexMap.get(headerName);
  }

  // 尝试模糊匹配（忽略括号内容和空格）
  const normalizedHeader = headerName.replace(/[（）()（重复项标亮）]/g, "").trim();

  for (const [key, value] of headerIndexMap.entries()) {
    const normalizedKey = key.replace(/[（）()（重复项标亮）]/g, "").trim();
    if (normalizedKey === normalizedHeader || key.includes(normalizedHeader) || normalizedHeader.includes(key)) {
      console.log(`表头模糊匹配: "${headerName}" -> "${key}" (索引: ${value})`);
      return value;
    }
  }

  console.warn(`未找到表头: "${headerName}"`);
  return undefined;
};

// 格式化日期为 yyyy-MM-dd
const formatDate = (value: any): string => {
  if (!value) return "";

  // 如果是 Date 对象
  if (value instanceof Date) {
    const year = value.getFullYear();
    const month = String(value.getMonth() + 1).padStart(2, "0");
    const day = String(value.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }

  // 如果是字符串，尝试解析
  const str = String(value).trim();

  // 已经是 yyyy-MM-dd 格式
  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
    return str;
  }

  // yyyy/MM/dd 格式
  if (/^\d{4}\/\d{2}\/\d{2}$/.test(str)) {
    return str.replace(/\//g, "-");
  }

  // 包含时间的格式 yyyy-MM-dd HH:mm:ss 或 yyyy/MM/dd HH:mm:ss
  const match = str.match(/^(\d{4})[-/](\d{2})[-/](\d{2})/);
  if (match) {
    return `${match[1]}-${match[2]}-${match[3]}`;
  }

  return str;
};

// 格式化日期时间为 yyyy-MM-dd HH:mm
const formatDateTime = (value: any): string => {
  if (!value) return "";

  // 如果是 Date 对象
  if (value instanceof Date) {
    const year = value.getFullYear();
    const month = String(value.getMonth() + 1).padStart(2, "0");
    const day = String(value.getDate()).padStart(2, "0");
    const hour = String(value.getHours()).padStart(2, "0");
    const minute = String(value.getMinutes()).padStart(2, "0");
    return `${year}-${month}-${day} ${hour}:${minute}`;
  }

  // 如果是字符串，尝试解析
  const str = String(value).trim();

  // 已经是 yyyy-MM-dd HH:mm 格式
  if (/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/.test(str)) {
    return str;
  }

  // yyyy-MM-dd HH:mm:ss 格式
  if (/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(str)) {
    return str.substring(0, 16);
  }

  // yyyy/MM/dd HH:mm:ss 格式
  const match = str.match(/^(\d{4})[-/](\d{2})[-/](\d{2})[ T](\d{2}):(\d{2})/);
  if (match) {
    return `${match[1]}-${match[2]}-${match[3]} ${match[4]}:${match[5]}`;
  }

  return str;
};

// 格式化金额为两位小数
const formatAmount = (value: any): string => {
  const num = parseFloat(value);
  if (isNaN(num)) return "";
  return num.toFixed(2);
};

// 转换机票数据
const transformFlightData = (
  rows: any[][],
  headerIndexMap: Map<string, number>
): any[][] => {
  const transformedData: any[][] = [];

  for (let i = 0; i < rows.length; i++) {
    const oldRow = rows[i];
    const newRow: any[] = new Array(FLIGHT_NEW_HEADERS.length).fill("");

    // 序号
    newRow[0] = i + 1;

    // 成本中心 - 空
    newRow[1] = "";

    // 根据映射填充数据（使用模糊匹配）
    for (const [newIdx, oldHeader] of Object.entries(FLIGHT_HEADER_MAPPING)) {
      const oldIdx = findHeaderIndex(headerIndexMap, oldHeader);
      if (oldIdx !== undefined) {
        newRow[parseInt(newIdx)] = oldRow[oldIdx] ?? "";
      }
    }

    // 预订日期格式化
    newRow[4] = formatDate(newRow[4]);

    // 出发时间格式化
    newRow[5] = formatDateTime(newRow[5]);

    // 格式化金额列
    newRow[13] = formatAmount(newRow[13]); // 员工自付
    newRow[14] = formatAmount(newRow[14]); // 机票价格
    newRow[15] = formatAmount(newRow[15]); // 机建费
    newRow[16] = formatAmount(newRow[16]); // 燃油费
    newRow[17] = formatAmount(newRow[17]); // 退票费
    newRow[18] = formatAmount(newRow[18]); // 改签费
    newRow[19] = formatAmount(newRow[19]); // 系统使用费

    // 消费金额 = 机票价格+机建费+燃油费+改签费+系统使用费
    const consumeAmount =
      parseFloat(newRow[14] || 0) +
      parseFloat(newRow[15] || 0) +
      parseFloat(newRow[16] || 0) +
      parseFloat(newRow[18] || 0) +
      parseFloat(newRow[19] || 0);
    newRow[12] = consumeAmount.toFixed(2);

    // 供应商 - 固定写"特航商旅"
    newRow[20] = "特航商旅";

    // 工资区域 - 空
    newRow[22] = "";

    // 行程单金额 - 空
    newRow[23] = "";

    transformedData.push(newRow);
  }

  return transformedData;
};

// 处理机票文件上传
const handleFlightFileChange = async (uploadFile: any) => {
  const file = uploadFile.raw;
  if (!file) return;

  if (!beforeUpload(file)) return;

  flightLoading.value = true;
  try {
    const result = await readExcelFile(file);
    const headerIndexMap = buildHeaderIndexMap(result.headers);

    // 转换数据
    const transformedData = transformFlightData(result.data, headerIndexMap);

    flightFile.value = file;
    flightData.value = {
      headers: result.headers,
      data: result.data,
      transformedData
    };

    console.log("机票表头映射:", Object.fromEntries(headerIndexMap));
    console.log(`机票数据转换完成，共 ${transformedData.length} 条`);

    ElMessage.success(`机票文件上传成功，共 ${transformedData.length} 条数据`);
  } catch (error: any) {
    ElMessage.error(error.message || "读取机票文件失败");
    flightFile.value = null;
    flightData.value = null;
  } finally {
    flightLoading.value = false;
  }
};

// 处理酒店文件上传
const handleHotelFileChange = async (uploadFile: any) => {
  const file = uploadFile.raw;
  if (!file) return;

  if (!beforeUpload(file)) return;

  hotelLoading.value = true;
  try {
    const result = await readExcelFile(file);
    hotelFile.value = file;
    hotelData.value = result;
    ElMessage.success(`酒店文件上传成功，共 ${result.data.length} 条数据`);
  } catch (error: any) {
    ElMessage.error(error.message || "读取酒店文件失败");
    hotelFile.value = null;
    hotelData.value = null;
  } finally {
    hotelLoading.value = false;
  }
};

// 清除机票文件
const clearFlightFile = () => {
  flightFile.value = null;
  flightData.value = null;
  showSummary.value = false;
  summaryData.value = null;
};

// 清除酒店文件
const clearHotelFile = () => {
  hotelFile.value = null;
  hotelData.value = null;
  showSummary.value = false;
  summaryData.value = null;
};

// 汇总数据
const summarizeData = () => {
  if (!flightData.value && !hotelData.value) {
    ElMessage.warning("请先上传机票或酒店文件");
    return;
  }

  summarizing.value = true;

  try {
    const allData: any[][] = [];

    // 使用机票新表头
    let headers = FLIGHT_NEW_HEADERS;

    // 添加转换后的机票数据
    if (flightData.value?.transformedData) {
      allData.push(...flightData.value.transformedData);
    }

    // TODO: 添加酒店数据处理

    summaryData.value = {
      headers,
      data: allData
    };
    showSummary.value = true;

    ElMessage.success(`汇总完成，共 ${allData.length} 条数据`);
  } catch (error) {
    console.error("汇总失败:", error);
    ElMessage.error("汇总失败");
  } finally {
    summarizing.value = false;
  }
};

// 生成并下载Excel
const generateExcel = async () => {
  if (!summaryData.value || summaryData.value.data.length === 0) {
    ElMessage.warning("没有可导出的数据");
    return;
  }

  generating.value = true;

  try {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("国内机票");

    // 计算动态日期（当前日期向前推1个月）
    const now = new Date();
    const targetDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    const year = targetDate.getFullYear();
    const month = targetDate.getMonth() + 1;
    const titleText = `跨越速运集团有限公司${year}年${month}月账单`;

    // 第一行：标题行
    const titleRow = worksheet.addRow([titleText]);
    titleRow.height = 30;
    const titleCell = titleRow.getCell(1);
    titleCell.font = { bold: true, size: 14 };
    titleCell.alignment = { horizontal: "center", vertical: "middle" };

    // 合并标题行单元格
    worksheet.mergeCells(1, 1, 1, summaryData.value.headers.length);

    // 第二行：表头
    worksheet.addRow(summaryData.value.headers);

    // 添加数据
    for (const row of summaryData.value.data) {
      worksheet.addRow(row);
    }

    // 设置表头样式（第2行）
    const headerRow = worksheet.getRow(2);
    headerRow.height = 22;
    headerRow.eachCell((cell) => {
      cell.font = { bold: true };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFFF99" }
      };
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" }
      };
    });

    // 设置数据行样式（从第3行开始）
    for (let i = 3; i <= worksheet.rowCount; i++) {
      const row = worksheet.getRow(i);
      row.height = 22;
      row.eachCell((cell) => {
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.font = { size: 10 };
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" }
        };
      });
    }

    // 检测重复订单号和票号并高亮
    // 订单号列（索引3，第4列D列），票号列（索引11，第12列L列）
    const duplicateCols = [
      { name: "订单号", col: 4 },   // D列
      { name: "票号", col: 12 }     // L列
    ];

    duplicateCols.forEach(({ name, col }) => {
      const valueMap: Map<string, number[]> = new Map();

      // 收集所有值及其行号
      for (let i = 3; i <= worksheet.rowCount; i++) {
        const cell = worksheet.getRow(i).getCell(col);
        const value = cell.value?.toString()?.trim();
        if (value) {
          if (!valueMap.has(value)) {
            valueMap.set(value, []);
          }
          valueMap.get(value)!.push(i);
        }
      }

      // 对重复的值填充底色
      for (const [, rows] of valueMap) {
        if (rows.length > 1) {
          rows.forEach((rowNum) => {
            const cell = worksheet.getRow(rowNum).getCell(col);
            cell.fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: "FFFF9900" }
            };
          });
        }
      }
    });

    // 自动调整列宽（跳过第1行标题，从第2行表头开始计算）
    worksheet.columns.forEach((column) => {
      let maxWidth = 8;
      column.eachCell?.({ includeEmpty: true }, (cell, rowNumber) => {
        // 跳过第1行标题行
        if (rowNumber === 1) return;
        const cellValue = cell.value?.toString() || "";
        let width = 0;
        for (const char of cellValue) {
          if (/[\u4e00-\u9fa5]/.test(char)) {
            width += 2;
          } else {
            width += 1;
          }
        }
        maxWidth = Math.max(maxWidth, width);
      });
      column.width = Math.min(maxWidth + 2, 30);
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    saveAs(blob, `跨越速运集团有限公司${year}年${month}月账单.xlsx`);
    ElMessage.success("文件生成成功！");
  } catch (error) {
    console.error("生成文件失败:", error);
    ElMessage.error("生成文件失败");
  } finally {
    generating.value = false;
  }
};

// 重置所有数据
const resetAll = () => {
  flightFile.value = null;
  flightData.value = null;
  hotelFile.value = null;
  hotelData.value = null;
  summaryData.value = null;
  showSummary.value = false;
};
</script>

<template>
  <div class="kuayue-bill-split">
    <!-- 文件上传区域 -->
    <div class="upload-section">
      <!-- 机票Excel上传 -->
      <el-card class="upload-card">
        <template #header>
          <div class="card-header">
            <span>机票Excel</span>
            <el-button
              v-if="flightFile"
              type="danger"
              size="small"
              @click="clearFlightFile"
            >
              清除
            </el-button>
          </div>
        </template>

        <el-upload
          v-if="!flightFile"
          class="upload-area"
          drag
          :auto-upload="false"
          :show-file-list="false"
          :on-change="handleFlightFileChange"
          accept=".xlsx,.xls"
        >
          <el-icon class="el-icon--upload" :size="50">
            <UploadFilled />
          </el-icon>
          <div class="el-upload__text">
            拖拽机票Excel文件到此处，或<em>点击上传</em>
          </div>
        </el-upload>

        <div v-else class="file-uploaded">
          <el-icon :size="40" color="#67C23A">
            <i-ep-circle-check-filled />
          </el-icon>
          <div class="file-info">
            <span class="file-name">{{ flightFile.name }}</span>
            <span class="file-count">共 {{ flightData?.transformedData?.length || 0 }} 条数据</span>
          </div>
        </div>
      </el-card>

      <!-- 酒店Excel上传 -->
      <el-card class="upload-card">
        <template #header>
          <div class="card-header">
            <span>酒店Excel</span>
            <el-button
              v-if="hotelFile"
              type="danger"
              size="small"
              @click="clearHotelFile"
            >
              清除
            </el-button>
          </div>
        </template>

        <el-upload
          v-if="!hotelFile"
          class="upload-area"
          drag
          :auto-upload="false"
          :show-file-list="false"
          :on-change="handleHotelFileChange"
          accept=".xlsx,.xls"
        >
          <el-icon class="el-icon--upload" :size="50">
            <UploadFilled />
          </el-icon>
          <div class="el-upload__text">
            拖拽酒店Excel文件到此处，或<em>点击上传</em>
          </div>
        </el-upload>

        <div v-else class="file-uploaded">
          <el-icon :size="40" color="#67C23A">
            <i-ep-circle-check-filled />
          </el-icon>
          <div class="file-info">
            <span class="file-name">{{ hotelFile.name }}</span>
            <span class="file-count">共 {{ hotelData?.data.length || 0 }} 条数据</span>
          </div>
        </div>
      </el-card>
    </div>

    <!-- 操作按钮 -->
    <div class="action-buttons">
      <el-button
        type="primary"
        size="large"
        :loading="summarizing"
        :disabled="!flightData && !hotelData"
        @click="summarizeData"
      >
        {{ summarizing ? "汇总中..." : "汇总数据" }}
      </el-button>
      <el-button size="large" @click="resetAll">
        重置
      </el-button>
    </div>

    <!-- 加载状态 -->
    <div v-if="flightLoading || hotelLoading" class="loading-container">
      <el-icon class="is-loading" :size="40">
        <i class="el-icon-loading" />
      </el-icon>
      <p>正在解析文件...</p>
    </div>

    <!-- 汇总结果预览 -->
    <el-card v-if="showSummary && summaryData" class="summary-card">
      <template #header>
        <div class="card-header">
          <span>汇总结果</span>
          <div class="header-actions">
            <span class="summary-count">共 {{ summaryData.data.length }} 条数据</span>
            <el-button
              type="primary"
              :loading="generating"
              @click="generateExcel"
            >
              {{ generating ? "生成中..." : "生成并下载" }}
            </el-button>
          </div>
        </div>
      </template>

      <el-table
        :data="summaryData.data.slice(0, 10)"
        border
        stripe
        max-height="400"
      >
        <el-table-column
          v-for="(header, index) in summaryData.headers"
          :key="index"
          :prop="String(index)"
          :label="header || `列${index}`"
          min-width="100"
        >
          <template #default="{ row }">
            {{ row[index] }}
          </template>
        </el-table-column>
      </el-table>
      <p class="preview-tip">
        仅显示前10条数据，共 {{ summaryData.data.length }} 条
      </p>
    </el-card>
  </div>
</template>

<style scoped>
.kuayue-bill-split {
  padding: 20px;
}

.upload-section {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 20px;
  margin-bottom: 20px;
}

@media (max-width: 768px) {
  .upload-section {
    grid-template-columns: 1fr;
  }
}

.upload-card {
  min-height: 200px;
}

.card-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
}

.header-actions {
  display: flex;
  align-items: center;
  gap: 15px;
}

.upload-area {
  width: 100%;
}

:deep(.el-upload-dragger) {
  width: 100%;
  height: 150px;
  display: flex;
  flex-direction: column;
  justify-content: center;
  align-items: center;
}

.file-uploaded {
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  padding: 30px;
  background: #f0f9eb;
  border-radius: 8px;
}

.file-info {
  margin-top: 10px;
  text-align: center;
}

.file-name {
  display: block;
  font-size: 14px;
  color: #303133;
  margin-bottom: 5px;
}

.file-count {
  font-size: 12px;
  color: #909399;
}

.action-buttons {
  display: flex;
  justify-content: center;
  gap: 20px;
  margin-bottom: 20px;
}

.loading-container {
  text-align: center;
  padding: 40px;
}

.loading-container p {
  margin-top: 15px;
  color: #909399;
}

.summary-card {
  margin-top: 20px;
}

.summary-count {
  font-size: 14px;
  color: #909399;
}

.preview-tip {
  color: #909399;
  font-size: 12px;
  margin-top: 10px;
  text-align: right;
}
</style>
