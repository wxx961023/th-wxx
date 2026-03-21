<script setup lang="ts">
import { ref } from "vue";
import { ElMessage } from "element-plus";
import { UploadFilled } from "@element-plus/icons-vue";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";

defineOptions({
  name: "XinteBillSplit"
});

const uploadedFile = ref<File | null>(null);
const sheetData = ref<{
  headers: any[];
  data: any[][];
} | null>(null);
const loading = ref(false);
const showData = ref(false);
const generating = ref(false);

// 新表头定义
const NEW_HEADERS = [
  "费用归属",
  "订单号",
  "记账日期",
  "类型",
  "出发城市",
  "到达城市",
  "出发日期",
  "出发时间",
  "航班号",
  "预订人",
  "乘机人",
  "票号",
  "舱位类型",
  "航线Y舱全价",
  "折扣",
  "订单状态",
  "票面价",
  "税费",
  "机建",
  "燃油",
  "机票费",
  "保险费",
  "退票费",
  "改签费",
  "系统使用费",
  "总金额"
];

// 国内机票表头映射：新表头索引 -> 旧表头名称
const DOMESTIC_HEADER_MAPPING: Record<number, string> = {
  0: "费用归属",    // 费用归属
  1: "订单号",      // 订单号
  2: "记账日期",    // 记账日期
  4: "出发城市",    // 出发城市
  5: "到达城市",    // 到达城市
  6: "出发日期",    // 出发日期
  7: "出发时间",    // 出发时间
  8: "航班号",      // 航班号
  9: "预订人",      // 预订人
  10: "乘机人",     // 乘机人
  11: "票号",       // 票号
  12: "舱位类型",   // 舱位类型
  13: "航线Y舱全价", // 航线Y舱全价
  14: "折扣",       // 折扣
  15: "订单状态",   // 订单状态
  16: "票面价",     // 票面价
  17: "税费",       // 税费
  18: "机建",       // 机建
  19: "燃油",       // 燃油
  20: "机票费",     // 机票费
  21: "保险费",     // 保险费
  22: "退票费",     // 退票费
  23: "改签费",     // 改签费
  24: "系统使用费",  // 系统使用费
  25: "总金额"      // 总金额
};

// 国际机票表头映射：新表头索引 -> 旧表头名称
const INTERNATIONAL_HEADER_MAPPING: Record<number, string> = {
  0: "费用归属",    // 费用归属
  1: "订单号",      // 订单号
  2: "记账日期",    // 记账日期
  8: "航班号",      // 航班号
  9: "预订人",      // 预订人
  10: "乘机人",     // 乘机人
  11: "票号",       // 票号
  12: "舱位类型",   // 舱位类型
  15: "订单状态",   // 订单状态
  16: "票面价",     // 票面价
  17: "税费",       // 税费
  20: "机票费",     // 机票费
  21: "保险费",     // 保险费
  22: "退票费",     // 退票费
  23: "改签费",     // 改签费
  24: "系统使用费",  // 系统使用费
  25: "总金额"      // 总金额
};

const handleFileChange = (uploadFile: any) => {
  const file = uploadFile.raw;
  if (!file) return;
  uploadedFile.value = file;
  readFile(file);
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

// 从行程/航程中提取出发城市和到达城市
const extractCities = (route: string): { from: string; to: string } => {
  if (!route) return { from: "", to: "" };

  // 如果有 "/" 分割，取第一段
  // 例如: "首尔-广州/广州-郑州/郑州-广州/广州-首尔" -> "首尔-广州"
  let firstSegment = route;
  if (route.includes("/")) {
    firstSegment = route.split("/")[0];
  }

  // 用 "-" 分割提取出发城市和到达城市
  // 例如: "首尔-广州" -> 出发: 首尔, 到达: 广州
  if (firstSegment.includes("-")) {
    const parts = firstSegment.split("-");
    if (parts.length >= 2) {
      return { from: parts[0].trim(), to: parts[1].trim() };
    }
  }

  return { from: firstSegment.trim(), to: "" };
};

// 从时间字符串中提取出发日期和出发时间
// 例如: "2026-02-12 10:55/2026-02-12 19:55/..." -> { date: "2026-02-12", time: "10:55" }
const extractDateTime = (dateTimeStr: string): { date: string; time: string } => {
  if (!dateTimeStr) return { date: "", time: "" };

  // 如果有 "/" 分割，取第一个时间
  let firstTime = dateTimeStr;
  if (dateTimeStr.includes("/")) {
    firstTime = dateTimeStr.split("/")[0];
  }

  firstTime = firstTime.trim();

  // 解析 "2026-02-12 10:55" 格式
  const match = firstTime.match(/^(\d{4}-\d{2}-\d{2})\s+(\d{2}:\d{2})$/);
  if (match) {
    return { date: match[1], time: match[2] };
  }

  // 尝试其他格式：只有日期或只有时间
  const dateOnly = firstTime.match(/^(\d{4}-\d{2}-\d{2})$/);
  if (dateOnly) {
    return { date: dateOnly[1], time: "" };
  }

  const timeOnly = firstTime.match(/^(\d{2}:\d{2})$/);
  if (timeOnly) {
    return { date: "", time: timeOnly[1] };
  }

  return { date: firstTime, time: "" };
};

// 从工作表读取数据
const readWorksheetData = (worksheet: ExcelJS.Worksheet): any[][] => {
  const rows: any[][] = [];
  worksheet.eachRow(row => {
    const rowData: any[] = [];
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      rowData[colNumber - 1] = cell.value;
    });
    rows.push(rowData);
  });
  return rows;
};

// 需要保留两位小数的金额列索引（N=13, Q=16, R=17, S=18, T=19, U=20, V=21, W=22, X=23, Y=24, Z=25）
const AMOUNT_COLUMNS = [13, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25];

// 格式化金额为两位小数
const formatAmount = (value: any): string => {
  const num = parseFloat(value);
  if (isNaN(num)) return "0.00";
  return num.toFixed(2);
};

// 构建表头索引映射
const buildHeaderIndexMap = (headers: any[]): Map<string, number> => {
  const map = new Map<string, number>();
  headers.forEach((h, i) => {
    if (h) {
      map.set(h.toString().trim(), i);
    }
  });
  return map;
};

// 转换国内机票数据
const transformDomesticData = (
  rows: any[][],
  headerIndexMap: Map<string, number>
): any[][] => {
  const transformedData: any[][] = [];

  for (let i = 1; i < rows.length; i++) {
    const oldRow = rows[i];
    const newRow: any[] = new Array(NEW_HEADERS.length).fill("");

    // 根据映射填充数据
    for (const [newIdx, oldHeader] of Object.entries(DOMESTIC_HEADER_MAPPING)) {
      const oldIdx = headerIndexMap.get(oldHeader);
      if (oldIdx !== undefined) {
        newRow[parseInt(newIdx)] = oldRow[oldIdx] ?? "";
      }
    }

    // 类型列固定填写"国内"
    newRow[3] = "国内";

    // 格式化金额列为两位小数
    AMOUNT_COLUMNS.forEach(colIdx => {
      newRow[colIdx] = formatAmount(newRow[colIdx]);
    });

    transformedData.push(newRow);
  }

  return transformedData;
};

// 转换国际机票数据
const transformInternationalData = (
  rows: any[][],
  headerIndexMap: Map<string, number>
): any[][] => {
  const routeColIndex = headerIndexMap.get("航程");
  const departTimeColIndex = headerIndexMap.get("出发时间");
  const transformedData: any[][] = [];

  for (let i = 1; i < rows.length; i++) {
    const oldRow = rows[i];
    const newRow: any[] = new Array(NEW_HEADERS.length).fill("");

    // 根据映射填充数据
    for (const [newIdx, oldHeader] of Object.entries(INTERNATIONAL_HEADER_MAPPING)) {
      const oldIdx = headerIndexMap.get(oldHeader);
      if (oldIdx !== undefined) {
        newRow[parseInt(newIdx)] = oldRow[oldIdx] ?? "";
      }
    }

    // 类型列固定填写"国际"
    newRow[3] = "国际";

    // 提取出发城市和到达城市（从"航程"列）
    if (routeColIndex !== undefined) {
      const route = oldRow[routeColIndex]?.toString() || "";
      const { from, to } = extractCities(route);
      newRow[4] = from;  // 出发城市
      newRow[5] = to;    // 到达城市
    }

    // 提取出发日期和出发时间（从"出发时间"列）
    if (departTimeColIndex !== undefined) {
      const dateTimeStr = oldRow[departTimeColIndex]?.toString() || "";
      const { date, time } = extractDateTime(dateTimeStr);
      newRow[6] = date;  // 出发日期
      newRow[7] = time;  // 出发时间
    }

    // 格式化金额列为两位小数
    AMOUNT_COLUMNS.forEach(colIdx => {
      newRow[colIdx] = formatAmount(newRow[colIdx]);
    });

    transformedData.push(newRow);
  }

  return transformedData;
};

const readFile = (file: File) => {
  loading.value = true;
  const reader = new FileReader();

  reader.onload = async e => {
    try {
      const buffer = e.target?.result as ArrayBuffer;
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);

      const allData: any[][] = [];
      let domesticCount = 0;
      let internationalCount = 0;

      // 读取"国内机票"工作表
      const domesticSheet = workbook.getWorksheet("国内机票");
      if (domesticSheet) {
        const rows = readWorksheetData(domesticSheet);
        if (rows.length > 1) {
          const headerIndexMap = buildHeaderIndexMap(rows[0]);
          console.log("国内机票表头映射:", Object.fromEntries(headerIndexMap));
          const transformedData = transformDomesticData(rows, headerIndexMap);
          allData.push(...transformedData);
          domesticCount = transformedData.length;
        }
      }

      // 读取"国际机票"工作表
      const internationalSheet = workbook.getWorksheet("国际机票");
      if (internationalSheet) {
        const rows = readWorksheetData(internationalSheet);
        if (rows.length > 1) {
          const headerIndexMap = buildHeaderIndexMap(rows[0]);
          console.log("国际机票表头映射:", Object.fromEntries(headerIndexMap));
          const transformedData = transformInternationalData(rows, headerIndexMap);
          allData.push(...transformedData);
          internationalCount = transformedData.length;
        }
      }

      if (allData.length === 0) {
        ElMessage.error("未找到有效数据，请检查Excel格式");
        loading.value = false;
        return;
      }

      sheetData.value = {
        headers: NEW_HEADERS,
        data: allData
      };

      console.log("读取到表头:", NEW_HEADERS);
      console.log(`国内机票: ${domesticCount} 条, 国际机票: ${internationalCount} 条, 总计: ${allData.length} 条`);

      showData.value = true;
      loading.value = false;
      ElMessage.success(`成功读取文件，国内${domesticCount}条，国际${internationalCount}条，共${allData.length}条数据！`);
    } catch (error) {
      console.error("读取Excel文件失败:", error);
      ElMessage.error("读取Excel文件失败");
      loading.value = false;
    }
  };

  reader.onerror = () => {
    ElMessage.error("文件读取失败");
    loading.value = false;
  };

  reader.readAsArrayBuffer(file);
};

// 生成Excel文件
const generateExcel = async (): Promise<Blob> => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("机票");

  // 添加表头
  worksheet.addRow(NEW_HEADERS);

  // 添加数据
  for (const row of sheetData.value!.data) {
    worksheet.addRow(row);
  }

  // 将金额列转换为数字类型（确保SUM公式能正确计算）
  const numericColumns = [14, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26]; // N, Q, R, S, T, U, V, W, X, Y, Z
  for (let i = 2; i <= sheetData.value!.data.length + 1; i++) {
    const row = worksheet.getRow(i);
    numericColumns.forEach(colIdx => {
      const cell = row.getCell(colIdx);
      const numValue = parseFloat(String(cell.value ?? 0)) || 0;
      cell.value = numValue;
      cell.numFmt = "0.00";
    });
  }

  // 添加合计行
  const dataRowCount = sheetData.value!.data.length;
  const lastDataRow = dataRowCount + 1; // 数据最后一行（表头在第1行，数据从第2行开始）
  const totalRow = worksheet.addRow([]);

  // 第一个单元格填充"合计"
  totalRow.getCell(1).value = "合计";

  // Z列（第26列）使用SUM公式求和
  const totalCell = totalRow.getCell(26);
  totalCell.value = {
    formula: `SUM(Z2:Z${lastDataRow})`
  };
  totalCell.numFmt = "0.00";

  // 设置合计行样式（所有26列都设置边框）
  totalRow.font = { bold: true, size: 10 };
  for (let col = 1; col <= 26; col++) {
    const cell = totalRow.getCell(col);
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" }
    };
  }

  // 设置表头样式
  const headerRow = worksheet.getRow(1);
  headerRow.height = 22;
  headerRow.eachCell(cell => {
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

  // 设置数据行样式
  for (let i = 2; i <= worksheet.rowCount; i++) {
    const row = worksheet.getRow(i);
    row.height = 22;
    row.eachCell(cell => {
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

  // 自动调整列宽
  worksheet.columns.forEach(column => {
    let maxWidth = 8;
    column.eachCell?.({ includeEmpty: true }, cell => {
      const cellValue = cell.value?.toString() || "";
      // 中文字符算2个宽度
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
  return new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
};

// 生成并下载文件
const generateAllFiles = async () => {
  if (!sheetData.value || sheetData.value.data.length === 0) {
    ElMessage.warning("没有可导出的数据");
    return;
  }

  generating.value = true;

  try {
    const blob = await generateExcel();
    saveAs(blob, `星特账单-国内机票.xlsx`);
    ElMessage.success("文件生成成功！");
  } catch (error) {
    console.error("生成文件失败:", error);
    ElMessage.error("生成文件失败");
  } finally {
    generating.value = false;
  }
};
</script>

<template>
  <div class="xinte-bill-split">
    <!-- 上传区域 -->
    <el-card class="upload-card">
      <template #header>
        <div class="card-header">
          <span>上传星特公司账单文件</span>
        </div>
      </template>

      <el-upload
        class="upload-area"
        drag
        :auto-upload="false"
        :show-file-list="false"
        :before-upload="beforeUpload"
        :on-change="handleFileChange"
        accept=".xlsx,.xls"
      >
        <el-icon class="el-icon--upload" :size="60">
          <UploadFilled />
        </el-icon>
        <div class="el-upload__text">
          将Excel文件拖到此处，或<em>点击上传</em>
        </div>
        <template #tip>
          <div class="el-upload__tip">
            支持 .xlsx/.xls 格式，文件大小不超过10MB
          </div>
        </template>
      </el-upload>

      <div v-if="uploadedFile" class="file-info">
        <el-tag type="success">{{ uploadedFile.name }}</el-tag>
      </div>
    </el-card>

    <!-- 加载状态 -->
    <div v-if="loading" class="loading-container">
      <el-icon class="is-loading" :size="40">
        <i class="el-icon-loading" />
      </el-icon>
      <p>正在解析文件...</p>
    </div>

    <!-- 数据预览 -->
    <el-card v-if="showData && sheetData" class="preview-card">
      <template #header>
        <div class="card-header">
          <span>数据预览</span>
          <el-button
            type="primary"
            :loading="generating"
            @click="generateAllFiles"
          >
            {{ generating ? "生成中..." : "生成并下载" }}
          </el-button>
        </div>
      </template>

      <!-- 数据表格预览 -->
      <el-table
        :data="sheetData.data.slice(0, 10)"
        border
        stripe
        max-height="300"
      >
        <el-table-column
          v-for="(header, index) in sheetData.headers"
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
        仅显示前10条数据，共 {{ sheetData.data.length }} 条
      </p>
    </el-card>

    <!-- 无数据提示 -->
    <el-empty v-if="showData && !sheetData" description="未找到有效数据" />
  </div>
</template>

<style scoped>
.xinte-bill-split {
  padding: 20px;
}

.upload-card,
.preview-card {
  margin-bottom: 20px;
}

.card-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
}

.upload-area {
  width: 100%;
}

.file-info {
  margin-top: 15px;
  text-align: center;
}

.loading-container {
  text-align: center;
  padding: 40px;
}

.loading-container p {
  margin-top: 15px;
  color: #909399;
}

.preview-tip {
  color: #909399;
  font-size: 12px;
  margin-top: 10px;
  text-align: right;
}

:deep(.el-upload-dragger) {
  width: 100%;
}
</style>
