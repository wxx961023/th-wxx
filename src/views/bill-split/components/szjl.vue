<script setup lang="ts">
import { ref } from "vue";
import { ElMessage } from "element-plus";
import { UploadFilled } from "@element-plus/icons-vue";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";

defineOptions({
  name: "SzjlBillSplit"
});

const uploadedFile = ref<File | null>(null);
const transformedData = ref<any[][]>([]);
const loading = ref(false);
const showData = ref(false);
const generating = ref(false);

// 新表头
const NEW_HEADERS = [
  "序号",
  "订单号",
  "预订人",
  "乘机人",
  "预订日期",
  "起飞时间",
  "国际/国内",
  "航程",
  "舱等",
  "航班",
  "成交净价",
  "民航基金",
  "燃油税",
  "商旅管理服务费",
  "改签费",
  "退票费",
  "实收实付（含后收）",
  "所属部门",
  "票号"
];

// 国内机票表头映射：新表头索引 -> 旧表头名称
const DOMESTIC_HEADER_MAPPING: Record<number, string> = {
  1: "订单号",
  2: "预订人",
  3: "乘机人",
  4: "记账日期",
  // 5: 起飞时间 = 出发日期 + 出发时间
  // 6: 国际/国内 固定"国内"
  // 7: 航程 = 出发城市 + 到达城市
  8: "舱位类型",
  9: "航班号",
  10: "票面价",
  11: "机建",
  12: "燃油",
  13: "系统使用费",
  14: "改签费",
  15: "退票费",
  16: "总金额",
  17: "费用归属",
  18: "票号"
};

// 国际机票表头映射：新表头索引 -> 旧表头名称
const INTERNATIONAL_HEADER_MAPPING: Record<number, string> = {
  1: "订单号",
  2: "预订人",
  3: "乘机人",
  4: "记账日期",
  // 5: 起飞时间 = 出发时间
  // 6: 国际/国内 固定"国际"
  7: "航程",
  8: "舱位类型",
  9: "航班号",
  10: "票面价",
  11: "税费", // 民航基金 = 税费
  // 12: 燃油税 空
  13: "系统使用费",
  14: "改签费",
  15: "退票费",
  16: "总金额",
  17: "费用归属",
  18: "票号"
};

// 获取上个月的日期范围字符串
const getLastMonthDateRange = (): string => {
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth();

  let lastMonthYear = year;
  let lastMonth = month;

  if (month === 0) {
    lastMonthYear = year - 1;
    lastMonth = 12;
  }

  const daysInLastMonth = new Date(lastMonthYear, lastMonth, 0).getDate();
  const startDate = `${lastMonthYear}${String(lastMonth).padStart(2, "0")}01`;
  const endDate = `${lastMonthYear}${String(lastMonth).padStart(2, "0")}${String(daysInLastMonth).padStart(2, "0")}`;

  return `${startDate}-${endDate}`;
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

// 转换国内机票数据
const transformDomesticData = (rows: any[][], startIndex: number): any[][] => {
  const headerIndexMap = buildHeaderIndexMap(rows[0]);

  const departDateIdx = headerIndexMap.get("出发日期");
  const departTimeIdx = headerIndexMap.get("出发时间");
  const fromCityIdx = headerIndexMap.get("出发城市");
  const toCityIdx = headerIndexMap.get("到达城市");

  const newData: any[][] = [];

  for (let i = 1; i < rows.length; i++) {
    const oldRow = rows[i];
    const newRow: any[] = new Array(NEW_HEADERS.length).fill("");

    // 序号
    newRow[0] = startIndex + i - 1;

    // 根据映射填充数据
    for (const [newIdx, oldHeader] of Object.entries(DOMESTIC_HEADER_MAPPING)) {
      const oldIdx = headerIndexMap.get(oldHeader);
      if (oldIdx !== undefined) {
        newRow[parseInt(newIdx)] = oldRow[oldIdx] ?? "";
      }
    }

    // 起飞时间 = 出发日期 + 出发时间
    const departDate =
      departDateIdx !== undefined ? (oldRow[departDateIdx] ?? "") : "";
    const departTime =
      departTimeIdx !== undefined ? (oldRow[departTimeIdx] ?? "") : "";
    newRow[5] = `${departDate} ${departTime}`.trim();

    // 国际/国内 固定"国内"
    newRow[6] = "国内";

    // 航程 = 出发城市 + 到达城市
    const fromCity =
      fromCityIdx !== undefined ? (oldRow[fromCityIdx] ?? "") : "";
    const toCity = toCityIdx !== undefined ? (oldRow[toCityIdx] ?? "") : "";
    newRow[7] = `${fromCity}-${toCity}`;

    newData.push(newRow);
  }

  return newData;
};

// 转换国际机票数据
const transformIndternationalData = (
  rows: any[][],
  startIndex: number
): any[][] => {
  const headerIndexMap = buildHeaderIndexMap(rows[0]);

  const departTimeIdx = headerIndexMap.get("出发时间");

  const newData: any[][] = [];

  for (let i = 1; i < rows.length; i++) {
    const oldRow = rows[i];
    const newRow: any[] = new Array(NEW_HEADERS.length).fill("");

    // 序号
    newRow[0] = startIndex + i - 1;

    // 根据映射填充数据
    for (const [newIdx, oldHeader] of Object.entries(
      INTERNATIONAL_HEADER_MAPPING
    )) {
      const oldIdx = headerIndexMap.get(oldHeader);
      if (oldIdx !== undefined) {
        newRow[parseInt(newIdx)] = oldRow[oldIdx] ?? "";
      }
    }

    // 起飞时间 = 出发时间
    newRow[5] =
      departTimeIdx !== undefined ? (oldRow[departTimeIdx] ?? "") : "";

    // 国际/国内 固定"国际"
    newRow[6] = "国际";

    // 燃油税保持为空
    newRow[12] = "";

    newData.push(newRow);
  }

  return newData;
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
          const transformedData = transformDomesticData(
            rows,
            allData.length + 1
          );
          allData.push(...transformedData);
          domesticCount = transformedData.length;
        }
      }

      // 读取"国际机票"工作表
      const internationalSheet = workbook.getWorksheet("国际机票");
      if (internationalSheet) {
        const rows = readWorksheetData(internationalSheet);
        if (rows.length > 1) {
          const transformedData = transformIndternationalData(
            rows,
            allData.length + 1
          );
          allData.push(...transformedData);
          internationalCount = transformedData.length;
        }
      }

      if (allData.length === 0) {
        ElMessage.error("未找到有效数据，请检查Excel格式");
        loading.value = false;
        return;
      }

      transformedData.value = allData;

      console.log(
        `国内机票: ${domesticCount} 条, 国际机票: ${internationalCount} 条, 总计: ${allData.length} 条`
      );

      showData.value = true;
      loading.value = false;
      ElMessage.success(
        `成功读取文件，国内${domesticCount}条，国际${internationalCount}条，共${allData.length}条数据！`
      );
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

  // 第1行：垫付机票消费明细报表（居中）
  const row1 = worksheet.addRow(["垫付机票消费明细报表"]);
  worksheet.mergeCells(1, 1, 1, NEW_HEADERS.length);
  row1.getCell(1).font = { bold: true, size: 14 };
  row1.getCell(1).alignment = { horizontal: "center", vertical: "middle" };
  row1.height = 25;
  // 第1行添加边框
  for (let col = 1; col <= NEW_HEADERS.length; col++) {
    row1.getCell(col).border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" }
    };
  }

  // 第2行：日期范围（加粗居中）
  const row2 = worksheet.addRow([getLastMonthDateRange()]);
  worksheet.mergeCells(2, 1, 2, NEW_HEADERS.length);
  row2.getCell(1).font = { bold: true };
  row2.getCell(1).alignment = { horizontal: "center", vertical: "middle" };
  row2.height = 22;
  // 第2行添加边框
  for (let col = 1; col <= NEW_HEADERS.length; col++) {
    row2.getCell(col).border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" }
    };
  }

  // 第3行：公司（加粗靠左）
  const row3 = worksheet.addRow(["公司: 金龙联合汽车工业（苏州）有限公司"]);
  worksheet.mergeCells(3, 1, 3, NEW_HEADERS.length);
  row3.getCell(1).font = { bold: true };
  row3.getCell(1).alignment = { horizontal: "left", vertical: "middle" };
  row3.height = 22;
  // 第3行添加边框
  for (let col = 1; col <= NEW_HEADERS.length; col++) {
    row3.getCell(col).border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" }
    };
  }

  // 第4行：结算币种（加粗靠左）
  const row4 = worksheet.addRow(["结算币种: CNY"]);
  worksheet.mergeCells(4, 1, 4, NEW_HEADERS.length);
  row4.getCell(1).font = { bold: true };
  row4.getCell(1).alignment = { horizontal: "left", vertical: "middle" };
  row4.height = 22;
  // 第4行添加边框
  for (let col = 1; col <= NEW_HEADERS.length; col++) {
    row4.getCell(col).border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" }
    };
  }

  // 第5行：新表头（无底色）
  const headerRow = worksheet.addRow(NEW_HEADERS);
  headerRow.height = 22;
  headerRow.eachCell(cell => {
    cell.font = { bold: true };
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" }
    };
  });

  // 添加数据行
  const amountColumns = [11, 12, 13, 14, 15, 16, 17]; // K、L、M、N、O、P、Q列
  for (const row of transformedData.value) {
    const dataRow = worksheet.addRow(row);
    dataRow.height = 22;
    dataRow.eachCell((cell, colNumber) => {
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.font = { size: 10 };
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" }
      };
      // K、L、M、N、O、P、Q列设置千分位+两位小数，空值或0显示为"-"
      if (amountColumns.includes(colNumber)) {
        const numValue = parseFloat(String(cell.value ?? 0)) || 0;
        cell.value = numValue;
        // 格式：正数千分位两位小数，负数千分位两位小数，零显示"-"，文本显示"-"
        cell.numFmt = '#,##0.00;-#,##0.00;"-";';
      }
    });
  }

  // 添加合计行
  const dataStartRow = 6; // 数据起始行（前5行是标题行）
  const dataEndRow = 5 + transformedData.value.length; // 数据结束行

  const totalRow: any[] = new Array(NEW_HEADERS.length).fill("");
  totalRow[0] = "合计";

  const totalExcelRow = worksheet.addRow(totalRow);
  totalExcelRow.height = 22;

  // K、L、M、N、O、P、Q列（第11-17列）使用SUM公式求和
  const sumColumns = [11, 12, 13, 14, 15, 16, 17];
  const colLetters = ["K", "L", "M", "N", "O", "P", "Q"];

  sumColumns.forEach((colNum, idx) => {
    const cell = totalExcelRow.getCell(colNum);
    cell.value = {
      formula: `SUM(${colLetters[idx]}${dataStartRow}:${colLetters[idx]}${dataEndRow})`
    };
    // 格式：正数千分位两位小数，负数千分位两位小数，零显示"-"
    cell.numFmt = '#,##0.00;-#,##0.00;"-";';
  });

  // 设置合计行样式
  totalExcelRow.eachCell(cell => {
    cell.font = { bold: true, size: 10 };
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" }
    };
  });

  // 设置列宽
  worksheet.columns.forEach((column, index) => {
    const colNum = index + 1;
    if (colNum === 4) {
      column.width = 15; // D列
    } else if (colNum === 5) {
      column.width = 18; // E列
    } else if (colNum === 6) {
      column.width = 20; // F列
    } else if (colNum === 8) {
      column.width = 30; // H列
    } else if (colNum === 9) {
      column.width = 24; // I列
    } else if (colNum === 14 || colNum === 19) {
      column.width = 15; // N, S列
    } else if (colNum === 17) {
      column.width = 20; // Q列
    } else if (colNum === 18) {
      column.width = 24; // R列
    } else {
      column.width = 10; // 其他列
    }
  });

  const buffer = await workbook.xlsx.writeBuffer();
  return new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
};

// 生成并下载文件
const generateAllFiles = async () => {
  if (transformedData.value.length === 0) {
    ElMessage.warning("没有可导出的数据");
    return;
  }

  generating.value = true;

  try {
    const blob = await generateExcel();
    saveAs(blob, `苏州金龙账单.xlsx`);
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
  <div class="szjl-bill-split">
    <!-- 上传区域 -->
    <el-card class="upload-card">
      <template #header>
        <div class="card-header">
          <span>上传苏州金龙账单文件</span>
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
    <el-card v-if="showData && transformedData.length > 0" class="preview-card">
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
        :data="transformedData.slice(0, 10)"
        border
        stripe
        max-height="300"
      >
        <el-table-column
          v-for="(header, index) in NEW_HEADERS"
          :key="index"
          :prop="String(index)"
          :label="header"
          min-width="100"
        >
          <template #default="{ row }">
            {{ row[index] }}
          </template>
        </el-table-column>
      </el-table>
      <p class="preview-tip">
        仅显示前10条数据，共 {{ transformedData.length }} 条
      </p>
    </el-card>

    <!-- 无数据提示 -->
    <el-empty
      v-if="showData && transformedData.length === 0"
      description="未找到有效数据"
    />
  </div>
</template>

<style scoped>
.szjl-bill-split {
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
