<script setup lang="ts">
import { ref, computed } from "vue";
import { ElMessage } from "element-plus";
import { UploadFilled } from "@element-plus/icons-vue";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";

defineOptions({
  name: "YalianBillSplit"
});

// ==================== 账单拆分功能 ====================
const summaryFile = ref<File | null>(null);
const summarySheetData = ref<{
  headers: any[];
  data: any[][];
  domesticCount: number;
  internationalCount: number;
} | null>(null);
const summaryLoading = ref(false);
const summaryGenerating = ref(false);

// ==================== 账单比对功能 ====================
const compareSummaryFile = ref<File | null>(null);
const compareSummaryData = ref<Map<string, { row: any[]; amount: number }>>(new Map());
const compareSummaryLoading = ref(false);

const customerFile = ref<File | null>(null);
const customerData = ref<Map<string, { row: any[]; amount: number }>>(new Map());
const customerLoading = ref(false);

// 比对结果项类型
interface CompareResultItem {
  ticketNo: string;        // 票号
  summaryAmount: number;   // 汇总金额
  customerAmount: number;  // 客户金额
  diff: number;            // 差额
  remark: string;          // 备注
  detail?: any[];          // 详细数据行
}

// 比对结果
const compareResult = ref<CompareResultItem[]>([]);
const showCompareResult = ref(false);

const compareGenerating = ref(false);

// ==================== 公共定义 ====================
// 新表头定义
const NEW_HEADERS = [
  "部门",
  "PNR",
  "票号",
  "乘机人",
  "出票日期",
  "乘机日期",
  "航段",
  "航班号",
  "折扣",
  "票价",
  "机建",
  "税费",
  "保险",
  "退票费",
  "改签费",
  "服务费",
  "应收金额",
  "备注"
];

// 国内机票表头映射
const DOMESTIC_HEADER_MAPPING: Record<number, string> = {
  0: "费用归属",
  1: "订单号",
  2: "票号",
  3: "乘机人",
  4: "记账日期",
  5: "出发日期",
  7: "航班号",
  8: "折扣",
  9: "票面价",
  10: "机建",
  11: "燃油",
  12: "保险费",
  13: "退票费",
  14: "改签费",
  15: "系统使用费",
  16: "总金额"
};

// 国际机票表头映射
const INTERNATIONAL_HEADER_MAPPING: Record<number, string> = {
  0: "费用归属",
  1: "订单号",
  2: "票号",
  3: "乘机人",
  4: "记账日期",
  5: "出发时间",
  7: "航班号",
  9: "票面价",
  11: "税费",
  12: "保险费",
  13: "退票费",
  14: "改签费",
  15: "系统使用费",
  16: "总金额"
};

const AMOUNT_COLUMNS = [9, 10, 11, 12, 13, 14, 15, 16];

// 解析金额值（支持数值和公式格式如 =-1120+735）
const parseAmountValue = (value: any): number => {
  if (value === null || value === undefined) return 0;

  // 如果已经是数值
  if (typeof value === "number") return value;

  const str = value.toString().trim();
  if (!str) return 0;

  // 如果是公式格式（以=开头）
  if (str.startsWith("=")) {
    try {
      // 移除等号，计算公式
      const formula = str.substring(1);
      // 安全计算：只允许数字、加减乘除、小数点、负号
      if (/^[\d+\-*/.\s]+$/.test(formula)) {
        // 使用 Function 安全执行
        const result = new Function(`return ${formula}`)();
        return typeof result === "number" ? result : 0;
      }
    } catch (e) {
      console.warn("公式解析失败:", str, e);
    }
    return 0;
  }

  // 普通数值解析
  return parseFloat(str) || 0;
};

// ==================== 公共方法 ====================
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

const buildHeaderIndexMap = (headers: any[]): Map<string, number> => {
  const map = new Map<string, number>();
  headers.forEach((h, i) => {
    if (h) {
      map.set(h.toString().trim(), i);
    }
  });
  return map;
};

const formatAmount = (value: any): string => {
  const num = parseFloat(value);
  if (isNaN(num)) return "0.00";
  return num.toFixed(2);
};

const extractRouteFromItinerary = (itinerary: string): string => {
  if (!itinerary) return "";
  let firstSegment = itinerary;
  if (itinerary.includes("/")) {
    firstSegment = itinerary.split("/")[0];
  }
  if (firstSegment.includes("-")) {
    const parts = firstSegment.split("-");
    if (parts.length >= 2) {
      return `${parts[0].trim()}+${parts[1].trim()}`;
    }
  }
  return firstSegment.trim();
};

const extractDateFromDateTime = (dateTime: string): string => {
  if (!dateTime) return "";
  let firstPart = dateTime;
  if (dateTime.includes("/")) {
    firstPart = dateTime.split("/")[0];
  }
  firstPart = firstPart.trim();
  const match = firstPart.match(/^(\d{4}-\d{2}-\d{2})/);
  if (match) {
    return match[1];
  }
  return firstPart;
};

const readWorksheetData = (worksheet: ExcelJS.Worksheet): any[][] => {
  const rows: any[][] = [];
  worksheet.eachRow(row => {
    const rowData: any[] = [];
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      // 处理公式单元格：如果有 result 则使用 result，否则使用 value
      if (cell.value && typeof cell.value === "object" && "result" in cell.value) {
        rowData[colNumber - 1] = cell.value.result;
      } else {
        rowData[colNumber - 1] = cell.value;
      }
    });
    rows.push(rowData);
  });
  return rows;
};

// ==================== 账单拆分功能 ====================
const handleSummaryFileChange = (uploadFile: any) => {
  const file = uploadFile.raw;
  if (!file) return;
  summaryFile.value = file;
  readSummaryFileForSplit(file);
};

const readSummaryFileForSplit = (file: File) => {
  summaryLoading.value = true;
  const reader = new FileReader();

  reader.onload = async e => {
    try {
      const buffer = e.target?.result as ArrayBuffer;
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);

      const allData: any[][] = [];
      let domesticCount = 0;
      let internationalCount = 0;

      workbook.eachSheet((ws) => {
        const name = ws.name.trim();
        if (name === "国内机票" || name === "国际机票") {
          const rows = readWorksheetData(ws);
          if (rows.length >= 2) {
            const headerIndexMap = buildHeaderIndexMap(rows[0]);
            const isDomestic = name === "国内机票";
            const mapping = isDomestic ? DOMESTIC_HEADER_MAPPING : INTERNATIONAL_HEADER_MAPPING;
            const departCityIdx = headerIndexMap.get("出发城市");
            const arriveCityIdx = headerIndexMap.get("到达城市");
            const itineraryIdx = headerIndexMap.get("航程");

            for (let i = 1; i < rows.length; i++) {
              const oldRow = rows[i];
              const newRow: any[] = new Array(NEW_HEADERS.length).fill("");

              for (const [newIdx, oldHeader] of Object.entries(mapping)) {
                const oldIdx = headerIndexMap.get(oldHeader);
                if (oldIdx !== undefined) {
                  let value = oldRow[oldIdx] ?? "";
                  if (parseInt(newIdx) === 5 && !isDomestic) {
                    value = extractDateFromDateTime(value?.toString() || "");
                  }
                  newRow[parseInt(newIdx)] = value;
                }
              }

              if (isDomestic && departCityIdx !== undefined && arriveCityIdx !== undefined) {
                const departCity = oldRow[departCityIdx]?.toString() || "";
                const arriveCity = oldRow[arriveCityIdx]?.toString() || "";
                newRow[6] = departCity && arriveCity ? `${departCity}-${arriveCity}` : "";
              } else if (!isDomestic && itineraryIdx !== undefined) {
                const itinerary = oldRow[itineraryIdx]?.toString() || "";
                newRow[6] = extractRouteFromItinerary(itinerary);
              }

              if (!isDomestic) {
                newRow[8] = "";
                newRow[10] = "";
              }

              newRow[17] = isDomestic ? "国内" : "国际";

              AMOUNT_COLUMNS.forEach(colIdx => {
                newRow[colIdx] = formatAmount(newRow[colIdx]);
              });

              allData.push(newRow);
            }

            if (isDomestic) {
              domesticCount = rows.length - 1;
            } else {
              internationalCount = rows.length - 1;
            }
          }
        }
      });

      if (allData.length === 0) {
        ElMessage.error("未找到【国内机票】或【国际机票】工作表，或无有效数据");
        summaryLoading.value = false;
        return;
      }

      summarySheetData.value = {
        headers: NEW_HEADERS,
        data: allData,
        domesticCount,
        internationalCount
      };

      summaryLoading.value = false;
      ElMessage.success(`读取成功，国内${domesticCount}条，国际${internationalCount}条，共${allData.length}条数据`);
    } catch (error) {
      console.error("读取文件失败:", error);
      ElMessage.error("读取文件失败");
      summaryLoading.value = false;
    }
  };

  reader.readAsArrayBuffer(file);
};

// 生成汇总Excel文件
const generateSummary = async () => {
  if (!summarySheetData.value || summarySheetData.value.data.length === 0) {
    ElMessage.warning("没有可导出的数据");
    return;
  }

  summaryGenerating.value = true;

  try {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("汇总");

    worksheet.addRow(NEW_HEADERS);

    for (const row of summarySheetData.value.data) {
      worksheet.addRow(row);
    }

    const numericColumns = [10, 11, 12, 13, 14, 15, 16, 17];
    for (let i = 2; i <= summarySheetData.value.data.length + 1; i++) {
      const row = worksheet.getRow(i);
      numericColumns.forEach(colIdx => {
        const cell = row.getCell(colIdx);
        const cellValue = cell.value;
        if (cellValue !== "" && cellValue !== null && cellValue !== undefined) {
          const numValue = parseFloat(String(cellValue)) || 0;
          cell.value = numValue;
          cell.numFmt = "0.00";
        }
      });
    }

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

    worksheet.columns.forEach(column => {
      let maxWidth = 8;
      column.eachCell?.({ includeEmpty: true }, cell => {
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
    saveAs(blob, `亚连账单汇总.xlsx`);
    ElMessage.success("文件生成成功！");
  } catch (error) {
    console.error("生成文件失败:", error);
    ElMessage.error("生成文件失败");
  } finally {
    summaryGenerating.value = false;
  }
};

// ==================== 账单比对功能 ====================
const handleCompareSummaryChange = (uploadFile: any) => {
  const file = uploadFile.raw;
  if (!file) return;
  compareSummaryFile.value = file;
  readCompareSummaryFile(file);
};

const readCompareSummaryFile = (file: File) => {
  compareSummaryLoading.value = true;
  const reader = new FileReader();

  reader.onload = async e => {
    try {
      const buffer = e.target?.result as ArrayBuffer;
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);

      // 读取"汇总"工作表
      const worksheet = workbook.getWorksheet("汇总");
      if (!worksheet) {
        ElMessage.error("未找到【汇总】工作表");
        compareSummaryLoading.value = false;
        return;
      }

      const rows = readWorksheetData(worksheet);
      if (rows.length < 2) {
        ElMessage.error("汇总工作表中没有有效数据");
        compareSummaryLoading.value = false;
        return;
      }

      const headers = rows[0];
      const headerIndexMap = buildHeaderIndexMap(headers);

      // 查找票号列和应收金额列
      const ticketNoIdx = headerIndexMap.get("票号");
      const amountIdx = headerIndexMap.get("应收金额");

      if (ticketNoIdx === undefined) {
        ElMessage.error("汇总文件中未找到【票号】列");
        compareSummaryLoading.value = false;
        return;
      }

      if (amountIdx === undefined) {
        ElMessage.error("汇总文件中未找到【应收金额】列");
        compareSummaryLoading.value = false;
        return;
      }

      // 构建票号 -> {行数据, 金额} 的映射
      // 同一票号可能有多条记录，需要累加金额
      const dataMap = new Map<string, { row: any[]; amount: number }>();
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        const ticketNo = row[ticketNoIdx]?.toString().trim() || "";
        const amount = parseAmountValue(row[amountIdx]);
        if (ticketNo) {
          if (dataMap.has(ticketNo)) {
            // 累加金额
            const existing = dataMap.get(ticketNo)!;
            existing.amount += amount;
          } else {
            dataMap.set(ticketNo, { row, amount });
          }
        }
      }

      compareSummaryData.value = dataMap;
      compareSummaryLoading.value = false;
      ElMessage.success(`汇总文件读取成功，共 ${rows.length - 1} 行，汇总后 ${dataMap.size} 个票号`);

      if (customerData.value.size > 0) {
        doCompare();
      }
    } catch (error) {
      console.error("读取汇总文件失败:", error);
      ElMessage.error("读取汇总文件失败");
      compareSummaryLoading.value = false;
    }
  };

  reader.readAsArrayBuffer(file);
};

const handleCustomerFileChange = (uploadFile: any) => {
  const file = uploadFile.raw;
  if (!file) return;
  customerFile.value = file;
  readCustomerFile(file);
};

const readCustomerFile = (file: File) => {
  customerLoading.value = true;
  const reader = new FileReader();

  reader.onload = async e => {
    try {
      const buffer = e.target?.result as ArrayBuffer;
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);

      const worksheet = workbook.worksheets[0];
      if (!worksheet) {
        ElMessage.error("客户文件中没有工作表");
        customerLoading.value = false;
        return;
      }

      const rows = readWorksheetData(worksheet);
      if (rows.length < 3) {
        ElMessage.error("客户文件中没有有效数据");
        customerLoading.value = false;
        return;
      }

      // 第一行是标题，第二行是表头
      const headers = rows[1];
      const headerIndexMap = buildHeaderIndexMap(headers);

      // 打印表头信息
      console.log("客户账单表头:", headers);
      console.log("客户账单表头映射:", Object.fromEntries(headerIndexMap));

      const ticketNoIdx = headerIndexMap.get("票号");
      const amountIdx = headerIndexMap.get("金额");

      if (ticketNoIdx === undefined) {
        ElMessage.error("客户文件中未找到【票号】列");
        customerLoading.value = false;
        return;
      }

      if (amountIdx === undefined) {
        ElMessage.error("客户文件中未找到【金额】列");
        customerLoading.value = false;
        return;
      }

      // 构建票号 -> {行数据, 金额} 的映射（从第三行开始是数据）
      // 同一票号可能有多条记录，需要累加金额
      // 跳过"小计"和"总计"行
      const dataMap = new Map<string, { row: any[]; amount: number }>();
      for (let i = 2; i < rows.length; i++) {
        const row = rows[i];
        const ticketNo = row[ticketNoIdx]?.toString().trim() || "";
        const amountRaw = row[amountIdx];
        const amount = parseAmountValue(amountRaw);

        // 跳过小计、总计行
        if (ticketNo === "小计" || ticketNo === "总计" || ticketNo.includes("小计") || ticketNo.includes("总计")) {
          continue;
        }

        if (ticketNo) {
          if (dataMap.has(ticketNo)) {
            // 累加金额
            const existing = dataMap.get(ticketNo)!;
            existing.amount += amount;
          } else {
            dataMap.set(ticketNo, { row, amount });
          }
        }
      }

      console.log("客户数据汇总后:", Array.from(dataMap.entries()).slice(0, 5).map(([k, v]) => ({ 票号: k, 金额: v.amount })));

      customerData.value = dataMap;
      customerLoading.value = false;
      ElMessage.success(`客户文件读取成功，共 ${rows.length - 2} 行，汇总后 ${dataMap.size} 个票号`);

      if (compareSummaryData.value.size > 0) {
        doCompare();
      }
    } catch (error) {
      console.error("读取客户文件失败:", error);
      ElMessage.error("读取客户文件失败");
      customerLoading.value = false;
    }
  };

  reader.readAsArrayBuffer(file);
};

const doCompare = () => {
  const results: CompareResultItem[] = [];

  // 汇总有、客户无
  compareSummaryData.value.forEach((value, ticketNo) => {
    if (!customerData.value.has(ticketNo)) {
      results.push({
        ticketNo,
        summaryAmount: value.amount,
        customerAmount: 0,
        diff: value.amount,
        remark: "汇总有客户无",
        detail: value.row
      });
    } else {
      const customerValue = customerData.value.get(ticketNo)!;
      if (Math.abs(value.amount - customerValue.amount) > 0.01) {
        results.push({
          ticketNo,
          summaryAmount: value.amount,
          customerAmount: customerValue.amount,
          diff: Math.abs(value.amount - customerValue.amount),
          remark: "金额不一致",
          detail: value.row
        });
      }
    }
  });

  // 客户有、汇总无
  customerData.value.forEach((value, ticketNo) => {
    if (!compareSummaryData.value.has(ticketNo)) {
      results.push({
        ticketNo,
        summaryAmount: 0,
        customerAmount: value.amount,
        diff: value.amount,
        remark: "客户有汇总无",
        detail: value.row
      });
    }
  });

  compareResult.value = results;
  showCompareResult.value = true;

  const summaryOnlyCount = results.filter(r => r.remark === "汇总有客户无").length;
  const customerOnlyCount = results.filter(r => r.remark === "客户有汇总无").length;
  const mismatchCount = results.filter(r => r.remark === "金额不一致").length;

  ElMessage.success(`比对完成：汇总独有${summaryOnlyCount}条，客户独有${customerOnlyCount}条，金额不一致${mismatchCount}条`);
};

const canCompare = computed(() => {
  return compareSummaryData.value.size > 0 && customerData.value.size > 0;
});

// 导出比对结果
const exportCompareResult = async () => {
  if (compareResult.value.length === 0) {
    ElMessage.warning("没有可导出的比对结果");
    return;
  }

  compareGenerating.value = true;

  try {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("比对结果");

    // 添加表头
    worksheet.columns = [
      { header: "票号", key: "ticketNo", width: 20 },
      { header: "汇总金额", key: "summaryAmount", width: 15 },
      { header: "客户金额", key: "customerAmount", width: 15 },
      { header: "差额", key: "diff", width: 12 },
      { header: "备注", key: "remark", width: 15 }
    ];

    // 添加数据
    for (const item of compareResult.value) {
      worksheet.addRow({
        ticketNo: item.ticketNo,
        summaryAmount: item.summaryAmount.toFixed(2),
        customerAmount: item.customerAmount.toFixed(2),
        diff: item.diff.toFixed(2),
        remark: item.remark
      });
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

    // 设置数据行样式（根据备注设置不同背景色）
    for (let i = 2; i <= worksheet.rowCount; i++) {
      const row = worksheet.getRow(i);
      const remark = worksheet.getCell(i, 5).value;

      let bgColor = "FFFFFFFF";
      if (remark === "金额不一致") {
        bgColor = "FFFFEBEE"; // 浅红色
      } else if (remark === "汇总有客户无") {
        bgColor = "FFE3F2FD"; // 浅蓝色
      } else if (remark === "客户有汇总无") {
        bgColor = "FFFFF3E0"; // 浅橙色
      }

      row.height = 20;
      row.eachCell(cell => {
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" }
        };
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: bgColor }
        };
      });
    }

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    saveAs(blob, `亚连账单比对结果_${new Date().toLocaleDateString().replace(/\//g, "-")}.xlsx`);
    ElMessage.success("导出成功");
  } catch (error) {
    console.error("导出失败:", error);
    ElMessage.error("导出失败");
  } finally {
    compareGenerating.value = false;
  }
};

const applyStyle = (worksheet: ExcelJS.Worksheet, colCount?: number) => {
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

  worksheet.columns.forEach(column => {
    let maxWidth = 8;
    column.eachCell?.({ includeEmpty: true }, cell => {
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
};
</script>

<template>
  <div class="yalian-bill-split">
    <!-- 第一行：账单拆分功能 -->
    <el-card class="upload-card">
      <template #header>
        <div class="card-header">
          <span>账单拆分</span>
          <el-button
            type="primary"
            :loading="summaryGenerating"
            :disabled="!summarySheetData || summarySheetData.data.length === 0"
            @click="generateSummary"
          >
            {{ summaryGenerating ? "生成中..." : "生成汇总" }}
          </el-button>
        </div>
      </template>

      <el-upload
        class="upload-area"
        drag
        :auto-upload="false"
        :show-file-list="false"
        :before-upload="beforeUpload"
        :on-change="handleSummaryFileChange"
        accept=".xlsx,.xls"
      >
        <el-icon class="el-icon--upload" :size="60">
          <UploadFilled />
        </el-icon>
        <div class="el-upload__text">
          将汇总文件拖到此处，或<em>点击上传</em>
        </div>
        <template #tip>
          <div class="el-upload__tip">
            上传包含【国内机票】和【国际机票】工作表的Excel文件
          </div>
        </template>
      </el-upload>

      <div v-if="summaryFile" class="file-info">
        <el-tag type="success">{{ summaryFile.name }}</el-tag>
      </div>

      <!-- 拆分预览 -->
      <div v-if="summarySheetData" class="preview-section">
        <div class="stats-tags">
          <el-tag type="success">国内: {{ summarySheetData.domesticCount }} 条</el-tag>
          <el-tag type="warning">国际: {{ summarySheetData.internationalCount }} 条</el-tag>
          <el-tag type="primary">总计: {{ summarySheetData.data.length }} 条</el-tag>
        </div>
        <el-table
          :data="summarySheetData.data.slice(0, 5)"
          border
          stripe
          max-height="200"
        >
          <el-table-column
            v-for="(header, index) in summarySheetData.headers"
            :key="index"
            :prop="String(index)"
            :label="header"
            min-width="80"
          >
            <template #default="{ row }">
              {{ row[index] }}
            </template>
          </el-table-column>
        </el-table>
        <p class="preview-tip">仅显示前5条数据</p>
      </div>

      <div v-if="summaryLoading" class="loading-container">
        <el-icon class="is-loading" :size="40">
          <i class="el-icon-loading" />
        </el-icon>
        <p>正在解析文件...</p>
      </div>
    </el-card>

    <!-- 第二行：账单比对功能 -->
    <el-card class="upload-card">
      <template #header>
        <div class="card-header">
          <span>账单比对</span>
          <el-button
            type="primary"
            :disabled="!canCompare"
            @click="doCompare"
          >
            执行比对
          </el-button>
        </div>
      </template>

      <el-row :gutter="20">
        <!-- 汇总文件上传 -->
        <el-col :span="12">
          <div class="upload-section">
            <h4>汇总机票数据</h4>
            <el-upload
              class="upload-area"
              drag
              :auto-upload="false"
              :show-file-list="false"
              :before-upload="beforeUpload"
              :on-change="handleCompareSummaryChange"
              accept=".xlsx,.xls"
            >
              <el-icon class="el-icon--upload" :size="40">
                <UploadFilled />
              </el-icon>
              <div class="el-upload__text">
                拖拽或<em>点击上传</em>
              </div>
            </el-upload>
            <div v-if="compareSummaryFile" class="file-info">
              <el-tag type="success">{{ compareSummaryFile.name }}</el-tag>
              <span class="count-info">共 {{ compareSummaryData.size }} 条</span>
            </div>
          </div>
        </el-col>

        <!-- 客户文件上传 -->
        <el-col :span="12">
          <div class="upload-section">
            <h4>客户账单</h4>
            <el-upload
              class="upload-area"
              drag
              :auto-upload="false"
              :show-file-list="false"
              :before-upload="beforeUpload"
              :on-change="handleCustomerFileChange"
              accept=".xlsx,.xls"
            >
              <el-icon class="el-icon--upload" :size="40">
                <UploadFilled />
              </el-icon>
              <div class="el-upload__text">
                拖拽或<em>点击上传</em>
              </div>
            </el-upload>
            <div v-if="customerFile" class="file-info">
              <el-tag type="success">{{ customerFile.name }}</el-tag>
              <span class="count-info">共 {{ customerData.size }} 条</span>
            </div>
          </div>
        </el-col>
      </el-row>

      <div v-if="compareSummaryLoading || customerLoading" class="loading-container">
        <el-icon class="is-loading" :size="40">
          <i class="el-icon-loading" />
        </el-icon>
        <p>正在解析文件...</p>
      </div>

      <!-- 比对结果 -->
      <div v-if="showCompareResult && compareResult.length > 0" class="result-section">
        <el-row :gutter="20">
          <el-col :span="8">
            <el-statistic title="汇总有、客户无" :value="compareResult.filter(r => r.remark === '汇总有客户无').length">
              <template #suffix>条</template>
            </el-statistic>
          </el-col>
          <el-col :span="8">
            <el-statistic title="客户有、汇总无" :value="compareResult.filter(r => r.remark === '客户有汇总无').length">
              <template #suffix>条</template>
            </el-statistic>
          </el-col>
          <el-col :span="8">
            <el-statistic title="金额不一致" :value="compareResult.filter(r => r.remark === '金额不一致').length">
              <template #suffix>条</template>
            </el-statistic>
          </el-col>
        </el-row>

        <!-- 比对结果表格 -->
        <div class="result-table">
          <div class="table-header">
            <h4>比对结果明细</h4>
            <el-button type="primary" size="small" @click="exportCompareResult">
              导出结果
            </el-button>
          </div>
          <el-table
            :data="compareResult"
            border
            stripe
            max-height="400"
          >
            <el-table-column prop="ticketNo" label="票号" min-width="150" />
            <el-table-column prop="summaryAmount" label="汇总金额" min-width="120">
              <template #default="{ row }">
                {{ row.summaryAmount.toFixed(2) }}
              </template>
            </el-table-column>
            <el-table-column prop="customerAmount" label="客户金额" min-width="120">
              <template #default="{ row }">
                {{ row.customerAmount.toFixed(2) }}
              </template>
            </el-table-column>
            <el-table-column prop="diff" label="差额" min-width="100">
              <template #default="{ row }">
                <span :style="{ color: row.diff > 0 ? '#F56C6C' : '#67C23A' }">
                  {{ row.diff.toFixed(2) }}
                </span>
              </template>
            </el-table-column>
            <el-table-column prop="remark" label="备注" min-width="120">
              <template #default="{ row }">
                <el-tag
                  :type="row.remark === '金额不一致' ? 'danger' : row.remark === '汇总有客户无' ? 'info' : 'warning'"
                >
                  {{ row.remark }}
                </el-tag>
              </template>
            </el-table-column>
          </el-table>
        </div>
      </div>
    </el-card>
  </div>
</template>

<style scoped>
.yalian-bill-split {
  padding: 20px;
}

.upload-card {
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

.upload-section {
  text-align: center;
}

.upload-section h4 {
  margin-bottom: 10px;
  color: #303133;
}

.count-info {
  color: #909399;
  font-size: 12px;
  margin-left: 8px;
}

.loading-container {
  text-align: center;
  padding: 20px;
}

.loading-container p {
  margin-top: 10px;
  color: #909399;
}

.preview-section {
  margin-top: 20px;
}

.stats-tags {
  margin-bottom: 10px;
  display: flex;
  gap: 8px;
}

.preview-tip {
  color: #909399;
  font-size: 12px;
  margin-top: 10px;
  text-align: right;
}

.result-section {
  margin-top: 20px;
  padding-top: 20px;
  border-top: 1px solid #ebeef5;
}

.result-table {
  margin-top: 15px;
}

.result-table h4 {
  margin-bottom: 10px;
  color: #303133;
  font-size: 14px;
}

.table-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 10px;
}

.table-header h4 {
  margin: 0;
  color: #303133;
  font-size: 14px;
}

:deep(.el-upload-dragger) {
  width: 100%;
}
</style>
