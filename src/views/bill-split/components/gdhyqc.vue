<script setup lang="ts">
import { ref } from "vue";
import { ElMessage } from "element-plus";
import { UploadFilled } from "@element-plus/icons-vue";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import JSZip from "jszip";
import { COST_CENTER_TO_INVOICE_UNIT_MAP } from "../costCenterConfig";

defineOptions({
  name: "GdhyqcBillSplit"
});

// 工作表名称
const SHEET_NAME_DOMESTIC = "机票明细(国内)";
const SHEET_NAME_INTERNATIONAL = "机票明细(国际)";
// 充值结算单模板路径
const SETTLEMENT_TEMPLATE_PATH = "/cxjg/充值结算单.xlsx";
// 表头行数（第1-3行为表头）
const HEADER_ROWS = 3;
// 数据起始行（从第4行开始）
const DATA_START_ROW = 4;
// 分组字段
const GROUP_FIELD = "开票单位";
// 应付金额字段名
const PAYABLE_AMOUNT_FIELD = "应付金额";
// 特殊字段映射：国际 -> 国内
const FIELD_MAPPING: Record<string, string> = {
  税费: "机建费"
};

// 成本中心字段名
const COST_CENTER_FIELD = "成本中心";

const uploadedFile = ref<File | null>(null);
const sheetData = ref<{
  headers: any[][]; // 多行表头
  data: any[][];
  groupColIndex: number; // 开票单位列索引
} | null>(null);
// 国内机票应付金额汇总
const domesticPayableAmount = ref<number>(0);
// 国际机票应付金额汇总
const internationalPayableAmount = ref<number>(0);
const loading = ref(false);
const showData = ref(false);
const generating = ref(false);

// 分组结果
interface CompanyGroup {
  companyName: string;
  rows: any[][];
  totalCount: number;
  editableFileName: string;
  groupName?: string;
}
const companyGroups = ref<CompanyGroup[]>([]);

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

// 计算上个月的日期范围
const getLastMonthDateRange = (): { startDate: string; endDate: string } => {
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth(); // 当前月份 (0-11)

  // 上个月的年份和月份
  let lastYear = year;
  let lastMonth = month - 1;
  if (lastMonth < 0) {
    lastMonth = 11;
    lastYear = year - 1;
  }

  // 计算上个月最后一天
  const lastDayOfLastMonth = new Date(lastYear, lastMonth + 1, 0).getDate();

  // 格式化日期
  const monthStr = String(lastMonth + 1).padStart(2, "0");
  const startDate = `${lastYear}-${monthStr}-01`;
  const endDate = `${lastYear}-${monthStr}-${String(lastDayOfLastMonth).padStart(2, "0")}`;

  return { startDate, endDate };
};

// 从工作表读取应付金额列的最后一行值
const readPayableAmountFromSheet = (
  worksheet: ExcelJS.Worksheet
): number => {
  // 读取最后一行表头查找应付金额列
  const headerRow = worksheet.getRow(HEADER_ROWS);
  let payableColIndex = -1;

  headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    const cellValue = cell.value?.toString().trim() || "";
    if (cellValue.includes(PAYABLE_AMOUNT_FIELD)) {
      payableColIndex = colNumber;
    }
  });

  if (payableColIndex === -1) {
    console.log(`未找到"${PAYABLE_AMOUNT_FIELD}"列`);
    return 0;
  }

  // 读取该列的最后一行数据（汇总行）
  const rowCount = worksheet.rowCount;
  const lastRow = worksheet.getRow(rowCount);
  const cell = lastRow.getCell(payableColIndex);
  let value = cell.value;

  // 处理公式结果
  if (value && typeof value === "object" && "result" in value) {
    value = value.result;
  }

  const numValue = Number(value);
  return isNaN(numValue) ? 0 : numValue;
};

// 读取工作表的表头
const readWorksheetHeaders = (
  worksheet: ExcelJS.Worksheet
): { headers: any[][]; lastHeaderRow: any[] } => {
  const headers: any[][] = [];
  for (let i = 1; i <= HEADER_ROWS; i++) {
    const row = worksheet.getRow(i);
    const headerRow: any[] = [];
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      headerRow[colNumber - 1] = cell.value;
    });
    headers.push(headerRow);
  }
  return { headers, lastHeaderRow: headers[HEADER_ROWS - 1] };
};

// 读取工作表的数据行
const readWorksheetData = (worksheet: ExcelJS.Worksheet): any[][] => {
  const data: any[][] = [];
  const rowCount = worksheet.rowCount;
  const colCount = worksheet.columnCount;

  for (let i = DATA_START_ROW; i <= rowCount; i++) {
    const row = worksheet.getRow(i);

    // 使用固定列数读取，确保每行数据长度一致
    const rowData: any[] = [];
    for (let col = 1; col <= colCount; col++) {
      const cell = row.getCell(col);
      rowData.push(cell.value);
    }

    // 检查是否为空行：检查前几个关键列是否都为空
    // 通常第1列（序号）或第2列（订单号）有值才算有效行
    const firstCellValue = rowData[0];
    const secondCellValue = rowData[1];

    const isEmptyRow =
      (firstCellValue === null ||
        firstCellValue === undefined ||
        (typeof firstCellValue === "string" && firstCellValue.trim() === "")) &&
      (secondCellValue === null ||
        secondCellValue === undefined ||
        (typeof secondCellValue === "string" && secondCellValue.trim() === ""));

    // 检查是否为汇总行（通常包含"合计"、"总计"等关键字）
    const isSummaryRow = rowData.some(cell => {
      if (typeof cell === "string") {
        const cellStr = cell.trim();
        return (
          cellStr === "合计" ||
          cellStr === "总计" ||
          cellStr.startsWith("合计") ||
          cellStr.startsWith("总计")
        );
      }
      return false;
    });

    // 只添加非空行且非汇总行
    if (!isEmptyRow && !isSummaryRow) {
      data.push(rowData);
    }
  }
  return data;
};

// 建立国际表头到国内表头的列索引映射
const buildColumnMapping = (
  domesticHeaders: any[],
  internationalHeaders: any[]
): Map<number, number> => {
  const mapping = new Map<number, number>();

  // 建立国内表头字段名到索引的映射
  const domesticFieldIndex = new Map<string, number>();
  for (let i = 0; i < domesticHeaders.length; i++) {
    const fieldName = domesticHeaders[i]?.toString().trim() || "";
    if (fieldName) {
      domesticFieldIndex.set(fieldName, i);
    }
  }

  // 遍历国际表头，建立映射关系
  for (let intlIdx = 0; intlIdx < internationalHeaders.length; intlIdx++) {
    const intlFieldName = internationalHeaders[intlIdx]?.toString().trim() || "";
    if (!intlFieldName) continue;

    // 检查是否有特殊映射
    const mappedFieldName = FIELD_MAPPING[intlFieldName] || intlFieldName;

    // 查找对应的国内列索引
    if (domesticFieldIndex.has(mappedFieldName)) {
      mapping.set(intlIdx, domesticFieldIndex.get(mappedFieldName)!);
    }
  }

  return mapping;
};

// 将国际数据行转换为国内数据格式
const convertInternationalRow = (
  intlRow: any[],
  columnMapping: Map<number, number>,
  domesticColCount: number
): any[] => {
  // 初始化为全0的数组
  const convertedRow: any[] = new Array(domesticColCount).fill(0);

  // 按映射关系填充数据
  for (const [intlIdx, domesticIdx] of columnMapping) {
    if (intlIdx < intlRow.length) {
      convertedRow[domesticIdx] = intlRow[intlIdx];
    }
  }

  return convertedRow;
};

const readFile = (file: File) => {
  loading.value = true;
  const reader = new FileReader();

  reader.onload = async e => {
    try {
      const buffer = e.target?.result as ArrayBuffer;
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);

      // 1. 读取国内工作表（作为基准）
      const domesticSheet = workbook.getWorksheet(SHEET_NAME_DOMESTIC);
      if (!domesticSheet) {
        ElMessage.error(`未找到工作表: ${SHEET_NAME_DOMESTIC}`);
        loading.value = false;
        return;
      }

      const { headers: domesticHeaders, lastHeaderRow: domesticLastHeader } =
        readWorksheetHeaders(domesticSheet);
      const domesticData = readWorksheetData(domesticSheet);

      // 读取国内机票应付金额汇总值
      domesticPayableAmount.value = readPayableAmountFromSheet(domesticSheet);
      console.log(`国内机票应付金额汇总: ${domesticPayableAmount.value}`);

      console.log(`国内工作表读取到 ${domesticData.length} 条数据`);

      // 在国内表头中查找"开票单位"列
      let groupColIndex = -1;
      for (let i = 0; i < domesticLastHeader.length; i++) {
        const cellValue = domesticLastHeader[i]?.toString() || "";
        if (cellValue.includes(GROUP_FIELD)) {
          groupColIndex = i;
          break;
        }
      }

      if (groupColIndex === -1) {
        ElMessage.error(`未找到"${GROUP_FIELD}"列`);
        loading.value = false;
        return;
      }

      console.log(`找到"${GROUP_FIELD}"列，索引: ${groupColIndex}`);

      // 2. 读取国际工作表（如果存在）
      const internationalSheet = workbook.getWorksheet(SHEET_NAME_INTERNATIONAL);
      let allData = [...domesticData];

      if (internationalSheet) {
        const { lastHeaderRow: intlLastHeader } =
          readWorksheetHeaders(internationalSheet);
        const internationalData = readWorksheetData(internationalSheet);

        // 读取国际机票应付金额汇总值
        internationalPayableAmount.value =
          readPayableAmountFromSheet(internationalSheet);
        console.log(`国际机票应付金额汇总: ${internationalPayableAmount.value}`);

        console.log(`国际工作表读取到 ${internationalData.length} 条数据`);

        // 建立列映射关系
        const columnMapping = buildColumnMapping(
          domesticLastHeader,
          intlLastHeader
        );
        console.log("列映射关系:", Object.fromEntries(columnMapping));

        // 转换国际数据并合并
        const domesticColCount = domesticLastHeader.length;
        for (const intlRow of internationalData) {
          const convertedRow = convertInternationalRow(
            intlRow,
            columnMapping,
            domesticColCount
          );
          allData.push(convertedRow);
        }

        console.log(
          `整合后共 ${allData.length} 条数据（国内 ${domesticData.length} + 国际 ${internationalData.length}）`
        );
      } else {
        // 没有国际工作表，国际机票金额为0
        internationalPayableAmount.value = 0;
        console.log(`未找到国际工作表: ${SHEET_NAME_INTERNATIONAL}，仅处理国内数据`);
      }

      // === 根据成本中心映射表转换开票单位 ===
      // 查找"成本中心"列索引
      let costCenterColIndex = -1;
      for (let i = 0; i < domesticLastHeader.length; i++) {
        const cellValue = domesticLastHeader[i]?.toString() || "";
        if (cellValue.includes(COST_CENTER_FIELD)) {
          costCenterColIndex = i;
          break;
        }
      }

      if (costCenterColIndex !== -1) {
        console.log(
          `找到"${COST_CENTER_FIELD}"列，索引: ${costCenterColIndex}`
        );
        console.log(`开始根据成本中心映射表转换开票单位...`);

        let convertedCount = 0;
        let notFoundList: string[] = [];

        for (const row of allData) {
          const costCenter = row[costCenterColIndex]?.toString().trim();
          if (costCenter && COST_CENTER_TO_INVOICE_UNIT_MAP[costCenter]) {
            const oldValue = row[groupColIndex];
            const newValue = COST_CENTER_TO_INVOICE_UNIT_MAP[costCenter];
            row[groupColIndex] = newValue;
            convertedCount++;

            // 只打印前5条转换记录作为示例
            if (convertedCount <= 5) {
              console.log(
                `转换: 成本中心="${costCenter}" -> 开票单位="${newValue}" (原值="${oldValue}")`
              );
            }
          } else if (costCenter && !notFoundList.includes(costCenter)) {
            // 记录未匹配的成本中心（去重）
            notFoundList.push(costCenter);
          }
        }

        console.log(`成本中心映射转换完成，共转换 ${convertedCount} 条记录`);
        if (notFoundList.length > 0) {
          console.log(`未找到映射的成本中心: ${notFoundList.join(", ")}`);
        }
      } else {
        console.log(`未找到"${COST_CENTER_FIELD}"列，跳过映射转换`);
      }

      sheetData.value = {
        headers: domesticHeaders,
        data: allData,
        groupColIndex
      };

      // 处理分组
      processGroups();
      showData.value = true;
      ElMessage.success("文件解析成功！");
    } catch (error) {
      console.error("解析文件失败:", error);
      ElMessage.error("解析文件失败，请检查文件格式");
    } finally {
      loading.value = false;
    }
  };

  reader.onerror = () => {
    ElMessage.error("文件读取失败");
    loading.value = false;
  };

  reader.readAsArrayBuffer(file);
};

// 处理分组逻辑
const processGroups = () => {
  if (!sheetData.value) return;

  const { data, groupColIndex } = sheetData.value;
  const groups = new Map<string, any[][]>();

  // 遍历数据行，按开票单位分组
  for (const row of data) {
    const companyName = row[groupColIndex]?.toString().trim();
    if (!companyName || companyName === "") continue;

    if (!groups.has(companyName)) {
      groups.set(companyName, []);
    }
    groups.get(companyName)!.push(row);
  }

  // 转换为数组
  companyGroups.value = Array.from(groups.entries()).map(
    ([companyName, rows]) => ({
      companyName,
      rows,
      totalCount: rows.length,
      editableFileName: companyName // 文件名使用开票单位名称
    })
  );

  console.log(`共分为 ${companyGroups.value.length} 个开票单位`);
};

// 加载充值结算单模板并填充数据
const loadAndFillSettlementSheet = async (
  workbook: ExcelJS.Workbook
): Promise<void> => {
  try {
    console.log("开始加载充值结算单模板...");

    // 尝试多个可能的路径
    const possiblePaths = [
      "./cxjg/充值结算单.xlsx",
      "/cxjg/充值结算单.xlsx",
      "cxjg/充值结算单.xlsx"
    ];

    let templateBuffer: ArrayBuffer | null = null;
    for (const path of possiblePaths) {
      try {
        const response = await fetch(path);
        if (response.ok) {
          templateBuffer = await response.arrayBuffer();
          console.log(`成功从 ${path} 加载模板`);
          break;
        }
      } catch (e) {
        console.log(`尝试路径 ${path} 失败`);
      }
    }

    if (!templateBuffer) {
      console.error("加载充值结算单模板失败：所有路径都无法访问");
      return;
    }

    if (templateBuffer.byteLength < 1000) {
      console.error("模板文件太小，可能不是有效的Excel文件");
      return;
    }

    const templateWorkbook = new ExcelJS.Workbook();
    await templateWorkbook.xlsx.load(templateBuffer);

    console.log(
      "成功解析充值结算单模板，工作表数量:",
      templateWorkbook.worksheets.length
    );

    const templateSheet = templateWorkbook.getWorksheet("充值结算单");
    if (!templateSheet) {
      console.error("模板中未找到充值结算单工作表");
      return;
    }

    // 创建新工作表
    const settlementSheet = workbook.addWorksheet("充值结算单");

    // 获取上月日期范围
    const { startDate, endDate } = getLastMonthDateRange();

    // Excel 主题颜色映射表（Office 默认主题）
    // theme 6 + tint 0.7999... 对应浅黄色 #FFF3CA
    const themeColorMap: Record<string, string> = {
      "6_0.799920651875362": "FFFFF3CA", // 浅黄色背景
      "6_0.5999938962981048": "FFFDE9A9", // 中黄色
      "6_0.3999755851924192": "FFFCDF7E", // 深黄色
      "1": "FF000000", // 黑色文字
      "0": "FFFFFFFF" // 白色
    };

    // 辅助函数：将主题颜色转换为 ARGB
    const convertThemeColor = (color: any): any => {
      if (!color) return color;

      // 如果已经有 argb 值，直接返回
      if (color.argb) return { argb: color.argb };

      // 如果是主题颜色，尝试转换
      if (color.theme !== undefined) {
        const key =
          color.tint !== undefined
            ? `${color.theme}_${color.tint}`
            : `${color.theme}`;

        if (themeColorMap[key]) {
          return { argb: themeColorMap[key] };
        }

        // 如果没有精确匹配，尝试模糊匹配
        for (const mapKey of Object.keys(themeColorMap)) {
          if (mapKey.startsWith(`${color.theme}_`)) {
            return { argb: themeColorMap[mapKey] };
          }
        }
      }

      // 如果是索引颜色，返回白色作为默认值
      if (color.indexed !== undefined) {
        return { argb: "FFFFFFFF" };
      }

      return color;
    };

    // 辅助函数：转换 fill 样式中的主题颜色
    const convertFillColors = (fill: any): any => {
      if (!fill) return fill;

      const newFill = { ...fill };
      if (newFill.fgColor) {
        newFill.fgColor = convertThemeColor(newFill.fgColor);
      }
      if (newFill.bgColor) {
        newFill.bgColor = convertThemeColor(newFill.bgColor);
      }
      return newFill;
    };

    // 复制模板结构和样式
    templateSheet.eachRow((row, rowNumber) => {
      row.eachCell((cell, colNumber) => {
        const newCell = settlementSheet.getCell(rowNumber, colNumber);

        // 复制原始值，并处理日期范围替换
        let cellValue = cell.value;
        if (typeof cellValue === "string") {
          // 替换结算款项描述中的日期范围
          if (cellValue.includes("结算款项列示如下：")) {
            cellValue = cellValue.replace(
              /本公司\d{4}-\d{2}-\d{2}至\d{4}-\d{2}-\d{2}与贵公司\([^)]+\)的结算款项列示如下：/,
              `本公司${startDate}至${endDate}与贵公司(广东鸿粤汽车销售集团有限公司)的结算款项列示如下：`
            );
          }
        }

        newCell.value = cellValue;

        // 增强样式复制 - 处理主题颜色转换
        if (cell.style) {
          const enhancedStyle: any = {};

          // 复制字体样式，转换主题颜色
          if (cell.style.font) {
            const fontCopy = JSON.parse(JSON.stringify(cell.style.font));
            if (fontCopy.color) {
              fontCopy.color = convertThemeColor(fontCopy.color);
            }
            enhancedStyle.font = fontCopy;
          }

          if (cell.style.alignment) {
            enhancedStyle.alignment = JSON.parse(
              JSON.stringify(cell.style.alignment)
            );
          }

          if (cell.style.border) {
            enhancedStyle.border = JSON.parse(
              JSON.stringify(cell.style.border)
            );
          }

          // 复制填充样式，转换主题颜色为 ARGB
          if (cell.style.fill) {
            enhancedStyle.fill = convertFillColors(
              JSON.parse(JSON.stringify(cell.style.fill))
            );
          }

          if (cell.style.numFmt) {
            enhancedStyle.numFmt = cell.style.numFmt;
          }

          if (cell.style.protection) {
            enhancedStyle.protection = JSON.parse(
              JSON.stringify(cell.style.protection)
            );
          }

          newCell.style = enhancedStyle;

          // 调试：打印 C23 单元格的样式转换
          if (rowNumber === 23 && colNumber === 3) {
            console.log("C23 原始 fill:", JSON.stringify(cell.style.fill));
            console.log("C23 转换后 fill:", JSON.stringify(enhancedStyle.fill));
          }
        }
      });
    });

    // 复制行高
    templateSheet.eachRow((row, rowNumber) => {
      if (row.height) {
        settlementSheet.getRow(rowNumber).height = row.height;
      }
    });

    // 复制列宽
    templateSheet.columns.forEach((column, index) => {
      if (column.width) {
        settlementSheet.getColumn(index + 1).width = column.width;
      }
    });

    // 复制合并单元格
    if (templateSheet.model && templateSheet.model.merges) {
      templateSheet.model.merges.forEach((merge: any) => {
        try {
          settlementSheet.mergeCells(merge);
        } catch (e) {
          // 忽略合并单元格错误
        }
      });
    }

    // === 任务1：隐藏金额为0的数据行 ===
    // 需要检查的行：10-国内酒店, 11-国际酒店, 12-国内火车, 13-国内用车, 14-国内外卖, 15-商务卡
    const rowsToCheck = [10, 11, 12, 13, 14, 15];
    const hiddenRows: number[] = [];

    rowsToCheck.forEach(rowNum => {
      const cellValue = settlementSheet.getCell(`D${rowNum}`).value;
      console.log(`检查第${rowNum}行 D列值:`, cellValue, typeof cellValue);

      // 检查值是否为 0、空、null 或 undefined
      const isEmpty =
        cellValue === null ||
        cellValue === undefined ||
        cellValue === 0 ||
        cellValue === "0" ||
        cellValue === "" ||
        (typeof cellValue === "number" && cellValue === 0);

      if (isEmpty) {
        const row = settlementSheet.getRow(rowNum);
        row.hidden = true; // 使用 hidden 属性隐藏行
        row.height = 0; // 同时设置高度为 0
        hiddenRows.push(rowNum);
        console.log(`隐藏第${rowNum}行`);
      }
    });

    if (hiddenRows.length > 0) {
      console.log("已隐藏金额为0的行:", hiddenRows);
    }

    // === 任务2：动态填充"X月余额"文本（C18单元格）===
    // X = 当前月份的前两个月
    const now = new Date();
    const currentYear = now.getFullYear();
    const currentMonth = now.getMonth() + 1; // 1-12

    let twoMonthsAgoMonth: number;
    if (currentMonth <= 2) {
      // 1月 -> 11月(上一年), 2月 -> 12月(上一年)
      twoMonthsAgoMonth = currentMonth + 10;
    } else {
      twoMonthsAgoMonth = currentMonth - 2;
    }

    const xMonthText = `${twoMonthsAgoMonth}月余额`;
    settlementSheet.getCell("C18").value = xMonthText;
    console.log(`C18 填充: ${xMonthText} (当前${currentMonth}月，前两个月是${twoMonthsAgoMonth}月)`);

    // === 任务3：动态填充"Y月总预存金额"和"截止Y月y日预存款余额"文本 ===
    // Y月 = 当前月份的上一个月，y日 = 上一个月的最后一天
    let lastMonthYear: number;
    let lastMonth: number;

    if (currentMonth === 1) {
      // 当前是1月，上个月是去年的12月
      lastMonthYear = currentYear - 1;
      lastMonth = 12;
    } else {
      lastMonthYear = currentYear;
      lastMonth = currentMonth - 1;
    }

    // 计算上个月的最后一天
    const lastDayOfLastMonth = new Date(lastMonthYear, lastMonth, 0).getDate();

    const yMonthPreText = `${lastMonth}月总预存金额`;
    const yMonthBalanceText = `截止${lastMonth}月${lastDayOfLastMonth}日预存款余额`;

    settlementSheet.getCell("D18").value = yMonthPreText;
    settlementSheet.getCell("F18").value = yMonthBalanceText;

    console.log(`D18 填充: ${yMonthPreText}`);
    console.log(`F18 填充: ${yMonthBalanceText}`);

    // === G18 单元格添加上下边框 ===
    // 边框颜色参考 cxjg.vue: FF95B3D7 (#95b3d7 蓝色)
    const g18Cell = settlementSheet.getCell("G18");
    g18Cell.border = {
      ...g18Cell.border, // 保留原有的左右边框
      top: { style: "thin", color: { argb: "FF95B3D7" } },
      bottom: { style: "thin", color: { argb: "FF95B3D7" } }
    };
    console.log("G18 边框设置完成 (颜色: #95B3D7)");

    // 填充数据：国内机票金额（D8单元格）
    settlementSheet.getCell("D8").value = domesticPayableAmount.value;

    // 填充数据：国际机票金额（D9单元格）
    settlementSheet.getCell("D9").value = internationalPayableAmount.value;

    console.log("充值结算单填充完成:", {
      dateRange: `${startDate} ~ ${endDate}`,
      domesticAmount: domesticPayableAmount.value,
      internationalAmount: internationalPayableAmount.value,
      xMonth: twoMonthsAgoMonth,
      yMonth: lastMonth,
      yDay: lastDayOfLastMonth
    });
  } catch (error) {
    console.error("处理充值结算单时出错:", error);
  }
};

// 需要从汇总表中删除的列名列表
const COLUMNS_TO_REMOVE = [
  "预订人工号",
  "乘机人工号",
  "乘机人工作地",
  "GP检验号",
  "行程单类型",
  "发票号",
  "可抵税额",
  "不可抵应付金额",
  "差标",
  "是否超标",
  "超规原因",
  "改签类型",
  "改签原因",
  "退票类型",
  "退票原因",
  "审批人",
  "出差单",
  "出差事由",
  "备注",
  "项目中心编码",
  "项目中心名称",
  "渠道编码",
  "渠道名称"
];

// 生成包含充值结算单的完整Excel文件
const generateSummaryExcel = async (): Promise<void> => {
  if (!sheetData.value) {
    ElMessage.warning("请先上传文件");
    return;
  }

  generating.value = true;

  try {
    const workbook = new ExcelJS.Workbook();

    // 1. 加载并填充充值结算单
    await loadAndFillSettlementSheet(workbook);

    // 2. 创建机票明细工作表
    const detailSheet = workbook.addWorksheet("机票明细");
    const { headers, data } = sheetData.value;

    // === 删除指定列 ===
    // 深拷贝表头和数据，避免修改原始数据
    const filteredHeaders = headers.map(row => [...row]);
    const filteredData = data.map(row => [...row]);

    // 从最后一行表头中查找要删除的列索引
    const lastHeaderRow = filteredHeaders[filteredHeaders.length - 1];
    const columnsToRemoveIndices: number[] = [];

    for (let i = 0; i < lastHeaderRow.length; i++) {
      const cellValue = lastHeaderRow[i]?.toString().trim() || "";
      if (COLUMNS_TO_REMOVE.includes(cellValue)) {
        columnsToRemoveIndices.push(i);
      }
    }

    // 从后往前排序索引（避免删除时索引变化）
    columnsToRemoveIndices.sort((a, b) => b - a);

    console.log(
      `准备删除 ${columnsToRemoveIndices.length} 列:`,
      columnsToRemoveIndices.map(i => `${lastHeaderRow[i]}(索引${i})`).join(", ")
    );

    // 从表头中删除列
    for (const headerRow of filteredHeaders) {
      for (const colIndex of columnsToRemoveIndices) {
        if (colIndex < headerRow.length) {
          headerRow.splice(colIndex, 1);
        }
      }
    }

    // 从数据行中删除列
    for (const rowData of filteredData) {
      for (const colIndex of columnsToRemoveIndices) {
        if (colIndex < rowData.length) {
          rowData.splice(colIndex, 1);
        }
      }
    }

    console.log(
      `删除列完成，原列数: ${headers[0]?.length || 0}，现列数: ${filteredHeaders[0]?.length || 0}`
    );

    // 添加多行表头
    for (const headerRow of filteredHeaders) {
      const row = detailSheet.addRow(headerRow);
      row.height = 20;
      row.eachCell(cell => {
        cell.font = { bold: true };
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" }
        };
      });
    }

    // === 合并表头中相邻的同名单元格（仅处理第1行和第2行）===
    const headerRowsToMerge = Math.min(2, filteredHeaders.length); // 只处理前两行
    for (let rowIdx = 0; rowIdx < headerRowsToMerge; rowIdx++) {
      const headerRow = filteredHeaders[rowIdx];
      const excelRowNum = rowIdx + 1; // Excel 行号从 1 开始
      const colCount = headerRow.length;

      if (colCount === 0) continue;

      let mergeStartIdx = 0; // 数组索引（从 0 开始）
      let currentValue = headerRow[0]?.toString().trim() || "";

      // 遍历从第 2 列开始（数组索引 1）
      for (let i = 1; i < colCount; i++) {
        const cellValue = headerRow[i]?.toString().trim() || "";

        if (cellValue !== currentValue) {
          // 值变化了，合并之前的单元格 (mergeStartIdx 到 i-1)
          if (i - 1 > mergeStartIdx && currentValue !== "") {
            // Excel 列号 = 数组索引 + 1
            detailSheet.mergeCells(
              excelRowNum,
              mergeStartIdx + 1,
              excelRowNum,
              i
            );
            console.log(
              `表头合并: 第${excelRowNum}行, 列${mergeStartIdx + 1}-${i} (${currentValue})`
            );
          }

          mergeStartIdx = i;
          currentValue = cellValue;
        }
      }

      // 循环结束后，处理最后一组相邻同名单元格
      if (colCount - 1 > mergeStartIdx && currentValue !== "") {
        detailSheet.mergeCells(
          excelRowNum,
          mergeStartIdx + 1,
          excelRowNum,
          colCount
        );
        console.log(
          `表头合并: 第${excelRowNum}行, 列${mergeStartIdx + 1}-${colCount} (${currentValue})`
        );
      }
    }
    console.log("表头合并完成");

    // 添加数据行（过滤空行）
    for (const rowData of filteredData) {
      // 检查是否为空行
      const isEmptyRow = rowData.every((cell: any) => {
        if (cell === null || cell === undefined) return true;
        if (typeof cell === "string" && cell.trim() === "") return true;
        return false;
      });

      // 跳过空行
      if (isEmptyRow) continue;

      const row = detailSheet.addRow(rowData);
      row.height = 20;
      // 使用 includeEmpty: true 确保空单元格也被设置样式
      const colCount = lastHeaderRow.length;
      for (let i = 1; i <= colCount; i++) {
        const cell = row.getCell(i);
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" }
        };
      }
    }

    // === 添加合计行 ===
    // 查找"应付金额"列的索引（复用之前定义的 lastHeaderRow）
    const payableAmountColIndex = lastHeaderRow.findIndex(
      (header: any) => header?.toString().trim() === "应付金额"
    );

    // 创建合计行数据（与数据列数相同）
    const summaryRowData = new Array(lastHeaderRow.length).fill("");
    summaryRowData[0] = "合计";

    const summaryRow = detailSheet.addRow(summaryRowData);

    // 如果找到应付金额列，添加 SUM 公式
    if (payableAmountColIndex !== -1) {
      const dataStartRow = filteredHeaders.length + 1; // 数据开始行（表头行数 + 1）
      const dataEndRow = detailSheet.rowCount - 1; // 数据结束行（当前行数 - 1，不包括合计行）

      // 将列索引转换为 Excel 列字母（支持超过26列）
      const getColumnLetter = (colIndex: number): string => {
        let letter = "";
        let temp = colIndex;
        while (temp >= 0) {
          letter = String.fromCharCode((temp % 26) + 65) + letter;
          temp = Math.floor(temp / 26) - 1;
        }
        return letter;
      };

      const colLetter = getColumnLetter(payableAmountColIndex);
      const sumFormula = `SUM(${colLetter}${dataStartRow}:${colLetter}${dataEndRow})`;

      summaryRow.getCell(payableAmountColIndex + 1).value = {
        formula: sumFormula
      };

      console.log(
        `合计行: 应付金额列索引=${payableAmountColIndex}, 列字母=${colLetter}, 公式=${sumFormula}`
      );
    } else {
      console.log("合计行: 未找到[应付金额]列");
    }

    // 设置合计行样式
    summaryRow.height = 20;
    // 遍历所有列设置样式（包括空单元格）
    for (let i = 1; i <= lastHeaderRow.length; i++) {
      const cell = summaryRow.getCell(i);
      cell.font = { bold: true };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" }
      };
    }

    console.log(`合计行添加完成，位于第 ${detailSheet.rowCount} 行`);

    // 自动调整列宽（优化版：区分中英文字符宽度）
    // 辅助函数：计算字符串的显示宽度
    const calculateStringWidth = (str: string): number => {
      let length = 0;
      for (const char of str) {
        // 中文字符范围判断（包括中文标点）
        if (/[\u4e00-\u9fa5\u3000-\u303f\uff00-\uffef]/.test(char)) {
          length += 2.2; // 中文字符宽度
        } else {
          length += 1; // 英文/数字宽度
        }
      }
      return length;
    };

    // 查找"行程单打印结果"列的索引（在所有表头行中查找）
    let itineraryPrintColIndex = -1;
    let foundInRow = -1;
    for (let rowIdx = 0; rowIdx < filteredHeaders.length; rowIdx++) {
      const headerRow = filteredHeaders[rowIdx];
      for (let i = 0; i < headerRow.length; i++) {
        if (headerRow[i]?.toString().trim() === "行程单打印结果") {
          itineraryPrintColIndex = i;
          foundInRow = rowIdx;
          break;
        }
      }
      if (itineraryPrintColIndex !== -1) break;
    }
    if (itineraryPrintColIndex !== -1) {
      console.log(
        `找到"行程单打印结果"列，索引: ${itineraryPrintColIndex}, 位于表头第${foundInRow + 1}行`
      );
    } else {
      console.log('未找到"行程单打印结果"列');
    }

    console.log("=== 汇总表列宽计算 ===");
    detailSheet.columns.forEach((column, colIndex) => {
      // 特殊处理：行程单打印结果列设置较大的默认宽度（7个中文字符 ≈ 15.4 + 边距）
      let maxLength = colIndex === itineraryPrintColIndex ? 20 : 8;
      column.eachCell?.({ includeEmpty: true }, cell => {
        const cellValue = cell.value?.toString() || "";
        const length = calculateStringWidth(cellValue);
        maxLength = Math.max(maxLength, length);
      });
      const finalWidth = Math.min(maxLength + 2, 35); // 最大宽度限制为35
      column.width = finalWidth;
      // 输出"行程单打印结果"列或前10列的调试信息
      if (colIndex < 10 || colIndex === itineraryPrintColIndex) {
        console.log(
          `列${colIndex + 1}${colIndex === itineraryPrintColIndex ? "(行程单打印结果)" : ""}: 最大字符宽度=${maxLength.toFixed(1)}, 设置宽度=${finalWidth.toFixed(1)}`
        );
      }
    });
    console.log("=== 列宽计算完成 ===");

    // 生成文件并下载
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });

    const { startDate, endDate } = getLastMonthDateRange();
    const fileName = `广东鸿粤汽车结算单_${startDate}_${endDate}.xlsx`;
    saveAs(blob, fileName);

    ElMessage.success("汇总文件生成成功！");
  } catch (error) {
    console.error("生成汇总文件失败:", error);
    ElMessage.error("生成汇总文件失败");
  } finally {
    generating.value = false;
  }
};

// 生成单个Excel文件
const generateExcelForCompany = async (group: CompanyGroup): Promise<Blob> => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("机票明细");

  if (!sheetData.value) {
    throw new Error("数据未加载");
  }

  const { headers } = sheetData.value;

  // 删除原始表头的第一行，只保留第二行及之后的表头
  // 深拷贝以避免修改原始数据
  const filteredHeaders = headers.slice(1).map(row => [...row]);
  const filteredRows = group.rows.map(row => [...row]);
  console.log(
    `拆分表表头: 原始行数=${headers.length}, 删除第一行后=${filteredHeaders.length}`
  );

  // === 删除指定的 37 列 ===
  const SPLIT_COLUMNS_TO_REMOVE = [
    "消费id",
    "订单类型",
    "预订人",
    "预订人工号",
    "预订人部门",
    "乘机人工号",
    "乘机人部门",
    "乘机人工作地",
    "舱等",
    "机票折扣",
    "GP检验号",
    "航班里程",
    "全价票费用",
    "行程单类型",
    "发票号",
    "可抵税额",
    "不可抵应付金额",
    "预订方式",
    "支付方式",
    "支付账户所属公司",
    "差标",
    "是否超标",
    "超规原因",
    "改签类型",
    "改签原因",
    "退票类型",
    "退票原因",
    "审批人",
    "出差单",
    "出差事由",
    "备注",
    "项目中心编码",
    "成本中心编码",
    "成本中心名称",
    "项目中心名称",
    "渠道编码",
    "渠道名称"
  ];

  // 从最后一行表头中查找要删除的列索引
  const splitLastHeaderRow = filteredHeaders[filteredHeaders.length - 1];
  const splitColumnsToRemoveIndices: number[] = [];

  for (let i = 0; i < splitLastHeaderRow.length; i++) {
    const cellValue = splitLastHeaderRow[i]?.toString().trim() || "";
    if (SPLIT_COLUMNS_TO_REMOVE.includes(cellValue)) {
      splitColumnsToRemoveIndices.push(i);
    }
  }

  // 从后往前排序索引（避免删除时索引变化）
  splitColumnsToRemoveIndices.sort((a, b) => b - a);

  console.log(
    `拆分表准备删除 ${splitColumnsToRemoveIndices.length} 列:`,
    splitColumnsToRemoveIndices
      .map(i => `${splitLastHeaderRow[i]}(索引${i})`)
      .join(", ")
  );

  // 从表头中删除列
  for (const headerRow of filteredHeaders) {
    for (const colIndex of splitColumnsToRemoveIndices) {
      if (colIndex < headerRow.length) {
        headerRow.splice(colIndex, 1);
      }
    }
  }

  // 从数据行中删除列
  for (const rowData of filteredRows) {
    for (const colIndex of splitColumnsToRemoveIndices) {
      if (colIndex < rowData.length) {
        rowData.splice(colIndex, 1);
      }
    }
  }

  console.log(
    `拆分表删除列完成，原列数: ${headers[0]?.length || 0}，现列数: ${filteredHeaders[0]?.length || 0}`
  );

  // 计算总列数（删除列后的列数 + 1 序号列）
  const originalColCount = filteredHeaders[0]?.length || 0;
  const totalColCount = originalColCount + 1; // 加上序号列

  // === 第1行：特航航空 ===
  const titleRow = worksheet.addRow(["特航航空"]);
  titleRow.height = 28;
  worksheet.mergeCells(1, 1, 1, totalColCount);
  const titleCell = worksheet.getCell("A1");
  titleCell.font = { bold: true, size: 18 };
  titleCell.alignment = { horizontal: "center", vertical: "middle" };
  titleCell.border = {
    top: { style: "thin" },
    left: { style: "thin" },
    bottom: { style: "thin" },
    right: { style: "thin" }
  };

  // === 第2行：动态日期标题（上一个月）===
  const now = new Date();
  let lastMonthYear = now.getFullYear();
  let lastMonth = now.getMonth(); // 0-11
  if (lastMonth === 0) {
    // 当前是1月，上一个月是去年12月
    lastMonthYear = lastMonthYear - 1;
    lastMonth = 12;
  }
  // 如果 lastMonth 不为 0，getMonth() 返回的就是上个月（因为是 0-based）
  const dateTitle = `${lastMonthYear}年${lastMonth}月份机票对账单`;
  console.log(`拆分表日期标题: ${dateTitle}`);

  const dateRow = worksheet.addRow([dateTitle]);
  dateRow.height = 25;
  worksheet.mergeCells(2, 1, 2, totalColCount);
  const dateCell = worksheet.getCell("A2");
  dateCell.font = { bold: true, size: 16 };
  dateCell.alignment = { horizontal: "center", vertical: "middle" };
  dateCell.border = {
    top: { style: "thin" },
    left: { style: "thin" },
    bottom: { style: "thin" },
    right: { style: "thin" }
  };

  // === 第3行（或第3-4行）：表头（带序号列）===
  const headerRowCount = filteredHeaders.length;
  for (let i = 0; i < headerRowCount; i++) {
    const headerRow = filteredHeaders[i];
    // 在表头前面添加"序号"占位符
    const newHeaderRow = ["序号", ...headerRow];
    const row = worksheet.addRow(newHeaderRow);
    row.height = 20;
    // 设置所有单元格样式
    for (let col = 1; col <= totalColCount; col++) {
      const cell = row.getCell(col);
      cell.font = { bold: true };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" }
      };
    }
  }

  // 合并序号列的表头单元格（如果表头有多行则合并）
  if (headerRowCount >= 2) {
    worksheet.mergeCells(3, 1, 3 + headerRowCount - 1, 1);
    console.log(`序号列表头合并: A3:A${3 + headerRowCount - 1}`);
  }

  // === 合并第三行表头中相邻的同名单元格（排除序号列）===
  const thirdRowNum = 3; // 第三行是表头的第一行
  const thirdRowData = filteredHeaders[0]; // 表头第一行数据（不含序号）

  if (thirdRowData && thirdRowData.length > 0) {
    // 从第2列开始（跳过序号列），数组索引从0开始对应Excel第2列
    let mergeStartCol = 2; // Excel 列号（从2开始，跳过序号列）
    let currentValue = thirdRowData[0]?.toString().trim() || "";

    for (let i = 1; i < thirdRowData.length; i++) {
      const cellValue = thirdRowData[i]?.toString().trim() || "";
      const excelCol = i + 2; // Excel 列号 = 数组索引 + 2（因为第1列是序号）

      if (cellValue !== currentValue) {
        // 值变化了，合并之前的单元格
        if (excelCol - 1 > mergeStartCol && currentValue !== "") {
          worksheet.mergeCells(thirdRowNum, mergeStartCol, thirdRowNum, excelCol - 1);
          console.log(
            `拆分表第三行表头合并: 列${mergeStartCol}-${excelCol - 1} (${currentValue})`
          );
        }
        mergeStartCol = excelCol;
        currentValue = cellValue;
      }
    }

    // 循环结束后，处理最后一组相邻同名单元格
    const lastExcelCol = thirdRowData.length + 1; // 最后一列的 Excel 列号
    if (lastExcelCol > mergeStartCol && currentValue !== "") {
      worksheet.mergeCells(thirdRowNum, mergeStartCol, thirdRowNum, lastExcelCol);
      console.log(
        `拆分表第三行表头合并: 列${mergeStartCol}-${lastExcelCol} (${currentValue})`
      );
    }
    console.log("拆分表第三行表头合并完成");
  }

  // === 添加数据行（带序号）===
  const dataStartRow = 3 + headerRowCount; // 数据开始行号
  let serialNumber = 1;
  for (const rowData of filteredRows) {
    // 在数据前面添加序号
    const newRowData = [serialNumber, ...rowData];
    const row = worksheet.addRow(newRowData);
    row.height = 20;
    // 设置所有单元格样式
    for (let col = 1; col <= totalColCount; col++) {
      const cell = row.getCell(col);
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" }
      };
    }
    serialNumber++;
  }

  console.log(
    `拆分表生成: ${group.groupName}, 数据行数: ${filteredRows.length}, 数据起始行: ${dataStartRow}`
  );

  // === 添加合计行 ===
  // 1. 查找"应付金额"列的索引
  const lastHeaderRow = filteredHeaders[filteredHeaders.length - 1];
  let payableColIndex = -1;
  for (let i = 0; i < lastHeaderRow.length; i++) {
    if (lastHeaderRow[i]?.toString().trim() === "应付金额") {
      payableColIndex = i + 2; // +1 是因为序号列，+1 是因为 Excel 列号从 1 开始
      break;
    }
  }

  // 记录数据最后一行（合计行之前）
  const dataLastRow = worksheet.rowCount;

  if (payableColIndex !== -1) {
    const totalRow = worksheet.addRow([]);
    totalRow.height = 20;

    // 设置所有单元格的边框
    for (let col = 1; col <= totalColCount; col++) {
      const cell = totalRow.getCell(col);
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" }
      };
      cell.alignment = { horizontal: "center", vertical: "middle" };
    }

    // 第1列显示"合计"
    const totalLabelCell = totalRow.getCell(1);
    totalLabelCell.value = "合计";
    totalLabelCell.font = { bold: true };

    // "应付金额"列使用 SUM 公式
    const payableColLetter = worksheet.getColumn(payableColIndex).letter;
    const totalAmountCell = totalRow.getCell(payableColIndex);
    totalAmountCell.value = {
      formula: `SUM(${payableColLetter}${dataStartRow}:${payableColLetter}${dataLastRow})`
    };
    totalAmountCell.font = { bold: true };

    console.log(
      `拆分表合计行: 应付金额列索引=${payableColIndex}(${payableColLetter}列), 数据范围=${dataStartRow}-${dataLastRow}, 合计行=${worksheet.rowCount}`
    );
  } else {
    console.log('拆分表合计行: 未找到"应付金额"列');
  }

  // === 添加提示行（合计行之后）===
  // 生成日期：当前年月 + 25日
  const tipDate = new Date();
  const tipYear = tipDate.getFullYear();
  const tipMonth = tipDate.getMonth() + 1; // 1-12
  const tipDateStr = `${tipYear}年${tipMonth}月25日`;
  const tipText = `请贵司在${tipDateStr}前结款，付款后请提供银行水单查询款项是否到账，谢谢合作！`;

  const tipRow = worksheet.addRow([tipText]);
  tipRow.height = 25;

  // 合并该行所有单元格
  const tipRowNum = worksheet.rowCount;
  worksheet.mergeCells(tipRowNum, 1, tipRowNum, totalColCount);

  // 设置提示行样式：红色字体、靠左对齐
  const tipCell = tipRow.getCell(1);
  tipCell.font = { color: { argb: "FFFF0000" }, bold: false }; // 红色字体
  tipCell.alignment = { horizontal: "left", vertical: "middle" }; // 靠左对齐
  tipCell.border = {
    top: { style: "thin" },
    left: { style: "thin" },
    bottom: { style: "thin" },
    right: { style: "thin" }
  };

  console.log(`拆分表提示行: ${tipText}, 位于第${tipRowNum}行`);

  // === 添加收款信息行（提示行之后）===
  const paymentInfo = `收款账户名称：深圳市特航航空服务有限公司\n收款账户：38980188000607612\n开户银行：光大银行深圳八卦岭支行`;

  const paymentRow = worksheet.addRow([paymentInfo]);
  paymentRow.height = 65; // 行高设置为65，足够显示三行文本

  // 合并前4个单元格（A列到D列）
  const paymentRowNum = worksheet.rowCount;
  worksheet.mergeCells(paymentRowNum, 1, paymentRowNum, 4);

  // 设置收款信息单元格样式：靠左对齐、垂直居中、自动换行、无边框
  const paymentCell = paymentRow.getCell(1);
  paymentCell.font = { bold: true, size: 12 }; // 加粗、12号字体
  paymentCell.alignment = {
    horizontal: "left",
    vertical: "middle",
    wrapText: true // 启用自动换行
  };
  // 不设置边框（无边框）

  console.log(`拆分表收款信息行: 位于第${paymentRowNum}行，合并A-D列`);

  // 自动调整列宽（只基于数据行，从 dataStartRow 到 dataLastRow）
  const dataEndRow = dataLastRow; // 使用数据最后一行（不包括合计行）

  // 辅助函数：计算字符串的显示宽度
  const calculateStringWidth = (str: string): number => {
    let length = 0;
    for (const char of str) {
      if (/[\u4e00-\u9fa5\u3000-\u303f\uff00-\uffef]/.test(char)) {
        length += 2.2; // 中文字符宽度
      } else {
        length += 1; // 英文/数字宽度
      }
    }
    return length;
  };

  console.log(`拆分表列宽计算（基于数据行${dataStartRow}-${dataEndRow}）:`);
  worksheet.columns.forEach((column, colIndex) => {
    // 获取该列的表头名称（序号列为"序号"，其他列从 filteredHeaders 获取）
    let colHeaderName = "";
    if (colIndex === 0) {
      colHeaderName = "序号";
    } else {
      // colIndex - 1 对应 filteredHeaders 中的索引（因为第0列是序号列）
      colHeaderName =
        filteredHeaders[filteredHeaders.length - 1]?.[colIndex - 1]
          ?.toString()
          .trim() || "";
    }

    // 根据列名设置不同的最小宽度
    let minWidth = 8; // 默认最小宽度
    if (colIndex === 0) {
      minWidth = 6; // 序号列
    } else if (colHeaderName === "行程单打印结果") {
      minWidth = 16; // 行程单打印结果列（7个中文字符 ≈ 15.4）
    }

    let maxLength = minWidth;

    // 只遍历数据行（从 dataStartRow 开始）
    for (let rowNum = dataStartRow; rowNum <= dataEndRow; rowNum++) {
      const cell = worksheet.getCell(rowNum, colIndex + 1);
      const cellValue = cell.value?.toString() || "";
      const length = calculateStringWidth(cellValue);
      maxLength = Math.max(maxLength, length);
    }

    const finalWidth = Math.min(maxLength + 2, 40); // 最大宽度限制为40
    column.width = finalWidth;

    // 输出所有列的调试信息
    console.log(
      `列${colIndex + 1}(${colHeaderName}): 最大字符宽度=${maxLength.toFixed(1)}, 设置宽度=${finalWidth.toFixed(1)}`
    );
  });
  console.log(`拆分表列宽计算完成，总列数: ${totalColCount}`);

  const buffer = await workbook.xlsx.writeBuffer();
  return new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
};

// 生成所有文件
const generateAllFiles = async () => {
  if (companyGroups.value.length === 0) {
    ElMessage.warning("没有可生成的数据");
    return;
  }

  generating.value = true;

  try {
    if (companyGroups.value.length === 1) {
      // 只有一个开票单位，直接下载单个文件
      const group = companyGroups.value[0];
      const blob = await generateExcelForCompany(group);
      saveAs(blob, `${group.editableFileName}.xlsx`);
      ElMessage.success("文件生成成功！");
    } else {
      // 多个开票单位，打包成ZIP
      const zip = new JSZip();

      for (const group of companyGroups.value) {
        const blob = await generateExcelForCompany(group);
        zip.file(`${group.editableFileName}.xlsx`, blob);
      }

      const zipBlob = await zip.generateAsync({ type: "blob" });
      saveAs(zipBlob, "广东鸿粤汽车账单拆分.zip");
      ElMessage.success(`成功生成 ${companyGroups.value.length} 个文件！`);
    }
  } catch (error) {
    console.error("生成文件失败:", error);
    ElMessage.error("生成文件失败");
  } finally {
    generating.value = false;
  }
};

// 更新文件名
const updateFileName = (index: number, newName: string) => {
  companyGroups.value[index].editableFileName = newName;
};
</script>

<template>
  <div class="gdhyqc-bill-split">
    <!-- 上传区域 -->
    <el-card class="upload-card">
      <template #header>
        <div class="card-header">
          <span>上传账单文件</span>
        </div>
      </template>

      <el-upload
        class="upload-area"
        drag
        :auto-upload="false"
        :show-file-list="false"
        :before-upload="beforeUpload"
        :on-change="handleFileChange"
        accept=".xlsx"
      >
        <el-icon class="el-icon--upload" :size="60">
          <UploadFilled />
        </el-icon>
        <div class="el-upload__text">
          将Excel文件拖到此处，或<em>点击上传</em>
        </div>
        <template #tip>
          <div class="el-upload__tip">
            支持 .xlsx 格式，文件大小不超过10MB<br />
            将按照"开票单位"字段进行分组拆分
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

    <!-- 汇总信息 -->
    <el-card v-if="showData" class="summary-card">
      <template #header>
        <div class="card-header">
          <span>汇总信息</span>
          <el-button
            type="success"
            :loading="generating"
            @click="generateSummaryExcel"
          >
            {{ generating ? "生成中..." : "生成汇总文件" }}
          </el-button>
        </div>
      </template>

      <el-descriptions :column="2" border>
        <el-descriptions-item label="国内机票应付金额">
          <el-tag type="primary">{{ domesticPayableAmount.toFixed(2) }} 元</el-tag>
        </el-descriptions-item>
        <el-descriptions-item label="国际机票应付金额">
          <el-tag type="success">{{ internationalPayableAmount.toFixed(2) }} 元</el-tag>
        </el-descriptions-item>
        <el-descriptions-item label="合计金额">
          <el-tag type="danger">
            {{ (domesticPayableAmount + internationalPayableAmount).toFixed(2) }} 元
          </el-tag>
        </el-descriptions-item>
        <el-descriptions-item label="数据条数">
          <el-tag>{{ sheetData?.data.length || 0 }} 条</el-tag>
        </el-descriptions-item>
      </el-descriptions>
    </el-card>

    <!-- 分组结果 -->
    <el-card v-if="showData && companyGroups.length > 0" class="result-card">
      <template #header>
        <div class="card-header">
          <span>分组结果（共 {{ companyGroups.length }} 个开票单位）</span>
          <el-button
            type="primary"
            :loading="generating"
            @click="generateAllFiles"
          >
            {{ generating ? "生成中..." : "生成拆分文件" }}
          </el-button>
        </div>
      </template>

      <el-table :data="companyGroups" border stripe>
        <el-table-column prop="companyName" label="开票单位" width="250" />
        <el-table-column prop="totalCount" label="数据条数" width="120" />
        <el-table-column label="文件名">
          <template #default="{ row, $index }">
            <el-input
              v-model="row.editableFileName"
              @change="updateFileName($index, $event)"
            >
              <template #append>.xlsx</template>
            </el-input>
          </template>
        </el-table-column>
      </el-table>
    </el-card>

    <!-- 无数据提示 -->
    <el-empty
      v-if="showData && companyGroups.length === 0"
      description="未找到有效数据"
    />
  </div>
</template>

<style scoped>
.gdhyqc-bill-split {
  padding: 20px;
}

.upload-card,
.summary-card,
.result-card {
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

:deep(.el-upload-dragger) {
  width: 100%;
}
</style>
