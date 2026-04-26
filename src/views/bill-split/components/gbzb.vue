<script setup lang="ts">
import { ref } from "vue";
import { ElMessage } from "element-plus";
import { UploadFilled } from "@element-plus/icons-vue";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";

defineOptions({
  name: "GbzbBillSplit"
});

const uploadedFile = ref<File | null>(null);
const sheetData = ref<{
  headers: any[];
  data: any[][];
  personColIndex: number;
} | null>(null);
const loading = ref(false);
const showData = ref(false);
const generating = ref(false);

// 指定人员名单（第一组，按此顺序排序）
const GROUP_A_PERSONS = ["付兴", "余泉", "周静", "韩东生", "姚国华", "盘国辉"];

// 表头映射：新表头 -> 旧表头（null表示不需要数据，空字符串表示空列）
const HEADER_MAPPING: { newHeader: string; oldHeader: string | null }[] = [
  { newHeader: "姓名", oldHeader: "乘机人" },
  { newHeader: "票号", oldHeader: "票号" },
  { newHeader: "起飞时间", oldHeader: "出发日期" },
  { newHeader: "航程名称", oldHeader: "行程" },
  { newHeader: "航班号", oldHeader: "航班号" },
  { newHeader: "航司名称", oldHeader: "航空公司" },
  { newHeader: "订单状态", oldHeader: "订单状态" },
  { newHeader: "票面价/改签补差", oldHeader: "票面价" },
  { newHeader: "销售价不含税金额", oldHeader: null },
  { newHeader: "销售价可抵扣税额", oldHeader: null },
  { newHeader: "机建", oldHeader: "机建" },
  { newHeader: "燃油费", oldHeader: "燃油" },
  { newHeader: "燃油费不含税金额", oldHeader: null },
  { newHeader: "燃油费可抵扣税额", oldHeader: null },
  { newHeader: "改签手续费", oldHeader: "改签费" },
  { newHeader: "行程单金额", oldHeader: "机票费" },
  { newHeader: "退票手续费", oldHeader: "退票费" },
  { newHeader: "服务费", oldHeader: "系统使用费" },
  { newHeader: "服务费不含税金额", oldHeader: null },
  { newHeader: "服务费可抵税金额", oldHeader: null },
  { newHeader: "结算金额", oldHeader: "总金额" },
  { newHeader: "不含税总额", oldHeader: null },
  { newHeader: "可抵税总额", oldHeader: null },
  { newHeader: "备注", oldHeader: null }
];

// 新表头列表
const NEW_HEADERS = HEADER_MAPPING.map(h => h.newHeader);

// 排序后的数据
const sortedData = ref<any[][]>([]);
// 分组数据（用于生成小计）
const groupAData = ref<any[][]>([]);
const groupBData = ref<any[][]>([]);
// 统计信息
const statsInfo = ref<{
  groupACount: number;
  groupBCount: number;
  groupAPersons: Map<string, number>;
} | null>(null);

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

const readFile = (file: File) => {
  loading.value = true;
  const reader = new FileReader();

  reader.onload = async e => {
    try {
      const buffer = e.target?.result as ArrayBuffer;
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);

      // 读取"国内机票"工作表
      const worksheet = workbook.getWorksheet("国内机票");
      if (!worksheet) {
        ElMessage.error("未找到'国内机票'工作表，请检查Excel格式");
        loading.value = false;
        return;
      }

      const rows: any[][] = [];
      worksheet.eachRow(row => {
        const rowData: any[] = [];
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          rowData[colNumber - 1] = cell.value;
        });
        rows.push(rowData);
      });

      if (rows.length === 0) {
        ElMessage.error("Excel文件中没有数据");
        loading.value = false;
        return;
      }

      // 第一行是表头
      const oldHeaders = rows[0];

      // 构建旧表头到列索引的映射
      const oldHeaderIndexMap = new Map<string, number>();
      oldHeaders.forEach((h, i) => {
        if (h) {
          oldHeaderIndexMap.set(h.toString().trim(), i);
        }
      });

      // 检测第二行是否为子表头行（包含"票面价/改签补差"等子表头关键词）
      const subHeaderKeywords = [
        "票面价/改签补差",
        "销售价不含税金额",
        "燃油费不含税金额",
        "服务费不含税金额"
      ];
      let hasSubHeaderRow = false;
      if (rows.length > 1) {
        const secondRowValues = rows[1].map((v: any) =>
          v ? v.toString().trim() : ""
        );
        hasSubHeaderRow = subHeaderKeywords.some(kw =>
          secondRowValues.includes(kw)
        );
      }

      // 数据起始行索引：有子表头行则从第3行开始，否则从第2行开始
      const dataStartIndex = hasSubHeaderRow ? 2 : 1;

      console.log("旧表头索引映射:", Object.fromEntries(oldHeaderIndexMap));
      console.log("是否有子表头行:", hasSubHeaderRow, "数据起始行索引:", dataStartIndex);

      // 查找"乘机人"列索引（用于分组）
      const personColIndex = oldHeaderIndexMap.get("乘机人");
      if (personColIndex === undefined) {
        ElMessage.error("未找到'乘机人'列，请检查Excel格式");
        loading.value = false;
        return;
      }

      // 将旧数据转换为新表格式
      const transformedData: any[][] = [];
      // 获取票面价和改签费的列索引（用于H列计算）
      const priceColIndex = oldHeaderIndexMap.get("票面价");
      const oldChangeFeeColIndex = oldHeaderIndexMap.get("改签费");

      // 数据从检测到的起始行开始
      for (let i = dataStartIndex; i < rows.length; i++) {
        const oldRow = rows[i];
        const newRow: any[] = [];

        for (const mapping of HEADER_MAPPING) {
          if (mapping.oldHeader === null) {
            // null 表示该列为空
            newRow.push("");
          } else {
            // 根据旧表头获取数据
            const oldIndex = oldHeaderIndexMap.get(mapping.oldHeader);
            if (oldIndex !== undefined) {
              newRow.push(oldRow[oldIndex] ?? "");
            } else {
              newRow.push("");
            }
          }
        }

        // H列（索引7）= 票面价 + 改签费
        if (priceColIndex !== undefined || oldChangeFeeColIndex !== undefined) {
          const price = parseFloat(oldRow[priceColIndex]) || 0;
          const changeFee = parseFloat(oldRow[oldChangeFeeColIndex]) || 0;
          newRow[7] = (price + changeFee).toFixed(2);
        }

        // P列（索引15）行程单金额 使用公式标记，在导出时生成公式
        // 公式：=H+K+L+O（票面价/改签补差 + 机建 + 燃油费 + 改签手续费）
        newRow[15] = "__FORMULA_ITINERARY__";

        // O列（索引14）改签手续费归零（已合并到H列）
        newRow[14] = "0.00";

        // K列（索引10）和L列（索引11）保留两位小数
        const kVal = parseFloat(newRow[10]) || 0;
        newRow[10] = kVal.toFixed(2);
        const lVal = parseFloat(newRow[11]) || 0;
        newRow[11] = lVal.toFixed(2);

        // Q列（索引16）、R列（索引17）、U列（索引20）保留两位小数
        const qVal = parseFloat(newRow[16]) || 0;
        newRow[16] = qVal.toFixed(2);
        const rVal = parseFloat(newRow[17]) || 0;
        newRow[17] = rVal.toFixed(2);
        const uVal = parseFloat(newRow[20]) || 0;
        newRow[20] = uVal.toFixed(2);

        transformedData.push(newRow);
      }

      sheetData.value = {
        headers: NEW_HEADERS,
        data: transformedData,
        personColIndex: 0 // 转换后，姓名列（原乘机人）在第0列
      };

      console.log("读取到表头:", NEW_HEADERS);

      // 自动执行分组
      autoGroupData();

      showData.value = true;
      loading.value = false;
      ElMessage.success(`成功读取文件，共 ${rows.length - 1} 条数据！`);
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

// 自动分组并排序
const autoGroupData = () => {
  if (!sheetData.value) return;

  const personColIndex = sheetData.value.personColIndex;

  // 按人员分组
  const groupAMap = new Map<string, any[][]>(); // 指定人员组
  const groupBRows: any[][] = []; // 其他人员组

  for (const row of sheetData.value.data) {
    const personName = row[personColIndex]?.toString().trim() || "";
    if (GROUP_A_PERSONS.includes(personName)) {
      if (!groupAMap.has(personName)) {
        groupAMap.set(personName, []);
      }
      groupAMap.get(personName)!.push(row);
    } else {
      groupBRows.push(row);
    }
  }

  // 按指定顺序排列：先是指定人员（按GROUP_A_PERSONS顺序），然后是其他人员
  sortedData.value = [];
  groupAData.value = [];
  groupBData.value = [];

  // 统计每个指定人员的数据条数
  const groupAPersons = new Map<string, number>();
  let groupACount = 0;

  for (const personName of GROUP_A_PERSONS) {
    const rows = groupAMap.get(personName) || [];
    if (rows.length > 0) {
      sortedData.value.push(...rows);
      groupAData.value.push(...rows);
      groupAPersons.set(personName, rows.length);
      groupACount += rows.length;
    }
  }

  // 添加其他人员数据
  sortedData.value.push(...groupBRows);
  groupBData.value = groupBRows;

  statsInfo.value = {
    groupACount,
    groupBCount: groupBRows.length,
    groupAPersons
  };

  console.log("排序结果:");
  console.log(`- 指定人员组: ${groupACount} 条`);
  GROUP_A_PERSONS.forEach(name => {
    const count = groupAPersons.get(name) || 0;
    if (count > 0) {
      console.log(`  - ${name}: ${count} 条`);
    }
  });
  console.log(`- 其他人员组: ${groupBRows.length} 条`);
  console.log(`- 总计: ${sortedData.value.length} 条`);
};

// 获取上一个月的年月字符串（用于标题）
const getLastMonthStr = () => {
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth(); // 0-11，getMonth()返回的月份减1就是上个月
  // 如果当前是1月，上个月是去年的12月
  if (month === 0) {
    return `${year - 1}年12月`;
  }
  return `${year}年${month}月`;
};

// 生成Excel文件
const generateExcel = async (): Promise<Blob> => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("国内机票");

  // 添加标题行
  const titleText = `深圳国宝造币有限公司${getLastMonthStr()}订单汇总账单-国内机票`;
  const titleRow = worksheet.addRow([titleText]);
  // 合并标题行单元格（合并所有列）
  worksheet.mergeCells(1, 1, 1, NEW_HEADERS.length);
  // 设置标题行样式
  titleRow.eachCell(cell => {
    cell.font = { bold: true, size: 14 };
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFCC99" }
    };
  });
  titleRow.height = 25;

  // 添加表头
  worksheet.addRow(NEW_HEADERS);

  // 需要计算小计的列索引（数字列）
  const sumColumns = [
    7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22
  ];

  // 需要转换为数字的列索引（这些列在读取时被存储为字符串）
  const numericColumns = [7, 10, 11, 14, 15, 16, 17, 20]; // H, K, L, O, P, Q, R, U

  // 添加指定人员组数据
  for (const row of groupAData.value) {
    const excelRow = worksheet.addRow(row);
    const rowNumber = excelRow.number;
    // 将字符串值转换为数字，否则SUM公式无法正确计算
    numericColumns.forEach(colIdx => {
      const cell = excelRow.getCell(colIdx + 1);
      const cellValue = cell.value;
      // 检查是否是行程单金额公式标记
      if (cellValue === "__FORMULA_ITINERARY__") {
        // P列行程单金额公式 = H+K+L+O（票面价/改签补差 + 机建 + 燃油费 + 改签手续费）
        cell.value = {
          formula: `H${rowNumber}+K${rowNumber}+L${rowNumber}+O${rowNumber}`
        };
        cell.numFmt = "0.00";
      } else {
        const numValue = parseFloat(String(cellValue ?? 0)) || 0;
        cell.value = numValue;
        cell.numFmt = "0.00";
      }
    });
    // 设置I列公式 = H列/1.09
    const cellI = excelRow.getCell(9);
    cellI.value = { formula: `ROUND(H${rowNumber}/1.09,2)` };
    cellI.numFmt = "0.00";
    // 设置J列公式 = I列*0.09
    const cellJ = excelRow.getCell(10);
    cellJ.value = { formula: `ROUND(I${rowNumber}*0.09,2)` };
    cellJ.numFmt = "0.00";
    // 设置M列公式 = L列/1.09
    const cellM = excelRow.getCell(13);
    cellM.value = { formula: `ROUND(L${rowNumber}/1.09,2)` };
    cellM.numFmt = "0.00";
    // 设置N列公式 = M列*0.09
    const cellN = excelRow.getCell(14);
    cellN.value = { formula: `ROUND(M${rowNumber}*0.09,2)` };
    cellN.numFmt = "0.00";
    // 设置S列公式 = R列/1.06
    const cellS = excelRow.getCell(19);
    cellS.value = { formula: `ROUND(R${rowNumber}/1.06,2)` };
    cellS.numFmt = "0.00";
    // 设置T列公式 = R列/1.06*0.06
    const cellT = excelRow.getCell(20);
    cellT.value = { formula: `ROUND(R${rowNumber}/1.06*0.06,2)` };
    cellT.numFmt = "0.00";
    // 设置V列公式 = I+K+M+S
    const cellV = excelRow.getCell(22);
    cellV.value = {
      formula: `ROUND(I${rowNumber},2)+K${rowNumber}+ROUND(M${rowNumber},2)+ROUND(S${rowNumber},2)`
    };
    cellV.numFmt = "0.00";
    // 设置W列公式 = J+N+T
    const cellW = excelRow.getCell(23);
    cellW.value = { formula: `J${rowNumber}+N${rowNumber}+T${rowNumber}` };
    cellW.numFmt = "0.00";
  }

  // 记录指定人员组数据的起始和结束行
  const groupAStartRow = 3;
  const groupAEndRow = 2 + groupAData.value.length;

  // 添加指定人员组小计行
  if (groupAData.value.length > 0) {
    const subtotalRow: any[] = new Array(NEW_HEADERS.length).fill("");
    subtotalRow[0] = "小计";

    const subtotalExcelRow = worksheet.addRow(subtotalRow);
    const rowNumber = subtotalExcelRow.number;

    // 合并A到G列
    worksheet.mergeCells(rowNumber, 1, rowNumber, 7);

    // 使用SUM公式计算各数字列的合计
    const colLetters = [
      "H",
      "I",
      "J",
      "K",
      "L",
      "M",
      "N",
      "O",
      "P",
      "Q",
      "R",
      "S",
      "T",
      "U",
      "V",
      "W"
    ];
    colLetters.forEach((col, idx) => {
      const cell = subtotalExcelRow.getCell(idx + 8);
      cell.value = {
        formula: `SUM(${col}${groupAStartRow}:${col}${groupAEndRow})`
      };
      cell.numFmt = "0.00";
    });

    // 设置小计行样式
    subtotalExcelRow.eachCell(cell => {
      cell.font = { bold: true };
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
  }

  // 添加其他人员组数据
  for (const row of groupBData.value) {
    const excelRow = worksheet.addRow(row);
    const rowNumber = excelRow.number;
    // 将字符串值转换为数字，否则SUM公式无法正确计算
    numericColumns.forEach(colIdx => {
      const cell = excelRow.getCell(colIdx + 1);
      const cellValue = cell.value;
      // 检查是否是行程单金额公式标记
      if (cellValue === "__FORMULA_ITINERARY__") {
        // P列行程单金额公式 = H+K+L+O（票面价/改签补差 + 机建 + 燃油费 + 改签手续费）
        cell.value = {
          formula: `H${rowNumber}+K${rowNumber}+L${rowNumber}+O${rowNumber}`
        };
        cell.numFmt = "0.00";
      } else {
        const numValue = parseFloat(String(cellValue ?? 0)) || 0;
        cell.value = numValue;
        cell.numFmt = "0.00";
      }
    });
    // 设置I列公式 = H列/1.09
    const cellI = excelRow.getCell(9);
    cellI.value = { formula: `ROUND(H${rowNumber}/1.09,2)` };
    cellI.numFmt = "0.00";
    // 设置J列公式 = I列*0.09
    const cellJ = excelRow.getCell(10);
    cellJ.value = { formula: `ROUND(I${rowNumber}*0.09,2)` };
    cellJ.numFmt = "0.00";
    // 设置M列公式 = L列/1.09
    const cellM = excelRow.getCell(13);
    cellM.value = { formula: `ROUND(L${rowNumber}/1.09,2)` };
    cellM.numFmt = "0.00";
    // 设置N列公式 = M列*0.09
    const cellN = excelRow.getCell(14);
    cellN.value = { formula: `ROUND(M${rowNumber}*0.09,2)` };
    cellN.numFmt = "0.00";
    // 设置S列公式 = R列/1.06
    const cellS = excelRow.getCell(19);
    cellS.value = { formula: `ROUND(R${rowNumber}/1.06,2)` };
    cellS.numFmt = "0.00";
    // 设置T列公式 = R列/1.06*0.06
    const cellT = excelRow.getCell(20);
    cellT.value = { formula: `ROUND(R${rowNumber}/1.06*0.06,2)` };
    cellT.numFmt = "0.00";
    // 设置V列公式 = I+K+M+S
    const cellV = excelRow.getCell(22);
    cellV.value = {
      formula: `ROUND(I${rowNumber},2)+K${rowNumber}+ROUND(M${rowNumber},2)+ROUND(S${rowNumber},2)`
    };
    cellV.numFmt = "0.00";
    // 设置W列公式 = J+N+T
    const cellW = excelRow.getCell(23);
    cellW.value = { formula: `J${rowNumber}+N${rowNumber}+T${rowNumber}` };
    cellW.numFmt = "0.00";
  }

  // 添加其他人员组小计行
  if (groupBData.value.length > 0) {
    // 计算groupB数据的起始和结束行
    const groupBStartRow = groupAEndRow + (groupAData.value.length > 0 ? 2 : 1); // 加上groupA小计行（如果有）
    const groupBEndRow = groupBStartRow + groupBData.value.length - 1;

    const subtotalRow: any[] = new Array(NEW_HEADERS.length).fill("");
    subtotalRow[0] = "小计";

    const subtotalExcelRow = worksheet.addRow(subtotalRow);
    const rowNumber = subtotalExcelRow.number;

    // 合并A到G列
    worksheet.mergeCells(rowNumber, 1, rowNumber, 7);

    // 使用SUM公式计算各数字列的合计
    const colLetters = [
      "H",
      "I",
      "J",
      "K",
      "L",
      "M",
      "N",
      "O",
      "P",
      "Q",
      "R",
      "S",
      "T",
      "U",
      "V",
      "W"
    ];
    colLetters.forEach((col, idx) => {
      const cell = subtotalExcelRow.getCell(idx + 8);
      cell.value = {
        formula: `SUM(${col}${groupBStartRow}:${col}${groupBEndRow})`
      };
      cell.numFmt = "0.00";
    });

    // 设置小计行样式
    subtotalExcelRow.eachCell(cell => {
      cell.font = { bold: true };
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
  }

  // 添加合计行（合并两组数据）
  if (groupAData.value.length > 0 || groupBData.value.length > 0) {
    const totalRow: any[] = new Array(NEW_HEADERS.length).fill("");
    totalRow[0] = "合计";

    const totalExcelRow = worksheet.addRow(totalRow);
    const rowNumber = totalExcelRow.number;

    // 合并A到G列
    worksheet.mergeCells(rowNumber, 1, rowNumber, 7);

    // 计算groupB数据的起始和结束行
    const groupBStartRow = groupAEndRow + (groupAData.value.length > 0 ? 2 : 1);
    const groupBEndRow = groupBStartRow + groupBData.value.length - 1;

    // 使用SUM公式计算各数字列的合计（两组数据相加）
    const colLetters = [
      "H",
      "I",
      "J",
      "K",
      "L",
      "M",
      "N",
      "O",
      "P",
      "Q",
      "R",
      "S",
      "T",
      "U",
      "V",
      "W"
    ];
    colLetters.forEach((col, idx) => {
      const cell = totalExcelRow.getCell(idx + 8);
      if (groupAData.value.length > 0 && groupBData.value.length > 0) {
        // 两组都有数据，相加
        cell.value = {
          formula: `SUM(${col}${groupAStartRow}:${col}${groupAEndRow})+SUM(${col}${groupBStartRow}:${col}${groupBEndRow})`
        };
      } else if (groupAData.value.length > 0) {
        // 只有groupA有数据
        cell.value = {
          formula: `SUM(${col}${groupAStartRow}:${col}${groupAEndRow})`
        };
      } else if (groupBData.value.length > 0) {
        // 只有groupB有数据
        cell.value = {
          formula: `SUM(${col}${groupBStartRow}:${col}${groupBEndRow})`
        };
      }
      cell.numFmt = "0.00";
    });

    // 设置合计行样式（无底色）
    totalExcelRow.eachCell(cell => {
      cell.font = { bold: true };
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" }
      };
    });
  }

  // 设置表头行样式（第2行，因为第1行是标题）
  const headerRow = worksheet.getRow(2);
  headerRow.height = 22;

  // 特殊颜色的列（I=9, J=10, M=13, N=14, S=19, T=20, V=22, W=23）
  const greenColumns = [9, 10, 13, 14, 19, 20, 22, 23];

  headerRow.eachCell((cell, colNumber) => {
    cell.font = { bold: true };
    cell.alignment = { horizontal: "center", vertical: "middle" };
    // 检查是否为特殊列
    if (greenColumns.includes(colNumber)) {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF92D050" }
      };
    } else {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFFF99" }
      };
    }
    cell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" }
    };
  });

  // 设置所有数据行样式（从第3行开始，第1行是标题，第2行是表头）
  for (let i = 3; i <= worksheet.rowCount; i++) {
    const row = worksheet.getRow(i);
    row.height = 22;
    row.eachCell(cell => {
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.font = { size: 10 };
      // 小计行已有边框样式，其他行添加边框
      if (row.getCell(1).value !== "小计") {
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" }
        };
      }
    });
  }

  // 计算字符串显示宽度（中文2个宽度，英文/数字1个宽度）
  const getDisplayWidth = (str: string): number => {
    let width = 0;
    for (const char of str) {
      // 判断是否为中文字符
      if (/[\u4e00-\u9fa5]/.test(char)) {
        width += 2;
      } else if (/[^\x00-\xff]/.test(char)) {
        // 其他全角字符
        width += 2;
      } else {
        width += 1;
      }
    }
    return width;
  };

  // 设置固定列宽
  worksheet.getColumn(1).width = 8; // A列
  worksheet.getColumn(4).width = 23; // D列
  worksheet.getColumn(5).width = 9; // E列
  worksheet.getColumn(6).width = 10; // F列
  worksheet.getColumn(7).width = 10; // G列
  worksheet.getColumn(8).width = 16; // H列
  worksheet.getColumn(11).width = 10; // K列
  worksheet.getColumn(12).width = 12; // L列
  worksheet.getColumn(18).width = 10; // R列

  // 固定宽度的列索引（0-based）
  const fixedColumns = [0, 3, 4, 5, 6, 7, 10, 11, 17];

  // 其他列自动调整列宽
  worksheet.columns.forEach((column, index) => {
    if (fixedColumns.includes(index)) return; // 跳过固定列
    let maxWidth = 4;
    column.eachCell?.({ includeEmpty: true }, cell => {
      const cellValue = cell.value?.toString() || "";
      const displayWidth = getDisplayWidth(cellValue);
      maxWidth = Math.max(maxWidth, displayWidth);
    });
    // 设置紧凑的列宽范围
    column.width = Math.max(4, Math.min(maxWidth + 1, 18));
  });

  const buffer = await workbook.xlsx.writeBuffer();
  return new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
};

// 生成并下载文件
const generateAllFiles = async () => {
  if (sortedData.value.length === 0) {
    ElMessage.warning("没有可导出的数据");
    return;
  }

  generating.value = true;

  try {
    const blob = await generateExcel();
    saveAs(blob, `国宝造币账单.xlsx`);
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
  <div class="gbzb-bill-split">
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

    <!-- 统计信息和下载 -->
    <el-card v-if="statsInfo" class="result-card">
      <template #header>
        <div class="card-header">
          <span>数据处理结果</span>
          <el-button
            type="primary"
            :loading="generating"
            @click="generateAllFiles"
          >
            {{ generating ? "生成中..." : "生成并下载" }}
          </el-button>
        </div>
      </template>

      <!-- 统计信息 -->
      <div class="stats-container">
        <el-descriptions :column="3" border>
          <el-descriptions-item label="指定人员数据">
            <el-tag type="success">{{ statsInfo.groupACount }} 条</el-tag>
          </el-descriptions-item>
          <el-descriptions-item label="其他人员数据">
            <el-tag type="warning">{{ statsInfo.groupBCount }} 条</el-tag>
          </el-descriptions-item>
          <el-descriptions-item label="总计">
            <el-tag type="primary">{{ sortedData.length }} 条</el-tag>
          </el-descriptions-item>
        </el-descriptions>

        <!-- 指定人员详情 -->
        <div class="person-detail">
          <p><strong>指定人员数据详情（按顺序排列）：</strong></p>
          <div class="person-tags">
            <el-tag
              v-for="person in GROUP_A_PERSONS"
              :key="person"
              :type="statsInfo.groupAPersons.get(person) ? 'success' : 'info'"
            >
              {{ person }}: {{ statsInfo.groupAPersons.get(person) || 0 }} 条
            </el-tag>
          </div>
        </div>
      </div>
    </el-card>

    <!-- 无数据提示 -->
    <el-empty v-if="showData && !sheetData" description="未找到有效数据" />
  </div>
</template>

<style scoped>
.gbzb-bill-split {
  padding: 20px;
}

.upload-card,
.preview-card,
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

.preview-tip {
  color: #909399;
  font-size: 12px;
  margin-top: 10px;
  text-align: right;
}

.stats-container {
  padding: 10px 0;
}

.person-detail {
  margin-top: 20px;
}

.person-detail p {
  margin-bottom: 10px;
}

.person-tags {
  display: flex;
  flex-wrap: wrap;
  gap: 10px;
}

:deep(.el-upload-dragger) {
  width: 100%;
}
</style>
