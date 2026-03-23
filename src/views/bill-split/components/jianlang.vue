<script setup lang="ts">
import { ref } from "vue";
import { ElMessage } from "element-plus";
import { UploadFilled } from "@element-plus/icons-vue";
import ExcelJS from "exceljs";
import * as XLSX from "xlsx";

defineOptions({
  name: "JianlangBillSplit"
});

// 客户账单文件上传相关
const customerBillFile = ref<File | null>(null);
const customerBillLoading = ref(false);
// 国内机票数据
const flightData = ref<{
  headers: any[];
  data: any[][];
} | null>(null);
// 酒店数据
const hotelData = ref<{
  headers: any[];
  data: any[][];
} | null>(null);

// TMC文件上传相关
const tmcFile = ref<File | null>(null);
const tmcLoading = ref(false);
// TMC出票数据
const tmcIssueData = ref<{
  headers: any[];
  data: any[][];
} | null>(null);
// TMC改签数据
const tmcChangeData = ref<{
  headers: any[];
  data: any[][];
} | null>(null);
// TMC退票数据
const tmcRefundData = ref<{
  headers: any[];
  data: any[][];
} | null>(null);

// 酒店账单文件上传相关
const hotelBillFile = ref<File | null>(null);
const hotelBillLoading = ref(false);
// 酒店账单数据
const hotelBillData = ref<{
  headers: any[];
  data: any[][];
} | null>(null);

// 对比结果
const comparing = ref(false);
const compareResult = ref<any[]>([]);
const showCompareResult = ref(false);
const compareFullscreen = ref(false);

// 读取Excel文件（单个工作表）
const readExcelSheet = async (
  file: File,
  sheetName?: string
): Promise<{ headers: any[]; data: any[][] }> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = async e => {
      try {
        const data = e.target?.result;

        if (file.name.endsWith(".xlsx")) {
          const workbook = new ExcelJS.Workbook();
          await workbook.xlsx.load(data as ArrayBuffer);

          let worksheet: ExcelJS.Worksheet;
          if (sheetName) {
            worksheet =
              workbook.getWorksheet(sheetName) || workbook.worksheets[0];
          } else {
            worksheet = workbook.worksheets[0];
          }

          // 获取单元格实际值（处理公式情况）
          const getCellValue = (cell: ExcelJS.Cell): any => {
            const value = cell.value;
            // 如果是公式单元格，value 是对象 { formula: '...', result: ... }
            if (value && typeof value === "object" && "result" in value) {
              return (value as any).result;
            }
            return value;
          };

          const headers: any[] = [];
          const rows: any[][] = [];

          worksheet.eachRow((row, rowNumber) => {
            const rowData: any[] = [];
            row.eachCell({ includeEmpty: true }, cell => {
              const colIndex = Number(cell.col) - 1;
              rowData[colIndex] = getCellValue(cell);
            });

            if (rowNumber === 1) {
              headers.push(...rowData);
            } else {
              rows.push(rowData);
            }
          });

          resolve({ headers, data: rows });
        } else {
          const workbook = XLSX.read(data, { type: "array" });
          const targetSheetName = sheetName || workbook.SheetNames[0];
          const worksheet =
            workbook.Sheets[targetSheetName] ||
            workbook.Sheets[workbook.SheetNames[0]];
          const jsonData = XLSX.utils.sheet_to_json(worksheet, {
            header: 1
          }) as any[][];

          const headers = jsonData[0] || [];
          const rows = jsonData.slice(1);

          resolve({ headers, data: rows });
        }
      } catch (error) {
        reject(error);
      }
    };

    reader.onerror = () => reject(new Error("文件读取失败"));
    reader.readAsArrayBuffer(file);
  });
};

// 读取Excel文件（多个工作表）
const readExcelMultipleSheets = async (
  file: File,
  sheetNames: string[]
): Promise<Record<string, { headers: any[]; data: any[][] }>> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = async e => {
      try {
        const data = e.target?.result;
        const result: Record<string, { headers: any[]; data: any[][] }> = {};

        if (file.name.endsWith(".xlsx")) {
          const workbook = new ExcelJS.Workbook();
          await workbook.xlsx.load(data as ArrayBuffer);

          // 获取单元格实际值（处理公式情况）
          const getCellValue = (cell: ExcelJS.Cell): any => {
            const value = cell.value;
            // 如果是公式单元格，value 是对象 { formula: '...', result: ... }
            if (value && typeof value === "object" && "result" in value) {
              return (value as any).result;
            }
            return value;
          };

          for (const sheetName of sheetNames) {
            const worksheet = workbook.getWorksheet(sheetName);

            if (worksheet) {
              const headers: any[] = [];
              const rows: any[][] = [];

              worksheet.eachRow((row, rowNumber) => {
                const rowData: any[] = [];
                row.eachCell({ includeEmpty: true }, cell => {
                  const colIndex = Number(cell.col) - 1;
                  rowData[colIndex] = getCellValue(cell);
                });

                if (rowNumber === 1) {
                  headers.push(...rowData);
                } else {
                  rows.push(rowData);
                }
              });

              result[sheetName] = { headers, data: rows };
            }
          }

          resolve(result);
        } else {
          const workbook = XLSX.read(data, { type: "array" });

          for (const sheetName of sheetNames) {
            if (workbook.SheetNames.includes(sheetName)) {
              const worksheet = workbook.Sheets[sheetName];
              const jsonData = XLSX.utils.sheet_to_json(worksheet, {
                header: 1
              }) as any[][];

              result[sheetName] = {
                headers: jsonData[0] || [],
                data: jsonData.slice(1)
              };
            }
          }

          resolve(result);
        }
      } catch (error) {
        reject(error);
      }
    };

    reader.onerror = () => reject(new Error("文件读取失败"));
    reader.readAsArrayBuffer(file);
  });
};

// 处理客户账单文件上传
const handleCustomerBillFileChange = async (uploadFile: any) => {
  const file = uploadFile.raw;
  if (!file) return;

  customerBillLoading.value = true;

  try {
    const result = await readExcelMultipleSheets(file, ["国内机票", "酒店"]);

    customerBillFile.value = file;

    // 处理国内机票数据
    if (result["国内机票"]) {
      flightData.value = result["国内机票"];
      console.log("=== 国内机票数据 ===");
      console.log("表头:", result["国内机票"].headers);
      console.log("数据行数:", result["国内机票"].data.length);
      console.log("前5行数据:", result["国内机票"].data.slice(0, 5));
    } else {
      flightData.value = null;
    }

    // 处理酒店数据
    if (result["酒店"]) {
      hotelData.value = result["酒店"];
      console.log("=== 酒店数据 ===");
      console.log("表头:", result["酒店"].headers);
      console.log("数据行数:", result["酒店"].data.length);
      console.log("前5行数据:", result["酒店"].data.slice(0, 5));
    } else {
      hotelData.value = null;
    }

    const flightCount = flightData.value?.data?.length || 0;
    const hotelCount = hotelData.value?.data?.length || 0;

    if (flightCount === 0 && hotelCount === 0) {
      ElMessage.warning("未找到'国内机票'或'酒店'工作表");
    } else {
      ElMessage.success(
        `客户账单上传成功，国内机票 ${flightCount} 条，酒店 ${hotelCount} 条`
      );
    }
  } catch (error) {
    console.error("读取客户账单文件失败:", error);
    ElMessage.error("读取客户账单文件失败");
  } finally {
    customerBillLoading.value = false;
  }
};

// 处理TMC文件上传
const handleTmcFileChange = async (uploadFile: any) => {
  const file = uploadFile.raw;
  if (!file) return;

  tmcLoading.value = true;

  try {
    const result = await readExcelMultipleSheets(file, ["出票", "改签", "退票"]);

    tmcFile.value = file;

    // 处理出票数据
    if (result["出票"]) {
      tmcIssueData.value = result["出票"];
      console.log("=== TMC出票数据 ===");
      console.log("表头:", result["出票"].headers);
      console.log("数据行数:", result["出票"].data.length);
      console.log("前5行数据:", result["出票"].data.slice(0, 5));
    } else {
      tmcIssueData.value = null;
    }

    // 处理改签数据
    if (result["改签"]) {
      tmcChangeData.value = result["改签"];
      console.log("=== TMC改签数据 ===");
      console.log("表头:", result["改签"].headers);
      console.log("数据行数:", result["改签"].data.length);
      console.log("前5行数据:", result["改签"].data.slice(0, 5));
    } else {
      tmcChangeData.value = null;
    }

    // 处理退票数据
    if (result["退票"]) {
      tmcRefundData.value = result["退票"];
      console.log("=== TMC退票数据 ===");
      console.log("表头:", result["退票"].headers);
      console.log("数据行数:", result["退票"].data.length);
      console.log("前5行数据:", result["退票"].data.slice(0, 5));
    } else {
      tmcRefundData.value = null;
    }

    const issueCount = tmcIssueData.value?.data?.length || 0;
    const changeCount = tmcChangeData.value?.data?.length || 0;
    const refundCount = tmcRefundData.value?.data?.length || 0;

    if (issueCount === 0 && changeCount === 0 && refundCount === 0) {
      ElMessage.warning("未找到'出票'、'改签'或'退票'工作表");
    } else {
      ElMessage.success(
        `TMC文件上传成功，出票 ${issueCount} 条，改签 ${changeCount} 条，退票 ${refundCount} 条`
      );
    }
  } catch (error) {
    console.error("读取TMC文件失败:", error);
    ElMessage.error("读取TMC文件失败");
  } finally {
    tmcLoading.value = false;
  }
};

// 清除客户账单文件
const clearCustomerBillFile = () => {
  customerBillFile.value = null;
  flightData.value = null;
  hotelData.value = null;
};

// 清除TMC文件
const clearTmcFile = () => {
  tmcFile.value = null;
  tmcIssueData.value = null;
  tmcChangeData.value = null;
  tmcRefundData.value = null;
};

// 处理酒店账单文件上传
const handleHotelBillFileChange = async (uploadFile: any) => {
  const file = uploadFile.raw;
  if (!file) return;

  hotelBillLoading.value = true;

  try {
    const result = await readExcelSheet(file, "酒店明细(国内)");

    hotelBillFile.value = file;
    hotelBillData.value = result;

    console.log("=== 酒店账单数据（酒店明细(国内)）===");
    console.log("表头:", result.headers);
    console.log("数据行数:", result.data.length);
    console.log("前5行数据:", result.data.slice(0, 5));

    if (result.data.length === 0) {
      ElMessage.warning("酒店账单文件无数据或未找到'酒店明细(国内)'工作表");
    } else {
      ElMessage.success(`酒店账单上传成功，共 ${result.data.length} 条数据`);
    }
  } catch (error) {
    console.error("读取酒店账单文件失败:", error);
    ElMessage.error("读取酒店账单文件失败");
  } finally {
    hotelBillLoading.value = false;
  }
};

// 清除酒店账单文件
const clearHotelBillFile = () => {
  hotelBillFile.value = null;
  hotelBillData.value = null;
};

// 对比数据
const compareData = () => {
  // 检查是否有可对比的数据
  const hasFlightCompare = flightData.value && (tmcIssueData.value || tmcChangeData.value || tmcRefundData.value);
  const hasHotelCompare = hotelData.value && hotelBillData.value;

  if (!hasFlightCompare && !hasHotelCompare) {
    ElMessage.warning("请先上传必要的文件进行对比");
    return;
  }

  comparing.value = true;
  compareResult.value = [];

  try {
    // 机票对比
    if (hasFlightCompare) {
      const flightHeaders = flightData.value!.headers;
      const flightRows = flightData.value!.data;

    // 找到客户账单的列索引
    // M列是第13列，索引为12
    const ticketNoColIndex = 12; // M列 - 承运人-票号
    // U列是第21列，索引为20
    const refundFeeColIndex = 20; // U列 - 退票手续费
    // V列是第22列，索引为21
    const changeFeeColIndex = 21; // V列 - 改签手续费
    // 查找结算金额列
    const settlementColIndex = flightHeaders.findIndex(
      h => h && h.toString().includes("结算金额")
    );

    console.log("=== 客户账单列索引 ===");
    console.log("承运人-票号 列索引:", ticketNoColIndex);
    console.log("退票手续费 列索引 (U列):", refundFeeColIndex);
    console.log("结算金额 列索引:", settlementColIndex);
    console.log("改签手续费 列索引 (V列):", changeFeeColIndex);

    // ========== 出票对比 ==========
    if (tmcIssueData.value) {
      const tmcHeaders = tmcIssueData.value.headers;
      const tmcRows = tmcIssueData.value.data;

      // 找到TMC出票的列索引
      const tmcTicketNoColIndex = tmcHeaders.findIndex(
        h => h && h.toString().includes("全票号")
      );
      const tmcAmountColIndex = tmcHeaders.findIndex(
        h => h && h.toString().includes("应收金额")
      );
      // L列是第12列，索引为11 - 乘机日
      const tmcFlightDateColIndex = 11;

      console.log("=== TMC出票列索引 ===");
      console.log("全票号 列索引:", tmcTicketNoColIndex);
      console.log("应收金额 列索引:", tmcAmountColIndex);
      console.log("乘机日 列索引 (L列):", tmcFlightDateColIndex);

      if (settlementColIndex === -1) {
        ElMessage.warning("客户账单未找到'结算金额'列");
      } else if (tmcTicketNoColIndex === -1) {
        ElMessage.warning("TMC出票未找到'全票号'列");
      } else if (tmcAmountColIndex === -1) {
        ElMessage.warning("TMC出票未找到'应收金额'列");
      } else {
        // 格式化日期为 yyyy/M/d 格式
        const formatDate = (dateValue: any): string => {
          if (!dateValue) return "";
          // 如果是Date对象
          if (dateValue instanceof Date) {
            const year = dateValue.getFullYear();
            const month = dateValue.getMonth() + 1;
            const day = dateValue.getDate();
            return `${year}/${month}/${day}`;
          }
          // 如果是字符串，尝试解析
          const str = dateValue.toString().trim();
          if (str) return str;
          return "";
        };

        // 构建TMC出票数据映射
        const tmcIssueMap = new Map<string, { row: any[]; amount: number; flightDate: string }>();
        tmcRows.forEach(row => {
          const ticketNo = row[tmcTicketNoColIndex]?.toString().trim();
          const amount = parseFloat(row[tmcAmountColIndex]) || 0;
          const flightDate = formatDate(row[tmcFlightDateColIndex]);
          if (ticketNo) {
            tmcIssueMap.set(ticketNo, { row, amount, flightDate });
          }
        });

        // 构建客户账单出票数据映射（按票号汇总结算金额）
        const customerIssueMap = new Map<string, { totalAmount: number; rowNumbers: number[] }>();
        flightRows.forEach((row, index) => {
          const ticketNo = row[ticketNoColIndex]?.toString().trim();
          const settlementAmount = parseFloat(row[settlementColIndex]) || 0;
          const changeFee = parseFloat(row[changeFeeColIndex]) || 0;

          // 结算金额需要大于0，且改签手续费需要等于0
          if (settlementAmount <= 0) return;
          if (changeFee !== 0) return;
          if (!ticketNo) return;

          if (customerIssueMap.has(ticketNo)) {
            const existing = customerIssueMap.get(ticketNo)!;
            existing.totalAmount += settlementAmount;
            existing.rowNumbers.push(index + 2);
          } else {
            customerIssueMap.set(ticketNo, {
              totalAmount: settlementAmount,
              rowNumbers: [index + 2]
            });
          }
        });

        const matchedTmcIssueTickets = new Set<string>();

        // 遍历客户账单汇总数据进行出票对比
        customerIssueMap.forEach((customerRecord, ticketNo) => {
          const tmcRecord = tmcIssueMap.get(ticketNo);

          if (tmcRecord) {
            matchedTmcIssueTickets.add(ticketNo);
            const tmcAmount = tmcRecord.amount;
            const diff = customerRecord.totalAmount - tmcAmount;

            if (Math.abs(diff) > 0.01) {
              compareResult.value.push({
                category: "出票",
                type: "金额不匹配",
                ticketNo,
                customerAmount: customerRecord.totalAmount,
                tmcAmount,
                diff: diff.toFixed(2),
                customerRow: customerRecord.rowNumbers.join(","),
                tmcRow: tmcRows.indexOf(tmcRecord.row) + 2
              });
            }
          } else {
            compareResult.value.push({
              category: "出票",
              type: "客户有TMC无",
              ticketNo,
              customerAmount: customerRecord.totalAmount,
              tmcAmount: "-",
              diff: customerRecord.totalAmount.toFixed(2),
              customerRow: customerRecord.rowNumbers.join(","),
              tmcRow: "-"
            });
          }
        });

        // TMC出票有但客户账单没有
        tmcRows.forEach((row, index) => {
          const ticketNo = row[tmcTicketNoColIndex]?.toString().trim();
          const tmcAmount = parseFloat(row[tmcAmountColIndex]) || 0;
          const flightDate = formatDate(row[tmcFlightDateColIndex]);

          if (!ticketNo) return;
          if (matchedTmcIssueTickets.has(ticketNo)) return;

          // 拼接票号和乘机日
          const displayTicketNo = flightDate ? `${ticketNo}（起飞时间：${flightDate}）` : ticketNo;

          compareResult.value.push({
            category: "出票",
            type: "TMC有客户无",
            ticketNo: displayTicketNo,
            customerAmount: "-",
            tmcAmount,
            diff: (-tmcAmount).toFixed(2),
            customerRow: "-",
            tmcRow: index + 2
          });
        });
      }
    }

    // ========== 改签对比 ==========
    if (tmcChangeData.value) {
      const tmcChangeHeaders = tmcChangeData.value.headers;
      const tmcChangeRows = tmcChangeData.value.data;

      // 找到TMC改签的列索引
      const tmcChangeTicketNoColIndex = tmcChangeHeaders.findIndex(
        h => h && h.toString().includes("票号")
      );
      // 客户改签费用
      const tmcCustomerChangeFeeColIndex = tmcChangeHeaders.findIndex(
        h => h && h.toString().includes("客户改签费用")
      );

      console.log("=== TMC改签列索引 ===");
      console.log("票号 列索引:", tmcChangeTicketNoColIndex);
      console.log("客户改签费用 列索引:", tmcCustomerChangeFeeColIndex);

      if (tmcChangeTicketNoColIndex === -1) {
        ElMessage.warning("TMC改签未找到'票号'列");
      } else if (tmcCustomerChangeFeeColIndex === -1) {
        ElMessage.warning("TMC改签未找到'客户改签费用'列");
      } else {
        // 票号处理函数：用 "-" 分割后取最后一项
        const getTicketNoKey = (ticketNo: string) => {
          const trimmed = ticketNo?.toString().trim();
          if (!trimmed) return "";
          const parts = trimmed.split("-");
          return parts[parts.length - 1] || trimmed;
        };

        // 构建TMC改签数据映射
        const tmcChangeMap = new Map<string, { row: any[]; customerChangeFee: number; originalTicketNo: string }>();
        tmcChangeRows.forEach(row => {
          const originalTicketNo = row[tmcChangeTicketNoColIndex]?.toString().trim();
          const ticketNoKey = getTicketNoKey(originalTicketNo);
          const customerChangeFee = parseFloat(row[tmcCustomerChangeFeeColIndex]) || 0;
          if (ticketNoKey) {
            tmcChangeMap.set(ticketNoKey, { row, customerChangeFee, originalTicketNo });
          }
        });

        const matchedTmcChangeTickets = new Set<string>();

        // 遍历客户账单数据进行改签对比
        flightRows.forEach((row, index) => {
          const originalTicketNo = row[ticketNoColIndex]?.toString().trim();
          const ticketNoKey = getTicketNoKey(originalTicketNo);
          const changeFee = parseFloat(row[changeFeeColIndex]) || 0;
          const settlementAmount = parseFloat(row[settlementColIndex]) || 0;

          // 改签手续费需要大于0
          if (changeFee <= 0) return;
          if (!ticketNoKey) return;

          const tmcRecord = tmcChangeMap.get(ticketNoKey);

          if (tmcRecord) {
            matchedTmcChangeTickets.add(ticketNoKey);

            // 对比结算金额与TMC客户改签费用
            const tmcCustomerChangeFee = tmcRecord.customerChangeFee;
            const diff = settlementAmount - tmcCustomerChangeFee;
            if (Math.abs(diff) > 0.01) {
              compareResult.value.push({
                category: "改签",
                type: "金额不匹配",
                ticketNo: originalTicketNo,
                customerAmount: settlementAmount,
                tmcAmount: tmcCustomerChangeFee,
                diff: diff.toFixed(2),
                customerRow: index + 2,
                tmcRow: tmcChangeRows.indexOf(tmcRecord.row) + 2
              });
            }
          } else {
            // 客户有改签数据，TMC没有
            compareResult.value.push({
              category: "改签",
              type: "客户有TMC无",
              ticketNo: originalTicketNo,
              customerAmount: settlementAmount,
              tmcAmount: "-",
              diff: settlementAmount.toFixed(2),
              customerRow: index + 2,
              tmcRow: "-"
            });
          }
        });

        // TMC改签有但客户账单没有
        tmcChangeRows.forEach((row, index) => {
          const originalTicketNo = row[tmcChangeTicketNoColIndex]?.toString().trim();
          const ticketNoKey = getTicketNoKey(originalTicketNo);
          const tmcCustomerChangeFee = parseFloat(row[tmcCustomerChangeFeeColIndex]) || 0;

          if (!ticketNoKey) return;
          if (matchedTmcChangeTickets.has(ticketNoKey)) return;
          if (tmcCustomerChangeFee <= 0) return;

          compareResult.value.push({
            category: "改签",
            type: "TMC有客户无",
            ticketNo: originalTicketNo,
            customerAmount: "-",
            tmcAmount: tmcCustomerChangeFee,
            diff: (-tmcCustomerChangeFee).toFixed(2),
            customerRow: "-",
            tmcRow: index + 2
          });
        });
      }
    }

    // ========== 退票对比 ==========
    if (tmcRefundData.value) {
      const tmcRefundHeaders = tmcRefundData.value.headers;
      const tmcRefundRows = tmcRefundData.value.data;

      // 找到TMC退票的列索引
      // N列是第14列，索引为13 - 票面_承运人-票号
      const tmcRefundTicketNoColIndex = tmcRefundHeaders.findIndex(
        h => h && h.toString().includes("票面_承运人-票号")
      );
      // AF列是第32列，索引为31 - 应退金额
      const tmcRefundAmountColIndex = tmcRefundHeaders.findIndex(
        h => h && h.toString().includes("应退金额")
      );

      console.log("=== TMC退票列索引 ===");
      console.log("票面_承运人-票号 列索引:", tmcRefundTicketNoColIndex);
      console.log("应退金额 列索引:", tmcRefundAmountColIndex);

      if (tmcRefundTicketNoColIndex === -1) {
        ElMessage.warning("TMC退票未找到'票面_承运人-票号'列");
      } else if (tmcRefundAmountColIndex === -1) {
        ElMessage.warning("TMC退票未找到'应退金额'列");
      } else {
        // 构建TMC退票数据映射（按票号汇总金额）
        const tmcRefundMap = new Map<string, { rows: any[]; totalAmount: number; rowNumbers: number[] }>();
        tmcRefundRows.forEach((row, idx) => {
          const ticketNo = row[tmcRefundTicketNoColIndex]?.toString().trim();
          const amount = parseFloat(row[tmcRefundAmountColIndex]) || 0;
          if (ticketNo) {
            if (tmcRefundMap.has(ticketNo)) {
              const existing = tmcRefundMap.get(ticketNo)!;
              existing.rows.push(row);
              existing.totalAmount += amount;
              existing.rowNumbers.push(idx + 2);
            } else {
              tmcRefundMap.set(ticketNo, {
                rows: [row],
                totalAmount: amount,
                rowNumbers: [idx + 2]
              });
            }
          }
        });

        console.log("=== TMC退票汇总数据 ===");
        tmcRefundMap.forEach((value, key) => {
          console.log(`票号: ${key}, 汇总金额: ${value.totalAmount}, 行号: ${value.rowNumbers.join(",")}`);
        });

        // 构建客户账单退票数据映射（按票号汇总结算金额）
        const customerRefundMap = new Map<string, { totalAmount: number; rowNumbers: number[] }>();
        flightRows.forEach((row, index) => {
          const ticketNo = row[ticketNoColIndex]?.toString().trim();
          const refundFee = parseFloat(row[refundFeeColIndex]) || 0;
          const settlementAmount = parseFloat(row[settlementColIndex]) || 0;

          // 退票手续费需要大于0
          if (refundFee <= 0) return;
          if (!ticketNo) return;

          const absSettlement = Math.abs(settlementAmount);
          if (customerRefundMap.has(ticketNo)) {
            const existing = customerRefundMap.get(ticketNo)!;
            existing.totalAmount += absSettlement;
            existing.rowNumbers.push(index + 2);
          } else {
            customerRefundMap.set(ticketNo, {
              totalAmount: absSettlement,
              rowNumbers: [index + 2]
            });
          }
        });

        const matchedTmcRefundTickets = new Set<string>();

        // 遍历客户账单汇总数据进行退票对比
        customerRefundMap.forEach((customerRecord, ticketNo) => {
          const tmcRecord = tmcRefundMap.get(ticketNo);

          if (tmcRecord) {
            matchedTmcRefundTickets.add(ticketNo);
            const diff = customerRecord.totalAmount - tmcRecord.totalAmount;

            if (Math.abs(diff) > 0.01) {
              compareResult.value.push({
                category: "退票",
                type: "金额不匹配",
                ticketNo,
                customerAmount: customerRecord.totalAmount,
                tmcAmount: tmcRecord.totalAmount,
                diff: diff.toFixed(2),
                customerRow: customerRecord.rowNumbers.join(","),
                tmcRow: tmcRecord.rowNumbers.join(",")
              });
            }
          } else {
            compareResult.value.push({
              category: "退票",
              type: "客户有TMC无",
              ticketNo,
              customerAmount: customerRecord.totalAmount,
              tmcAmount: "-",
              diff: customerRecord.totalAmount.toFixed(2),
              customerRow: customerRecord.rowNumbers.join(","),
              tmcRow: "-"
            });
          }
        });

        // TMC退票有但客户账单没有（使用汇总后的数据，需要二次校验）
        tmcRefundMap.forEach((record, ticketNo) => {
          if (matchedTmcRefundTickets.has(ticketNo)) return;
          if (record.totalAmount <= 0) return;

          // 二次校验：在客户原始数据中查找该票号的所有记录
          // W列是第23列，索引为22 - TMC服务费
          const tmcServiceFeeColIndex = 22;
          let matchedCustomerRow: number | null = null;

          flightRows.forEach((row, idx) => {
            const customerTicketNo = row[ticketNoColIndex]?.toString().trim();
            if (customerTicketNo === ticketNo) {
              const settlementAmount = parseFloat(row[settlementColIndex]) || 0;
              // 只处理正数结算金额
              if (settlementAmount > 0) {
                const tmcServiceFee = parseFloat(row[tmcServiceFeeColIndex]) || 0;
                // 正数结算金额 - W列TMC服务费
                const adjustedAmount = settlementAmount - tmcServiceFee;

                // 如果处理后的金额等于TMC应退金额，则匹配成功
                if (Math.abs(adjustedAmount - record.totalAmount) <= 0.01) {
                  matchedCustomerRow = idx + 2;
                  console.log(`退票二次校验匹配成功: ${ticketNo}, TMC金额: ${record.totalAmount}, 客户结算金额: ${settlementAmount}, TMC服务费: ${tmcServiceFee}, 处理后金额: ${adjustedAmount}`);
                }
              }
            }
          });

          // 如果二次校验匹配成功，不算差异
          if (matchedCustomerRow !== null) {
            return;
          }

          compareResult.value.push({
            category: "退票",
            type: "TMC有客户无",
            ticketNo,
            customerAmount: "-",
            tmcAmount: record.totalAmount,
            diff: (-record.totalAmount).toFixed(2),
            customerRow: "-",
            tmcRow: record.rowNumbers.join(",")
          });
        });
      }
    }
    } // 结束 if (hasFlightCompare)

    // ========== 酒店对比 ==========
    if (hotelData.value && hotelBillData.value) {
      const hotelHeaders = hotelData.value.headers;
      const hotelRows = hotelData.value.data;
      const tmcHotelHeaders = hotelBillData.value.headers;
      const tmcHotelRows = hotelBillData.value.data;

      // 客户账单酒店工作表列索引
      // H列是第8列，索引为7 - 供应订单编号
      const customerOrderNoColIndex = 7;
      // Y列是第25列，索引为24 - 结算金额
      const customerSettlementColIndex = 24;

      // TMC酒店账单列索引
      // P列是第16列，索引为15 - 订单编号
      const tmcOrderNoColIndex = 15;
      // AG列是第33列，索引为32 - 应付金额
      const tmcAmountColIndex = 32;

      console.log("=== 酒店对比列索引 ===");
      console.log("客户账单-供应订单编号 列索引 (H列):", customerOrderNoColIndex);
      console.log("客户账单-结算金额 列索引 (Y列):", customerSettlementColIndex);
      console.log("TMC酒店-订单编号 列索引 (P列):", tmcOrderNoColIndex);
      console.log("TMC酒店-应付金额 列索引 (AG列):", tmcAmountColIndex);

      // M列是第13列，索引为12 - 入住人
      const guestNameColIndex = 12;

      // 构建客户账单酒店数据映射（按供应订单编号汇总结算金额）
      const customerHotelMap = new Map<string, { totalAmount: number; rowNumbers: number[]; guestNames: string[] }>();
      hotelRows.forEach((row, index) => {
        const orderNo = row[customerOrderNoColIndex]?.toString().trim();
        const settlementAmount = parseFloat(row[customerSettlementColIndex]) || 0;
        const guestName = row[guestNameColIndex]?.toString().trim() || "";

        if (!orderNo) return;

        if (customerHotelMap.has(orderNo)) {
          const existing = customerHotelMap.get(orderNo)!;
          existing.totalAmount += settlementAmount;
          existing.rowNumbers.push(index + 2);
          if (guestName && !existing.guestNames.includes(guestName)) {
            existing.guestNames.push(guestName);
          }
        } else {
          customerHotelMap.set(orderNo, {
            totalAmount: settlementAmount,
            rowNumbers: [index + 2],
            guestNames: guestName ? [guestName] : []
          });
        }
      });

      // 移除汇总金额为0的记录
      customerHotelMap.forEach((value, key) => {
        if (Math.abs(value.totalAmount) < 0.01) {
          customerHotelMap.delete(key);
        }
      });

      console.log("=== 客户账单酒店汇总数据 ===");
      customerHotelMap.forEach((value, key) => {
        console.log(`订单号: ${key}, 汇总金额: ${value.totalAmount}, 行号: ${value.rowNumbers.join(",")}`);
      });

      // 构建TMC酒店数据映射（按订单编号汇总应付金额，金额为0的不加入）
      const tmcHotelMap = new Map<string, { rows: any[]; totalAmount: number; rowNumbers: number[] }>();
      tmcHotelRows.forEach((row, idx) => {
        const orderNo = row[tmcOrderNoColIndex]?.toString().trim();
        const amount = parseFloat(row[tmcAmountColIndex]) || 0;
        if (orderNo) {
          if (tmcHotelMap.has(orderNo)) {
            const existing = tmcHotelMap.get(orderNo)!;
            existing.rows.push(row);
            existing.totalAmount += amount;
            existing.rowNumbers.push(idx + 2);
          } else {
            tmcHotelMap.set(orderNo, {
              rows: [row],
              totalAmount: amount,
              rowNumbers: [idx + 2]
            });
          }
        }
      });

      // 移除汇总金额为0的记录
      tmcHotelMap.forEach((value, key) => {
        if (Math.abs(value.totalAmount) < 0.01) {
          tmcHotelMap.delete(key);
        }
      });

      console.log("=== TMC酒店汇总数据 ===");
      tmcHotelMap.forEach((value, key) => {
        console.log(`订单号: ${key}, 汇总金额: ${value.totalAmount}, 行号: ${value.rowNumbers.join(",")}`);
      });

      const matchedTmcHotelOrders = new Set<string>();

      // 遍历客户账单汇总数据进行酒店对比
      customerHotelMap.forEach((customerRecord, orderNo) => {
        const tmcRecord = tmcHotelMap.get(orderNo);

        if (tmcRecord) {
          matchedTmcHotelOrders.add(orderNo);
          const tmcAmount = tmcRecord.totalAmount;
          const diff = customerRecord.totalAmount - tmcAmount;

          if (Math.abs(diff) > 0.01) {
            compareResult.value.push({
              category: "酒店",
              type: "金额不匹配",
              ticketNo: orderNo,
              customerAmount: customerRecord.totalAmount,
              tmcAmount,
              diff: diff.toFixed(2),
              customerRow: customerRecord.rowNumbers.join(","),
              tmcRow: tmcRecord.rowNumbers.join(",")
            });
          }
        } else {
          // 拼接订单号和入住人
          const guestNames = customerRecord.guestNames;
          const displayOrderNo = guestNames.length > 0 
            ? `${orderNo}（入住人：${guestNames.join(",")}）` 
            : orderNo;

          compareResult.value.push({
            category: "酒店",
            type: "客户有TMC无",
            ticketNo: displayOrderNo,
            customerAmount: customerRecord.totalAmount,
            tmcAmount: "-",
            diff: customerRecord.totalAmount.toFixed(2),
            customerRow: customerRecord.rowNumbers.join(","),
            tmcRow: "-"
          });
        }
      });

      // TMC酒店有但客户账单没有（使用汇总后的数据）
      tmcHotelMap.forEach((record, orderNo) => {
        if (matchedTmcHotelOrders.has(orderNo)) return;

        compareResult.value.push({
          category: "酒店",
          type: "TMC有客户无",
          ticketNo: orderNo,
          customerAmount: "-",
          tmcAmount: record.totalAmount,
          diff: (-record.totalAmount).toFixed(2),
          customerRow: "-",
          tmcRow: record.rowNumbers.join(",")
        });
      });
    }

    console.log("=== 对比结果 ===");
    console.log("差异总数:", compareResult.value.length);

    ElMessage.success(`对比完成，共发现 ${compareResult.value.length} 条差异`);
    showCompareResult.value = true;
  } catch (error) {
    console.error("对比失败:", error);
    ElMessage.error("对比失败");
  } finally {
    comparing.value = false;
  }
};

// 导出对比结果
const exportCompareResult = async () => {
  if (compareResult.value.length === 0) {
    ElMessage.warning("没有可导出的数据");
    return;
  }

  try {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("对比结果");

    // 添加表头
    worksheet.addRow([
      "序号",
      "类别",
      "差异类型",
      "票号",
      "客户账单金额",
      "TMC金额",
      "差额",
      "客户账单行号",
      "TMC行号"
    ]);

    // 添加数据
    compareResult.value.forEach((item, index) => {
      worksheet.addRow([
        index + 1,
        item.category,
        item.type,
        item.ticketNo,
        item.customerAmount,
        item.tmcAmount,
        item.diff,
        item.customerRow,
        item.tmcRow
      ]);
    });

    // 生成并下载文件
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `坚朗对比结果_${new Date().toLocaleDateString().replace(/\//g, "-")}.xlsx`;
    link.click();
    URL.revokeObjectURL(url);

    ElMessage.success("导出成功");
  } catch (error) {
    console.error("导出失败:", error);
    ElMessage.error("导出失败");
  }
};

// 重置所有数据
const resetAll = () => {
  customerBillFile.value = null;
  flightData.value = null;
  hotelData.value = null;
  tmcFile.value = null;
  tmcIssueData.value = null;
  tmcChangeData.value = null;
  tmcRefundData.value = null;
  hotelBillFile.value = null;
  hotelBillData.value = null;
  compareResult.value = [];
  showCompareResult.value = false;
};
</script>

<template>
  <div class="jianlang-bill-split">
    <!-- 文件上传区域 -->
    <div class="upload-section">
      <!-- 客户账单文件上传 -->
      <el-card class="upload-card">
        <template #header>
          <div class="card-header">
            <span>客户账单</span>
            <el-button
              v-if="customerBillFile"
              type="danger"
              size="small"
              @click="clearCustomerBillFile"
            >
              清除
            </el-button>
          </div>
        </template>

        <el-upload
          v-if="!customerBillFile"
          class="upload-area"
          drag
          :auto-upload="false"
          :show-file-list="false"
          :on-change="handleCustomerBillFileChange"
          accept=".xlsx,.xls"
        >
          <el-icon class="el-icon--upload" :size="50">
            <UploadFilled />
          </el-icon>
          <div class="el-upload__text">
            拖拽客户账单文件到此处，或<em>点击上传</em>
          </div>
          <div class="el-upload__tip">支持"国内机票"和"酒店"工作表</div>
        </el-upload>

        <div v-else class="file-uploaded">
          <el-icon :size="40" color="#67C23A">
            <i-ep-circle-check-filled />
          </el-icon>
          <div class="file-info">
            <span class="file-name">{{ customerBillFile.name }}</span>
            <span class="file-count">
              国内机票: {{ flightData?.data?.length || 0 }} 条 | 酒店:
              {{ hotelData?.data?.length || 0 }} 条
            </span>
          </div>
        </div>
      </el-card>

      <!-- TMC文件上传 -->
      <el-card class="upload-card">
        <template #header>
          <div class="card-header">
            <span>TMC文件</span>
            <el-button
              v-if="tmcFile"
              type="danger"
              size="small"
              @click="clearTmcFile"
            >
              清除
            </el-button>
          </div>
        </template>

        <el-upload
          v-if="!tmcFile"
          class="upload-area"
          drag
          :auto-upload="false"
          :show-file-list="false"
          :on-change="handleTmcFileChange"
          accept=".xlsx,.xls"
        >
          <el-icon class="el-icon--upload" :size="50">
            <UploadFilled />
          </el-icon>
          <div class="el-upload__text">
            拖拽TMC文件到此处，或<em>点击上传</em>
          </div>
          <div class="el-upload__tip">支持"出票"、"改签"和"退票"工作表</div>
        </el-upload>

        <div v-else class="file-uploaded">
          <el-icon :size="40" color="#67C23A">
            <i-ep-circle-check-filled />
          </el-icon>
          <div class="file-info">
            <span class="file-name">{{ tmcFile.name }}</span>
            <span class="file-count">
              出票: {{ tmcIssueData?.data?.length || 0 }} 条 | 改签:
              {{ tmcChangeData?.data?.length || 0 }} 条 | 退票:
              {{ tmcRefundData?.data?.length || 0 }} 条
            </span>
          </div>
        </div>
      </el-card>

      <!-- 酒店账单文件上传 -->
      <el-card class="upload-card">
        <template #header>
          <div class="card-header">
            <span>酒店账单</span>
            <el-button
              v-if="hotelBillFile"
              type="danger"
              size="small"
              @click="clearHotelBillFile"
            >
              清除
            </el-button>
          </div>
        </template>

        <el-upload
          v-if="!hotelBillFile"
          class="upload-area"
          drag
          :auto-upload="false"
          :show-file-list="false"
          :on-change="handleHotelBillFileChange"
          accept=".xlsx,.xls"
        >
          <el-icon class="el-icon--upload" :size="50">
            <UploadFilled />
          </el-icon>
          <div class="el-upload__text">
            拖拽酒店账单文件到此处，或<em>点击上传</em>
          </div>
          <div class="el-upload__tip">支持Excel文件</div>
        </el-upload>

        <div v-else class="file-uploaded">
          <el-icon :size="40" color="#67C23A">
            <i-ep-circle-check-filled />
          </el-icon>
          <div class="file-info">
            <span class="file-name">{{ hotelBillFile.name }}</span>
            <span class="file-count">
              酒店账单: {{ hotelBillData?.data?.length || 0 }} 条
            </span>
          </div>
        </div>
      </el-card>
    </div>

    <!-- 操作按钮 -->
    <div class="action-buttons">
      <el-button
        type="primary"
        size="large"
        :loading="comparing"
        :disabled="!flightData && !hotelData && (!tmcIssueData && !tmcChangeData && !tmcRefundData) && !hotelBillData"
        @click="compareData"
      >
        {{ comparing ? "对比中..." : "开始对比" }}
      </el-button>
      <el-button size="large" @click="resetAll"> 重置 </el-button>
    </div>

    <!-- 加载状态 -->
    <div v-if="customerBillLoading || tmcLoading || hotelBillLoading" class="loading-container">
      <el-icon class="is-loading" :size="40">
        <i class="el-icon-loading" />
      </el-icon>
      <p>正在解析文件...</p>
    </div>

    <!-- 对比结果 -->
    <div
      v-if="showCompareResult && compareResult.length > 0"
      class="compare-result"
    >
      <el-card>
        <template #header>
          <div class="card-header">
            <span>对比结果</span>
            <span class="result-count"
              >共 {{ compareResult.length }} 条差异</span
            >
            <div class="header-actions">
              <el-button
                type="primary"
                size="small"
                @click="exportCompareResult"
              >
                导出结果
              </el-button>
              <el-button
                type="success"
                size="small"
                @click="compareFullscreen = true"
              >
                全屏查看
              </el-button>
            </div>
          </div>
        </template>

        <el-table :data="compareResult" border stripe max-height="400">
          <el-table-column prop="category" label="类别" width="110">
            <template #default="{ row }">
              <el-tag
                :type="
                  row.category === '出票'
                    ? 'primary'
                    : row.category === '改签'
                      ? 'success'
                      : row.category === '退票'
                        ? 'warning'
                        : 'info'
                "
              >
                {{ row.category }}
              </el-tag>
            </template>
          </el-table-column>
          <el-table-column prop="type" label="差异类型" width="110">
            <template #default="{ row }">
              <el-tag
                :type="
                  row.type === '金额不匹配'
                    ? 'danger'
                    : row.type === '客户有TMC无'
                      ? 'warning'
                      : 'info'
                "
              >
                {{ row.type }}
              </el-tag>
            </template>
          </el-table-column>
          <el-table-column prop="ticketNo" label="票号" min-width="180" />
          <el-table-column
            prop="customerAmount"
            label="客户账单金额"
            width="130"
          />
          <el-table-column prop="tmcAmount" label="TMC金额" width="120" />
          <el-table-column prop="diff" label="差额" width="100">
            <template #default="{ row }">
              <span
                :style="{
                  color: parseFloat(row.diff) > 0 ? '#F56C6C' : '#67C23A'
                }"
              >
                {{ row.diff }}
              </span>
            </template>
          </el-table-column>
          <el-table-column
            prop="customerRow"
            label="客户账单行号"
            width="120"
          />
          <el-table-column prop="tmcRow" label="TMC行号" width="100" />
        </el-table>
      </el-card>
    </div>

    <!-- 全屏弹窗 -->
    <el-dialog
      v-model="compareFullscreen"
      title="对比结果"
      fullscreen
      :close-on-click-modal="false"
    >
      <div class="fullscreen-content">
        <div class="fullscreen-toolbar">
          <span class="result-count">共 {{ compareResult.length }} 条差异</span>
          <el-button type="primary" size="small" @click="exportCompareResult">
            导出结果
          </el-button>
        </div>
        <el-table
          :data="compareResult"
          border
          stripe
          height="calc(100vh - 180px)"
        >
          <el-table-column prop="category" label="类别" width="110" fixed>
            <template #default="{ row }">
              <el-tag
                :type="
                  row.category === '出票'
                    ? 'primary'
                    : row.category === '改签'
                      ? 'success'
                      : row.category === '退票'
                        ? 'warning'
                        : 'info'
                "
              >
                {{ row.category }}
              </el-tag>
            </template>
          </el-table-column>
          <el-table-column prop="type" label="差异类型" width="110">
            <template #default="{ row }">
              <el-tag
                :type="
                  row.type === '金额不匹配'
                    ? 'danger'
                    : row.type === '客户有TMC无'
                      ? 'warning'
                      : 'info'
                "
              >
                {{ row.type }}
              </el-tag>
            </template>
          </el-table-column>
          <el-table-column prop="ticketNo" label="票号" min-width="200" />
          <el-table-column
            prop="customerAmount"
            label="客户账单金额"
            width="140"
          />
          <el-table-column prop="tmcAmount" label="TMC金额" width="130" />
          <el-table-column prop="diff" label="差额" width="120">
            <template #default="{ row }">
              <span
                :style="{
                  color: parseFloat(row.diff) > 0 ? '#F56C6C' : '#67C23A'
                }"
              >
                {{ row.diff }}
              </span>
            </template>
          </el-table-column>
          <el-table-column
            prop="customerRow"
            label="客户账单行号"
            width="130"
          />
          <el-table-column prop="tmcRow" label="TMC行号" width="110" />
        </el-table>
      </div>
    </el-dialog>
  </div>
</template>

<style scoped>
.jianlang-bill-split {
  padding: 20px;
}

.upload-section {
  display: flex;
  gap: 20px;
  margin-bottom: 20px;
}

.upload-card {
  flex: 1;
}

.card-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
}

.upload-area {
  width: 100%;
}

.upload-area :deep(.el-upload-dragger) {
  width: 100%;
  height: 150px;
  display: flex;
  flex-direction: column;
  justify-content: center;
  align-items: center;
}

.el-upload__tip {
  margin-top: 10px;
  color: #909399;
  font-size: 12px;
}

.file-uploaded {
  display: flex;
  align-items: center;
  gap: 15px;
  padding: 20px;
}

.file-info {
  display: flex;
  flex-direction: column;
  gap: 5px;
}

.file-name {
  font-size: 16px;
  font-weight: 500;
  color: #303133;
}

.file-count {
  font-size: 14px;
  color: #909399;
}

.action-buttons {
  display: flex;
  justify-content: center;
  gap: 20px;
  margin-bottom: 20px;
}

.loading-container {
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  padding: 40px;
  color: #909399;
}

.compare-result {
  margin-top: 20px;
}

.header-actions {
  display: flex;
  gap: 10px;
}

.result-count {
  font-size: 14px;
  color: #909399;
}

.fullscreen-content {
  height: 100%;
}

.fullscreen-toolbar {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 15px;
}
</style>
