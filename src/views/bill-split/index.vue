<script setup lang="ts">
import { ref } from "vue";
import { ElMessage } from "element-plus";
import { UploadFilled } from "@element-plus/icons-vue";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import JSZip from "jszip";

defineOptions({
  name: "BillSplitIndex"
});

const uploadedFile = ref<File | null>(null);
const excelData = ref<any[]>([]);
const allSheetData = ref<Record<string, any[]>>({}); // 存储所有工作表数据
const originalWorkbook = ref<any>(null);
const loading = ref(false);
const showData = ref(false);
const generating = ref(false);
const editableFileNames = ref<
  { groupName: string; fileName: string; sheetType?: string }[]
>([]);

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
        .then(() => {
          // 处理多个工作表
          const sheetProcessors = [
            {
              name: "酒店明细(国内)",
              key: "hotel",
              departmentKeyword: "预订人部门"
            },
            {
              name: "酒店明细(国际)",
              key: "internationalHotel",
              departmentKeyword: "预订人部门"
            },
            {
              name: "火车票明细",
              key: "train",
              departmentKeyword: "预订人部门"
            },
            {
              name: "机票明细(国内)",
              key: "flight",
              departmentKeyword: "预订人部门"
            },
            {
              name: "机票明细(国际)",
              key: "internationalFlight",
              departmentKeyword: "预订人部门"
            }
          ];

          const sheetData: Record<string, any[]> = {};
          let processedSheets = 0;
          let totalSheets = 0;

          // 计算需要处理的工作表数量
          sheetProcessors.forEach(processor => {
            const worksheet = workbook.getWorksheet(processor.name);
            if (worksheet) totalSheets++;
          });

          if (totalSheets === 0) {
            ElMessage.error("未找到任何需要处理的工作表");
            console.log(
              "可用的工作表:",
              workbook.worksheets.map(ws => ws.name)
            );
            loading.value = false;
            return;
          }

          // 打印Excel文件信息
          console.log("Excel文件信息:");
          console.log("文件名:", file.name);
          console.log("文件大小:", file.size, "bytes");
          console.log("工作表数量:", workbook.worksheets.length);
          console.log(
            "所有工作表名称:",
            workbook.worksheets.map(ws => ws.name)
          );
          console.log(`找到 ${totalSheets} 个需要处理的工作表`);

          // 处理每个工作表
          sheetProcessors.forEach(processor => {
            const worksheet = workbook.getWorksheet(processor.name);
            if (!worksheet) {
              console.log(`跳过不存在的工作表: ${processor.name}`);
              return;
            }

            console.log(
              `\n========== 处理工作表: ${processor.name} ==========`
            );

            // 读取数据为二维数组
            const jsonData: any[][] = [];
            worksheet.eachRow((row, rowNumber) => {
              const rowData: any[] = [];
              row.eachCell((cell, colNumber) => {
                rowData.push(cell.value);
              });
              jsonData.push(rowData);
            });

            sheetData[processor.key] = jsonData;

            console.log(`${processor.name} - 数据行数:`, jsonData.length);
            console.log(
              `${processor.name} - 数据列数:`,
              (jsonData[0] as any[])?.length || 0
            );

            // 查找部门列
            if (jsonData.length > 2) {
              const headers = jsonData[0] as any[];
              const thirdRow = jsonData[2] as any[];

              // 在第三行中查找包含部门关键字的单元格
              const departmentColumnIndex = thirdRow.findIndex(
                (cell: any) =>
                  cell && cell.toString().includes(processor.departmentKeyword)
              );

              if (departmentColumnIndex !== -1) {
                const departmentColumnName =
                  headers[departmentColumnIndex] ||
                  `第${departmentColumnIndex + 1}列`;
                const departmentData = jsonData.map((row, index) => ({
                  行号: index + 1,
                  [processor.departmentKeyword]: row[departmentColumnIndex]
                }));

                console.log(
                  `\n========== "${processor.departmentKeyword}"列信息 ==========`
                );
                console.log(
                  `找到"${processor.departmentKeyword}"单元格位置: 第3行，第${departmentColumnIndex + 1}列`
                );
                console.log(`对应表头列名: ${departmentColumnName}`);
                console.log(`该列完整数据:`, departmentData);

                // 过滤掉空值，只显示有数据的行
                const validDepartmentData = departmentData.filter(
                  item =>
                    item[processor.departmentKeyword] &&
                    item[processor.departmentKeyword].toString().trim() !== ""
                );
                console.log(
                  `\n有效数据（共${validDepartmentData.length}条）:`,
                  validDepartmentData
                );

                // 对行数大于等于4的有效数据进行分组处理
                const validDataFromRow4 = validDepartmentData.filter(
                  item => item.行号 >= 4
                );

                if (validDataFromRow4.length > 0) {
                  console.log(
                    `\n========== 分组处理结果（第4行起数据） ==========`
                  );

                  // 分组处理
                  const groups = new Map<string, typeof validDataFromRow4>();

                  validDataFromRow4.forEach(item => {
                    const department =
                      item[processor.departmentKeyword].toString();
                    const firstPart = department.split("-")[0];

                    if (!groups.has(firstPart)) {
                      groups.set(firstPart, []);
                    }
                    groups.get(firstPart)!.push(item);
                  });

                  // 打印分组结果
                  console.log(`${processor.name} - 共分为 ${groups.size} 组:`);

                  groups.forEach((groupItems, groupName) => {
                    console.log(`\n--- 组名: ${groupName} ---`);
                    console.log(`包含数据条数: ${groupItems.length}`);
                    console.log(`具体数据:`);
                    groupItems.forEach(item => {
                      console.log(
                        `  行${item.行号}: ${item[processor.departmentKeyword]}`
                      );
                    });
                  });

                  // 生成分组统计
                  console.log(`\n========== 分组统计 ==========`);
                  const groupStats = Array.from(groups.entries()).map(
                    ([name, items]) => ({
                      组名: name,
                      数据条数: items.length,
                      行号范围: `${Math.min(...items.map(i => i.行号))}-${Math.max(...items.map(i => i.行号))}`
                    })
                  );
                  console.log(`${processor.name} - 分组统计表:`, groupStats);
                } else {
                  console.log(
                    `\n${processor.name} - 没有第4行及以后的有效数据进行分组处理`
                  );
                }
              } else {
                console.log(`${processor.name} - 第三行数据:`, thirdRow);
                console.log(`${processor.name} - 所有表头:`, headers);
                console.warn(
                  `在第三行中未找到'${processor.departmentKeyword}'单元格`
                );
              }
            }

            processedSheets++;

            // 当所有工作表都处理完成后显示结果
            if (processedSheets === totalSheets) {
              allSheetData.value = sheetData;
              excelData.value = sheetData.hotel || []; // 优先显示酒店数据
              originalWorkbook.value = workbook;
              showData.value = true;
              loading.value = false;

              ElMessage.success(
                `成功读取 ${totalSheets} 个工作表！请在控制台查看详细信息`
              );
            }
          });
        })
        .catch(error => {
          console.error("读取Excel文件失败:", error);
          ElMessage.error("读取Excel文件失败");
          loading.value = false;
        });
    } catch (error) {
      console.error("读取Excel文件失败:", error);
      ElMessage.error("读取Excel文件失败");
      loading.value = false;
    }
  };

  reader.onerror = () => {
    console.error("文件读取失败");
    ElMessage.error("文件读取失败");
    loading.value = false;
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

// 获取分组信息（按公司名称汇总）
const getGroupInfo = () => {
  const companyGroups = new Map<
    string,
    {
      groupName: string;
      hotelInfo?: { count: number; rowRange: string };
      internationalHotelInfo?: { count: number; rowRange: string };
      trainInfo?: { count: number; rowRange: string };
      flightInfo?: { count: number; rowRange: string };
      internationalFlightInfo?: { count: number; rowRange: string };
      totalCount: number;
      editableFileName: string;
    }
  >();

  // 处理酒店数据
  if (allSheetData.value.hotel && allSheetData.value.hotel.length > 0) {
    const hotelGroups = processSheetData(
      allSheetData.value.hotel,
      "酒店明细(国内)",
      "预订人部门",
      "hotel"
    );

    hotelGroups.forEach(group => {
      if (!companyGroups.has(group.groupName)) {
        const generatedFileName = generateFileName(group.groupName);
        companyGroups.set(group.groupName, {
          groupName: group.groupName,
          totalCount: 0,
          editableFileName: generatedFileName
        });
      }

      const companyGroup = companyGroups.get(group.groupName)!;
      companyGroup.hotelInfo = {
        count: group.count,
        rowRange: group.rowRange
      };
      companyGroup.totalCount += group.count;
    });
  }

  // 处理国际酒店数据
  if (
    allSheetData.value.internationalHotel &&
    allSheetData.value.internationalHotel.length > 0
  ) {
    const internationalHotelGroups = processSheetData(
      allSheetData.value.internationalHotel,
      "酒店明细(国际)",
      "预订人部门",
      "internationalHotel"
    );

    internationalHotelGroups.forEach(group => {
      if (!companyGroups.has(group.groupName)) {
        const existing = editableFileNames.value.find(
          item => item.groupName === group.groupName
        );
        companyGroups.set(group.groupName, {
          groupName: group.groupName,
          totalCount: 0,
          editableFileName: generateFileName(
            group.groupName,
            existing?.fileName
          )
        });
      }

      const companyGroup = companyGroups.get(group.groupName)!;
      companyGroup.internationalHotelInfo = {
        count: group.count,
        rowRange: group.rowRange
      };
      companyGroup.totalCount += group.count;
    });
  }

  // 处理火车票数据
  if (allSheetData.value.train && allSheetData.value.train.length > 0) {
    const trainGroups = processSheetData(
      allSheetData.value.train,
      "火车票明细",
      "预订人部门",
      "train"
    );

    trainGroups.forEach(group => {
      if (!companyGroups.has(group.groupName)) {
        const existing = editableFileNames.value.find(
          item => item.groupName === group.groupName
        );
        companyGroups.set(group.groupName, {
          groupName: group.groupName,
          totalCount: 0,
          editableFileName: generateFileName(
            group.groupName,
            existing?.fileName
          )
        });
      }

      const companyGroup = companyGroups.get(group.groupName)!;
      companyGroup.trainInfo = {
        count: group.count,
        rowRange: group.rowRange
      };
      companyGroup.totalCount += group.count;
    });
  }

  // 处理机票数据
  if (allSheetData.value.flight && allSheetData.value.flight.length > 0) {
    const flightGroups = processSheetData(
      allSheetData.value.flight,
      "机票明细(国内)",
      "预订人部门",
      "flight"
    );

    flightGroups.forEach(group => {
      if (!companyGroups.has(group.groupName)) {
        const existing = editableFileNames.value.find(
          item => item.groupName === group.groupName
        );
        companyGroups.set(group.groupName, {
          groupName: group.groupName,
          totalCount: 0,
          editableFileName: generateFileName(
            group.groupName,
            existing?.fileName
          )
        });
      }

      const companyGroup = companyGroups.get(group.groupName)!;
      companyGroup.flightInfo = {
        count: group.count,
        rowRange: group.rowRange
      };
      companyGroup.totalCount += group.count;
    });
  }

  // 处理国际机票数据
  if (
    allSheetData.value.internationalFlight &&
    allSheetData.value.internationalFlight.length > 0
  ) {
    const internationalFlightGroups = processSheetData(
      allSheetData.value.internationalFlight,
      "机票明细(国际)",
      "预订人部门",
      "internationalFlight"
    );

    internationalFlightGroups.forEach(group => {
      if (!companyGroups.has(group.groupName)) {
        const existing = editableFileNames.value.find(
          item => item.groupName === group.groupName
        );
        companyGroups.set(group.groupName, {
          groupName: group.groupName,
          totalCount: 0,
          editableFileName: generateFileName(
            group.groupName,
            existing?.fileName
          )
        });
      }

      const companyGroup = companyGroups.get(group.groupName)!;
      companyGroup.internationalFlightInfo = {
        count: group.count,
        rowRange: group.rowRange
      };
      companyGroup.totalCount += group.count;
    });
  }

  const result = Array.from(companyGroups.values());
  return result;
};

// 处理单个工作表数据
const processSheetData = (
  sheetData: any[][],
  sheetName: string,
  departmentKeyword: string,
  sheetType: string
) => {
  if (!sheetData || sheetData.length === 0) {
    return [];
  }

  // 重新计算分组数据
  const departmentColumnIndex = (sheetData[2] as any[]).findIndex(
    (cell: any) => cell && cell.toString().includes(departmentKeyword)
  );

  if (departmentColumnIndex === -1) {
    console.warn(`在${sheetName}中未找到${departmentKeyword}列`);
    return [];
  }

  const validDataFromRow4 = sheetData
    .slice(3)
    .filter((row: any[], index: number) => {
      const departmentValue = row[departmentColumnIndex];
      if (!departmentValue || departmentValue.toString().trim() === "") {
        return false;
      }

      const departmentText = departmentValue.toString();

      // 过滤掉合计行、总计行等非数据行
      const summaryKeywords = [
        "合计",
        "总计",
        "总和",
        "小计",
        "sum",
        "total",
        "summary"
      ];
      const isSummaryRow = summaryKeywords.some(keyword =>
        departmentText.toLowerCase().includes(keyword.toLowerCase())
      );

      // 过滤纯数字（可能是金额合计）
      const isPureNumber = /^\d+(\.\d+)?$/.test(departmentText.trim());

      // 过滤空值、特殊字符
      const isEmptyOrSpecial =
        departmentText.trim() === "" ||
        /^[\-_=+]+$/.test(departmentText.trim()) ||
        departmentText.length < 2;

      // 如果是空值、合计行、纯数字或特殊字符，不进行分组处理
      if (isSummaryRow || isPureNumber || isEmptyOrSpecial) {
        console.log(
          `跳过非数据行: 行号${index + 4}, 内容: ${departmentText}, 类型: ${
            isSummaryRow ? "合计行" : isPureNumber ? "纯数字" : "空值/特殊字符"
          }`
        );
        return false;
      }

      return true;
    })
    .map((row: any[], index: number) => ({
      行号: index + 4,
      [departmentKeyword]: row[departmentColumnIndex],
      完整行数据: row
    }));

  // 分组处理
  const groups = new Map<string, typeof validDataFromRow4>();
  validDataFromRow4.forEach(item => {
    const department = item[departmentKeyword].toString();

    // 只有包含"-"的部门信息才进行拆分，否则使用完整的部门名称
    let groupName: string;
    if (department.includes("-")) {
      groupName = department.split("-")[0].trim();
    } else {
      groupName = department.trim();
    }

    // 确保分组名称不为空
    if (groupName) {
      if (!groups.has(groupName)) {
        groups.set(groupName, []);
      }
      groups.get(groupName)!.push(item);
    }
  });

  // 返回分组信息
  const groupInfo = Array.from(groups.entries()).map(([name, items]) => {
    // 查找是否已有该分组的编辑文件名（区分工作表类型）
    const existing = editableFileNames.value.find(
      item => item.groupName === name && item.sheetType === sheetType
    );
    return {
      groupName: name,
      sheetName: sheetName,
      sheetType: sheetType,
      count: items.length,
      rowRange: `${Math.min(...items.map(i => i.行号))}-${Math.max(...items.map(i => i.行号))}`,
      fileName: `${name}.xlsx`,
      editableFileName: existing ? existing.fileName : name // 使用已保存的文件名或默认名称
    };
  });

  // 如果有新的分组，添加到editableFileNames中
  groupInfo.forEach(item => {
    const exists = editableFileNames.value.find(
      f => f.groupName === item.groupName && f.sheetType === item.sheetType
    );
    if (!exists) {
      editableFileNames.value.push({
        groupName: item.groupName,
        fileName: item.editableFileName,
        sheetType: item.sheetType
      });
    }
  });

  return groupInfo;
};

// 获取分组数量
const getGroupCount = () => {
  return getGroupInfo().length;
};

// 更新文件名（不再区分工作表类型）
const updateFileName = (
  groupName: string,
  newFileName: string,
  sheetType?: string
) => {
  const existing = editableFileNames.value.find(
    item => item.groupName === groupName
  );
  if (existing) {
    existing.fileName = newFileName;
  } else {
    editableFileNames.value.push({
      groupName: groupName,
      fileName: newFileName
    });
  }
};

// 应用工作表样式
const applyWorksheetStyling = async (
  worksheet: ExcelJS.Worksheet,
  data: any[][],
  departmentKeyword: string
) => {
  // 写入数据
  data.forEach((row, rowIndex) => {
    row.forEach((cellValue, colIndex) => {
      const cell = worksheet.getCell(rowIndex + 1, colIndex + 1);
      cell.value = cellValue;

      if (rowIndex < 3) {
        // 前三行应用居中对齐、边框和字体样式
        cell.alignment = {
          horizontal: "center",
          vertical: "middle"
        };

        cell.border = {
          top: { style: "thin" },
          bottom: { style: "thin" },
          left: { style: "thin" },
          right: { style: "thin" }
        };

        // 设置字体样式
        if (rowIndex === 0) {
          cell.font = { name: "微软雅黑", size: 16, bold: true };
        } else if (rowIndex === 1) {
          cell.font = { name: "微软雅黑", size: 11, bold: true };
        } else if (rowIndex === 2) {
          cell.font = { name: "微软雅黑", size: 11, bold: true };
        }
      } else {
        // 第四行及以后应用居中对齐和边框样式
        cell.alignment = {
          horizontal: "center",
          vertical: "middle"
        };

        cell.border = {
          top: { style: "thin" },
          bottom: { style: "thin" },
          left: { style: "thin" },
          right: { style: "thin" }
        };
      }
    });
  });

  // 检查前三行中相邻的单元格，如果值相等就合并
  for (let rowIndex = 0; rowIndex < 3; rowIndex++) {
    let startCol = 0;
    for (let colIndex = 1; colIndex < data[rowIndex].length; colIndex++) {
      const currentValue = data[rowIndex][colIndex];
      const previousValue = data[rowIndex][startCol];

      if (
        currentValue &&
        previousValue &&
        currentValue.toString() === previousValue.toString()
      ) {
        continue;
      } else {
        if (colIndex - 1 > startCol) {
          worksheet.mergeCells(
            rowIndex + 1,
            startCol + 1,
            rowIndex + 1,
            colIndex
          );
          const mergedCell = worksheet.getCell(rowIndex + 1, startCol + 1);
          mergedCell.alignment = { horizontal: "center", vertical: "middle" };
        }
        startCol = colIndex;
      }
    }

    if (data[rowIndex].length - 1 > startCol) {
      worksheet.mergeCells(
        rowIndex + 1,
        startCol + 1,
        rowIndex + 1,
        data[rowIndex].length
      );
      const mergedCell = worksheet.getCell(rowIndex + 1, startCol + 1);
      mergedCell.alignment = { horizontal: "center", vertical: "middle" };
    }
  }

  // 设置行高
  worksheet.eachRow((row, rowNumber) => {
    row.height = 40;
  });

  // 查找第三行中"应付金额"所在的列
  let paymentAmountColIndex = -1;
  const thirdRow = data[2];
  if (thirdRow) {
    thirdRow.forEach((cellValue, index) => {
      if (cellValue && cellValue.toString() === "应付金额") {
        paymentAmountColIndex = index;
      }
    });
  }

  // 对应付金额列进行求和并添加合计行
  if (paymentAmountColIndex !== -1 && data.length > 3) {
    let sum = 0;
    for (let i = 3; i < data.length; i++) {
      const value = data[i][paymentAmountColIndex];
      if (value !== null && value !== undefined && value !== "") {
        const numValue = parseFloat(value.toString());
        if (!isNaN(numValue)) {
          sum += numValue;
        }
      }
    }

    const summaryRow = new Array(data[0].length).fill("");
    summaryRow[0] = "合计";
    summaryRow[paymentAmountColIndex] = sum;

    const summaryRowIndex = data.length + 1;
    summaryRow.forEach((cellValue, colIndex) => {
      const cell = worksheet.getCell(summaryRowIndex, colIndex + 1);
      cell.value = cellValue;

      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = {
        top: { style: "thin" },
        bottom: { style: "thin" },
        left: { style: "thin" },
        right: { style: "thin" }
      };

      if (colIndex === 0 || colIndex === paymentAmountColIndex) {
        cell.font = { name: "微软雅黑", size: 10, bold: true };
      }
    });

    worksheet.getRow(summaryRowIndex).height = 40;
  }

  // 自动调整列宽
  worksheet.columns.forEach((column, index) => {
    let maxLength = 0;
    column.eachCell((cell, rowNumber) => {
      if (cell.value) {
        const text = cell.value.toString();
        const charWidth = text.split("").reduce((width, char) => {
          return width + (char.charCodeAt(0) > 127 ? 2 : 1);
        }, 0);
        if (charWidth > maxLength) {
          maxLength = charWidth;
        }
      }
    });
    column.width = Math.max(maxLength * 1.1, 15);
  });
};

// 导入公司配置
import companyConfig from "./companyConfig";

// 生成上个月日期范围字符串（处理上个月的账单）
const generateCurrentMonthDateRange = (): string => {
  const now = new Date();
  const currentYear = now.getFullYear();
  const currentMonth = now.getMonth(); // 0-11
  // 计算上个月的年月
  let targetYear, targetMonth;

  if (currentMonth === 0) {
    // 当前是1月，上个月是去年的12月
    targetYear = currentYear - 1;
    targetMonth = 12; // 12月
  } else {
    targetYear = currentYear;
    targetMonth = currentMonth; // 因为getMonth()返回0-11，所以直接使用
  }

  // 计算上个月的最后一天：当前月份第1天减去1天
  const currentMonthFirstDay = new Date(currentYear, currentMonth, 1);
  const lastDayOfTargetMonth = new Date(
    currentMonthFirstDay.getTime() - 24 * 60 * 60 * 1000
  );

  // 格式化为 M.D-M.D 格式（如 8.1-8.31，12.1-12.31）
  const startDate = `${targetMonth}.1`;
  const endDate = `${targetMonth}.${lastDayOfTargetMonth.getDate()}`;
  const dateRange = `${startDate}-${endDate}`;

  return dateRange;
};

// 生成带日期的文件名
const generateFileName = (
  groupName: string,
  existingFileName?: string
): string => {
  const companyInfo = companyConfig.getCompanyInfo(groupName);

  const dateRange = generateCurrentMonthDateRange();

  // 优先使用 shortName + 日期
  const finalFileName = `${companyInfo.shortName}${dateRange}`;

  return finalFileName;
};

// 生成分组Excel文件并打包成ZIP
const generateGroupedExcelFiles = async () => {
  if (!originalWorkbook.value || Object.keys(allSheetData.value).length === 0) {
    ElMessage.error("请先上传并处理Excel文件");
    return;
  }

  generating.value = true;

  try {
    const groupInfo = getGroupInfo();
    console.log(`准备为 ${groupInfo.length} 个公司生成Excel文件`);

    // 创建ZIP文件
    const zip = new JSZip();

    // 预加载结算单模板
    let templateWorkbook: ExcelJS.Workbook;
    try {
      console.log("开始加载结算单模板...");

      // 尝试多个可能的路径
      const possiblePaths = [
        "./cxjg/结算单.xlsx", // public目录下的文件
        "/cxjg/结算单.xlsx", // 根路径访问
        "cxjg/结算单.xlsx" // 相对路径
      ];

      let templateBuffer: ArrayBuffer | null = null;
      let successPath = "";

      for (const path of possiblePaths) {
        try {
          console.log(`尝试路径: ${path}`);
          const templateResponse = await fetch(path);

          if (templateResponse.ok) {
            console.log(
              `文件找到，大小: ${templateResponse.headers.get("content-length")} bytes`
            );
            templateBuffer = await templateResponse.arrayBuffer();
            successPath = path;
            break;
          } else {
            console.log(`路径 ${path} 返回状态: ${templateResponse.status}`);
          }
        } catch (pathError) {
          console.log(`路径 ${path} 访问失败:`, pathError);
        }
      }

      if (!templateBuffer) {
        throw new Error(
          `结算单模板文件未找到，请确保模板文件存在于以下任一路径:\n${possiblePaths.join("\n")}`
        );
      }

      console.log(
        `成功加载模板文件: ${successPath}, 大小: ${templateBuffer.byteLength} bytes`
      );

      // 验证文件格式
      if (templateBuffer.byteLength < 100) {
        throw new Error("模板文件太小，可能不是有效的Excel文件");
      }

      templateWorkbook = new ExcelJS.Workbook();
      await templateWorkbook.xlsx.load(templateBuffer);

      console.log(
        "成功解析结算单模板，工作表数量:",
        templateWorkbook.worksheets.length
      );

      // 验证结算单工作表是否存在
      const settlementWorksheet = templateWorkbook.getWorksheet("结算单");
      if (!settlementWorksheet) {
        throw new Error('模板文件中没有找到"结算单"工作表，请检查模板文件');
      }

      console.log("成功找到结算单工作表，可以开始使用");
    } catch (error) {
      console.error("加载结算单模板失败:", error);
      let errorMessage = "加载结算单模板失败";

      if (error instanceof Error) {
        if (error.message.includes("Can't find end of central directory")) {
          errorMessage =
            "结算单模板文件格式错误或文件损坏，请确保是有效的Excel文件(.xlsx格式)";
        } else {
          errorMessage = `加载结算单模板失败: ${error.message}`;
        }
      }

      ElMessage.error(errorMessage);
      generating.value = false;
      return;
    }

    // 为每个公司生成Excel文件
    for (const companyGroup of groupInfo) {
      // 始终使用 shortName + 日期生成最新的文件名
      const latestFileName = generateFileName(companyGroup.groupName);
      console.log(
        `生成文件: ${latestFileName}.xlsx，公司: ${companyGroup.groupName}`
      );

      // 创建新的工作簿
      const newWorkbook = new ExcelJS.Workbook();

      // 添加结算单工作表（使用模板）
      const templateWorksheet = templateWorkbook.getWorksheet("结算单");

      if (!templateWorksheet) {
        throw new Error("无法获取模板的结算单工作表");
      }

      // 获取公司配置信息
      const companyInfo = companyConfig.getCompanyInfo(companyGroup.groupName);
      console.log(`公司配置信息:`, companyInfo);

      // 打印结算单模板数据
      console.log(`===== 结算单模板数据内容 (${companyGroup.groupName}) =====`);
      const templateData: any[][] = [];
      templateWorksheet.eachRow((row, rowNumber) => {
        const rowData: any[] = [];
        row.eachCell((cell, colNumber) => {
          rowData[colNumber] = cell.value;
        });
        templateData[rowNumber] = rowData;

        // 打印有内容的行
        if (
          rowData.some(
            cell => cell !== null && cell !== undefined && cell !== ""
          )
        ) {
          console.log(`第${rowNumber}行:`, rowData);
        }
      });
      console.log(`===== 结算单模板数据结束，共${templateData.length}行 =====`);

      // 单独打印第九行数据
      if (templateData[9]) {
        console.log(`===== 第九行数据单独打印 =====`);
        console.log(`第九行数据:`, templateData[9]);
        console.log(`第九行详细信息:`);
        templateData[9].forEach((cellValue, index) => {
          if (
            cellValue !== null &&
            cellValue !== undefined &&
            cellValue !== ""
          ) {
            console.log(`  第${index}列: ${cellValue}`);
          }
        });
        console.log(`===== 第九行数据打印结束 =====`);
      } else {
        console.log("第九行数据不存在或为空");
      }

      const summaryWorksheet = newWorkbook.addWorksheet("结算单");

      // 复制模板结构和样式
      templateWorksheet.eachRow((row, rowNumber) => {
        row.eachCell((cell, colNumber) => {
          const newCell = summaryWorksheet.getCell(rowNumber, colNumber);

          // 复制原始值，并替换公司名称、联系人和日期
          let cellValue = cell.value;
          if (typeof cellValue === "string") {
            let hasChanges = false;
            const originalValue = cellValue;

            // 替换《》内的公司名称
            if (cellValue.includes("《") && cellValue.includes("》")) {
              cellValue = cellValue.replace(
                /《([^》]+)》/g,
                `《${companyGroup.groupName}》`
              );
              hasChanges = true;
            }

            // 替换"收件方："后面的公司名称
            if (cellValue.includes("收件方：")) {
              cellValue = cellValue.replace(
                /收件方：[^，,\n]+/,
                `收件方：${companyGroup.groupName}`
              );
              hasChanges = true;
            }

            // 替换联系人信息
            if (cellValue.includes("收件人：") && companyInfo.contact) {
              cellValue = cellValue.replace(
                /收件人：[^，,\n]+/,
                `收件人：${companyInfo.contact}`
              );
              hasChanges = true;
            }

            // 替换手机号信息
            if (cellValue.includes("电话：") && companyInfo.phone) {
              // 获取当前电话
              const currentPhoneMatch = cellValue.match(/电话：([^，,\n]+)/);
              if (currentPhoneMatch) {
                const currentPhone = currentPhoneMatch[1];
                // 如果当前电话不等于 15768628831，则保留原电话并添加公司配置的电话
                if (currentPhone !== "15768628831") {
                  cellValue = cellValue.replace(
                    /电话：[^，,\n]+/,
                    `电话：${companyInfo.phone}`
                  );
                  hasChanges = true;
                }
              }
            }

            // 动态计算上个月的最后一天（处理上个月的账单）
            if (cellValue.includes("最晚结算日：")) {
              const now = new Date();
              const currentYear = now.getFullYear();
              const currentMonth = now.getMonth(); // 0-11

              // 计算上个月的年月
              let targetYear, targetMonth;

              if (currentMonth === 0) {
                // 当前是1月，上个月是去年的12月
                targetYear = currentYear - 1;
                targetMonth = 12;
              } else {
                targetYear = currentYear;
                targetMonth = currentMonth;
              }

              // 计算上个月的最后一天：当前月份第1天减去1天
              const currentMonthFirstDay = new Date(
                currentYear,
                currentMonth,
                1
              );
              const lastDayOfTargetMonth = new Date(
                currentMonthFirstDay.getTime() - 24 * 60 * 60 * 1000
              );

              // 格式化为 YYYY-MM-DD
              const formattedDate = `${targetYear}-${String(targetMonth).padStart(2, "0")}-${String(lastDayOfTargetMonth.getDate()).padStart(2, "0")}`;

              cellValue = cellValue.replace(
                /最晚结算日：\d{4}-\d{2}-\d{2}/,
                `最晚结算日：${formattedDate}`
              );
              hasChanges = true;
            }

            // 动态生成结算款项描述文本
            if (cellValue.includes("结算款项列示如下：")) {
              const now = new Date();
              const currentYear = now.getFullYear();
              const currentMonth = now.getMonth(); // 0-11

              // 计算上个月的年月（与上面逻辑一致）
              let targetYear, targetMonth;

              if (currentMonth === 0) {
                // 当前是1月，上个月是去年的12月
                targetYear = currentYear - 1;
                targetMonth = 12;
              } else {
                targetYear = currentYear;
                targetMonth = currentMonth;
              }

              // 计算上个月的最后一天
              const currentMonthFirstDay = new Date(
                currentYear,
                currentMonth,
                1
              );
              const lastDayOfTargetMonth = new Date(
                currentMonthFirstDay.getTime() - 24 * 60 * 60 * 1000
              );

              // 格式化日期
              const startDate = `${targetYear}-${String(targetMonth).padStart(2, "0")}-01`;
              const endDate = `${targetYear}-${String(targetMonth).padStart(2, "0")}-${String(lastDayOfTargetMonth.getDate()).padStart(2, "0")}`;

              // 替换结算款项描述
              cellValue = cellValue.replace(
                /本公司\d{4}-\d{2}-\d{2}至\d{4}-\d{2}-\d{2}与贵公司\([^)]+\)的结算款项列示如下：/,
                `本公司${startDate}至${endDate}与贵公司(${companyGroup.groupName})的结算款项列示如下：`
              );
              hasChanges = true;
            }

            // 动态生成付款到期日
            if (cellValue.includes("前，将本期应还金额付款到以下账户：")) {
              const now = new Date();
              const currentYear = now.getFullYear();
              const currentMonth = now.getMonth() + 1; // 1-12

              // 计算当月的最后一天
              const nextMonthFirstDay = new Date(currentYear, currentMonth, 1);
              const lastDayOfCurrentMonth = new Date(
                nextMonthFirstDay.getTime() - 24 * 60 * 60 * 1000
              );

              // 格式化为 YYYY-MM-DD
              const formattedDate = `${currentYear}-${String(currentMonth).padStart(2, "0")}-${String(lastDayOfCurrentMonth.getDate()).padStart(2, "0")}`;

              // 替换付款到期日
              cellValue = cellValue.replace(
                /请在\d{4}-\d{2}-\d{2}前，将本期应还金额付款到以下账户：/,
                `请在${formattedDate}前，将本期应还金额付款到以下账户：`
              );
              hasChanges = true;
              console.log(`动态生成付款到期日: ${formattedDate}`);
            }

            // if (hasChanges) {
            //   console.log(
            //     `第${rowNumber}行第${colNumber}列替换内容:`,
            //     originalValue,
            //     "->",
            //     cellValue
            //   );
            // }
          }

          newCell.value = cellValue;

          // 复制样式
          newCell.style = cell.style;
        });
      });

      // 复制行高
      templateWorksheet.eachRow((row, rowNumber) => {
        if (row.height) {
          summaryWorksheet.getRow(rowNumber).height = row.height;
        }
      });

      // 复制列宽
      templateWorksheet.columns.forEach((column, index) => {
        if (column.width) {
          summaryWorksheet.getColumn(index + 1).width = column.width;
        }
      });

      // 复制合并单元格
      templateWorksheet.model.merges?.forEach((merge: any) => {
        summaryWorksheet.mergeCells(merge);
      });

      console.log(`使用结算单模板为 ${companyGroup.groupName} 生成工作表`);

      // 检查是否有酒店数据并添加工作表
      if (companyGroup.hotelInfo && allSheetData.value.hotel) {
        const hotelData = allSheetData.value.hotel;
        const departmentKeyword = "预订人部门";
        const departmentColumnIndex = (hotelData[2] as any[]).findIndex(
          (cell: any) => cell && cell.toString().includes(departmentKeyword)
        );

        if (departmentColumnIndex !== -1) {
          const newWorksheet = newWorkbook.addWorksheet("国内酒店", {
            views: [{ showGridLines: true }]
          });
          newWorksheet.properties.defaultRowHeight = 40;

          // 筛选该公司的酒店数据
          const companyHotelData = hotelData
            .slice(3)
            .filter((row: any[], index: number) => {
              const departmentValue = row[departmentColumnIndex];
              if (!departmentValue) return false;

              const departmentText = departmentValue.toString();

              // 过滤掉合计行、总计行等非数据行
              const summaryKeywords = [
                "合计",
                "总计",
                "总和",
                "小计",
                "sum",
                "total",
                "summary"
              ];
              const isSummaryRow = summaryKeywords.some(keyword =>
                departmentText.toLowerCase().includes(keyword.toLowerCase())
              );

              // 过滤纯数字（可能是金额合计）
              const isPureNumber = /^\d+(\.\d+)?$/.test(departmentText.trim());

              // 过滤空值、特殊字符
              const isEmptyOrSpecial =
                departmentText.trim() === "" ||
                /^[\-_=+]+$/.test(departmentText.trim()) ||
                departmentText.length < 2;

              if (isSummaryRow || isPureNumber || isEmptyOrSpecial)
                return false;

              // 获取分组名称进行匹配
              let groupName: string;
              if (departmentText.includes("-")) {
                groupName = departmentText.split("-")[0].trim();
              } else {
                groupName = departmentText.trim();
              }

              return groupName === companyGroup.groupName;
            });

          // 复制原始前三行
          const headerRows = hotelData.slice(0, 3);
          const newData = [...headerRows, ...companyHotelData];

          console.log(`  酒店工作表: ${newData.length} 行数据`);

          // 应用样式和格式
          await applyWorksheetStyling(newWorksheet, newData, departmentKeyword);
        }
      }

      // 检查是否有国际酒店数据并添加工作表
      if (
        companyGroup.internationalHotelInfo &&
        allSheetData.value.internationalHotel
      ) {
        const internationalHotelData = allSheetData.value.internationalHotel;
        const departmentKeyword = "预订人部门";
        const departmentColumnIndex = (
          internationalHotelData[2] as any[]
        ).findIndex(
          (cell: any) => cell && cell.toString().includes(departmentKeyword)
        );

        if (departmentColumnIndex !== -1) {
          const newWorksheet = newWorkbook.addWorksheet("国际酒店", {
            views: [{ showGridLines: true }]
          });
          newWorksheet.properties.defaultRowHeight = 40;

          // 筛选该公司的国际酒店数据
          const companyInternationalHotelData = internationalHotelData
            .slice(3)
            .filter((row: any[], index: number) => {
              const departmentValue = row[departmentColumnIndex];
              if (!departmentValue) return false;

              const departmentText = departmentValue.toString();

              // 过滤掉合计行、总计行等非数据行
              const summaryKeywords = [
                "合计",
                "总计",
                "总和",
                "小计",
                "sum",
                "total",
                "summary"
              ];
              const isSummaryRow = summaryKeywords.some(keyword =>
                departmentText.toLowerCase().includes(keyword.toLowerCase())
              );

              // 过滤纯数字（可能是金额合计）
              const isPureNumber = /^\d+(\.\d+)?$/.test(departmentText.trim());

              // 过滤空值、特殊字符
              const isEmptyOrSpecial =
                departmentText.trim() === "" ||
                /^[\-_=+]+$/.test(departmentText.trim()) ||
                departmentText.length < 2;

              if (isSummaryRow || isPureNumber || isEmptyOrSpecial)
                return false;

              // 获取分组名称进行匹配
              let groupName: string;
              if (departmentText.includes("-")) {
                groupName = departmentText.split("-")[0].trim();
              } else {
                groupName = departmentText.trim();
              }

              return groupName === companyGroup.groupName;
            });

          // 复制原始前三行
          const headerRows = internationalHotelData.slice(0, 3);
          const newData = [...headerRows, ...companyInternationalHotelData];

          console.log(`  国际酒店工作表: ${newData.length} 行数据`);

          // 应用样式和格式
          await applyWorksheetStyling(newWorksheet, newData, departmentKeyword);
        }
      }

      // 检查是否有火车票数据并添加工作表
      if (companyGroup.trainInfo && allSheetData.value.train) {
        const trainData = allSheetData.value.train;
        const departmentKeyword = "预订人部门";
        const departmentColumnIndex = (trainData[2] as any[]).findIndex(
          (cell: any) => cell && cell.toString().includes(departmentKeyword)
        );

        if (departmentColumnIndex !== -1) {
          const newWorksheet = newWorkbook.addWorksheet("火车票", {
            views: [{ showGridLines: true }]
          });
          newWorksheet.properties.defaultRowHeight = 40;

          // 筛选该公司的火车票数据
          const companyTrainData = trainData
            .slice(3)
            .filter((row: any[], index: number) => {
              const departmentValue = row[departmentColumnIndex];
              if (!departmentValue) return false;

              const departmentText = departmentValue.toString();

              // 过滤掉合计行、总计行等非数据行
              const summaryKeywords = [
                "合计",
                "总计",
                "总和",
                "小计",
                "sum",
                "total",
                "summary"
              ];
              const isSummaryRow = summaryKeywords.some(keyword =>
                departmentText.toLowerCase().includes(keyword.toLowerCase())
              );

              // 过滤纯数字（可能是金额合计）
              const isPureNumber = /^\d+(\.\d+)?$/.test(departmentText.trim());

              // 过滤空值、特殊字符
              const isEmptyOrSpecial =
                departmentText.trim() === "" ||
                /^[\-_=+]+$/.test(departmentText.trim()) ||
                departmentText.length < 2;

              if (isSummaryRow || isPureNumber || isEmptyOrSpecial)
                return false;

              // 获取分组名称进行匹配
              let groupName: string;
              if (departmentText.includes("-")) {
                groupName = departmentText.split("-")[0].trim();
              } else {
                groupName = departmentText.trim();
              }

              return groupName === companyGroup.groupName;
            });

          // 复制原始前三行
          const headerRows = trainData.slice(0, 3);
          const newData = [...headerRows, ...companyTrainData];

          console.log(`  火车票工作表: ${newData.length} 行数据`);

          // 应用样式和格式
          await applyWorksheetStyling(newWorksheet, newData, departmentKeyword);
        }
      }

      // 检查是否有机票数据并添加工作表
      if (companyGroup.flightInfo && allSheetData.value.flight) {
        const flightData = allSheetData.value.flight;
        const departmentKeyword = "预订人部门";
        const departmentColumnIndex = (flightData[2] as any[]).findIndex(
          (cell: any) => cell && cell.toString().includes(departmentKeyword)
        );

        if (departmentColumnIndex !== -1) {
          const newWorksheet = newWorkbook.addWorksheet("国内机票", {
            views: [{ showGridLines: true }]
          });
          newWorksheet.properties.defaultRowHeight = 40;

          // 筛选该公司的机票数据
          const companyFlightData = flightData
            .slice(3)
            .filter((row: any[], index: number) => {
              const departmentValue = row[departmentColumnIndex];
              if (!departmentValue) return false;

              const departmentText = departmentValue.toString();

              // 过滤掉合计行、总计行等非数据行
              const summaryKeywords = [
                "合计",
                "总计",
                "总和",
                "小计",
                "sum",
                "total",
                "summary"
              ];
              const isSummaryRow = summaryKeywords.some(keyword =>
                departmentText.toLowerCase().includes(keyword.toLowerCase())
              );

              // 过滤纯数字（可能是金额合计）
              const isPureNumber = /^\d+(\.\d+)?$/.test(departmentText.trim());

              // 过滤空值、特殊字符
              const isEmptyOrSpecial =
                departmentText.trim() === "" ||
                /^[\-_=+]+$/.test(departmentText.trim()) ||
                departmentText.length < 2;

              if (isSummaryRow || isPureNumber || isEmptyOrSpecial)
                return false;

              // 获取分组名称进行匹配
              let groupName: string;
              if (departmentText.includes("-")) {
                groupName = departmentText.split("-")[0].trim();
              } else {
                groupName = departmentText.trim();
              }

              return groupName === companyGroup.groupName;
            });

          // 复制原始前三行
          const headerRows = flightData.slice(0, 3);
          const newData = [...headerRows, ...companyFlightData];

          console.log(`  机票工作表: ${newData.length} 行数据`);

          // 应用样式和格式
          await applyWorksheetStyling(newWorksheet, newData, departmentKeyword);
        }
      }

      // 检查是否有国际机票数据并添加工作表
      if (
        companyGroup.internationalFlightInfo &&
        allSheetData.value.internationalFlight
      ) {
        const internationalFlightData = allSheetData.value.internationalFlight;
        const departmentKeyword = "预订人部门";
        const departmentColumnIndex = (
          internationalFlightData[2] as any[]
        ).findIndex(
          (cell: any) => cell && cell.toString().includes(departmentKeyword)
        );

        if (departmentColumnIndex !== -1) {
          const newWorksheet = newWorkbook.addWorksheet("国际机票", {
            views: [{ showGridLines: true }]
          });
          newWorksheet.properties.defaultRowHeight = 40;

          // 筛选该公司的国际机票数据
          const companyInternationalFlightData = internationalFlightData
            .slice(3)
            .filter((row: any[], index: number) => {
              const departmentValue = row[departmentColumnIndex];
              if (!departmentValue) return false;

              const departmentText = departmentValue.toString();

              // 过滤掉合计行、总计行等非数据行
              const summaryKeywords = [
                "合计",
                "总计",
                "总和",
                "小计",
                "sum",
                "total",
                "summary"
              ];
              const isSummaryRow = summaryKeywords.some(keyword =>
                departmentText.toLowerCase().includes(keyword.toLowerCase())
              );

              // 过滤纯数字（可能是金额合计）
              const isPureNumber = /^\d+(\.\d+)?$/.test(departmentText.trim());

              // 过滤空值、特殊字符
              const isEmptyOrSpecial =
                departmentText.trim() === "" ||
                /^[\-_=+]+$/.test(departmentText.trim()) ||
                departmentText.length < 2;

              if (isSummaryRow || isPureNumber || isEmptyOrSpecial)
                return false;

              // 获取分组名称进行匹配
              let groupName: string;
              if (departmentText.includes("-")) {
                groupName = departmentText.split("-")[0].trim();
              } else {
                groupName = departmentText.trim();
              }

              return groupName === companyGroup.groupName;
            });

          // 复制原始前三行
          const headerRows = internationalFlightData.slice(0, 3);
          const newData = [...headerRows, ...companyInternationalFlightData];

          console.log(`  国际机票工作表: ${newData.length} 行数据`);

          // 应用样式和格式
          await applyWorksheetStyling(newWorksheet, newData, departmentKeyword);
        }
      }

      // 辅助函数：从费用类型工作表中查找服务费列并求和
      const getServiceFeeTotal = (worksheetName: string): number | null => {
        console.log(`===== 开始查找 ${worksheetName} 的服务费总额 =====`);

        const worksheet = newWorkbook.getWorksheet(worksheetName);
        if (!worksheet) {
          console.log(`未找到工作表: ${worksheetName}`);
          return null;
        }

        // 获取工作表的所有数据
        const worksheetData: any[][] = [];
        worksheet.eachRow((row, rowNumber) => {
          const rowData: any[] = [];
          row.eachCell((cell, colNumber) => {
            rowData[colNumber] = cell.value;
          });
          worksheetData[rowNumber] = rowData;
        });

        console.log(`${worksheetName} 工作表共${worksheetData.length}行数据`);

        // 查找第三行中含有"服务费"的列
        if (worksheetData.length < 3) {
          console.log(`${worksheetName} 工作表数据不足3行`);
          return null;
        }

        const thirdRow = worksheetData[2]; // 第三行（索引2）
        console.log(`第三行数据:`, thirdRow);

        let serviceFeeColIndex = -1;
        for (let col = 1; col < thirdRow.length; col++) {
          const cellValue = thirdRow[col];
          if (cellValue && typeof cellValue === 'string' && cellValue.includes('服务费')) {
            serviceFeeColIndex = col;
            console.log(`在第${col + 1}列找到包含"服务费"的单元格: ${cellValue}`);
            break;
          }
        }

        if (serviceFeeColIndex === -1) {
          console.log(`在${worksheetName}的第三行中未找到包含"服务费"的列`);
          return null;
        }

        // 对该列的数据求和（从第四行开始，跳过标题行）
        let totalAmount = 0;
        console.log(`开始对第${serviceFeeColIndex + 1}列的数据求和（从第4行开始）`);

        for (let row = 3; row < worksheetData.length; row++) { // 从索引3开始（第4行）
          const rowData = worksheetData[row];
          const cellValue = rowData[serviceFeeColIndex];

          if (cellValue !== null && cellValue !== undefined) {
            let numericValue = 0;

            if (typeof cellValue === 'number') {
              numericValue = cellValue;
            } else if (typeof cellValue === 'string') {
              // 提取字符串中的数字
              const numberMatch = cellValue.match(/-?\d+\.?\d*/);
              if (numberMatch) {
                numericValue = parseFloat(numberMatch[0]);
              }
            }

            if (numericValue !== 0) {
              console.log(`  第${row + 1}行: ${cellValue} -> ${numericValue}`);
              totalAmount += numericValue;
            }
          }
        }

        console.log(`${worksheetName} 服务费总额: ${totalAmount}`);
        console.log(`===== ${worksheetName} 服务费总额查找结束 =====`);
        return totalAmount;
      };

      // 辅助函数：根据费用类型名称获取对应的工作表名
      const getWorksheetNameByExpenseType = (expenseType: string): string | null => {
        const expenseTypeMap: { [key: string]: string } = {
          '国内机票': '国内机票',
          '国内酒店': '国内酒店',
          '火车票': '火车票',
          '国际机票': '国际机票',
          '国际酒店': '国际酒店'
        };

        return expenseTypeMap[expenseType] || null;
      };

      // 在所有工作表处理完成后，处理第九行的总计金额替换
      console.log(`===== 开始处理第九行总计金额替换 =====`);

      // 通用的费用类型处理函数
      const processExpenseTypeRow = (
        rowIndex: number,
        rowDescription: string
      ) => {
        console.log(`===== 开始处理第${rowIndex}行 (${rowDescription}) =====`);

        if (!summaryWorksheet) {
          console.log(`未找到结算单工作表`);
          return;
        }

        const targetRow = summaryWorksheet.getRow(rowIndex);
        if (!targetRow) {
          console.log(`结算单工作表中没有第${rowIndex}行数据`);
          return;
        }

        console.log(`检查结算单第${rowIndex}行数据...`);

        // 查找费用类型名称（如"国内机票"、"国内酒店"等）
        let expenseTypeColIndex = -1;
        let expenseType = "";

        for (let col = 1; col <= 50; col++) {
          const cell = targetRow.getCell(col);
          const cellValue = cell.value;
          if (
            cellValue &&
            typeof cellValue === "string" &&
            ["国内机票", "国内酒店", "火车票", "国际机票", "国际酒店"].includes(
              cellValue
            )
          ) {
            expenseTypeColIndex = col;
            expenseType = cellValue;
            console.log(`在第${col}列找到费用类型: ${expenseType}`);
            break;
          }
        }

        if (!expenseType) {
          console.log(`第${rowIndex}行中未找到有效的费用类型`);
          return;
        }

        // 获取第六列的原始值
        const sixthColCell = targetRow.getCell(6);
        const originalValue = sixthColCell.value;
        console.log(`第${rowIndex}行第6列原始值: ${originalValue} (类型: ${typeof originalValue})`);

        // 根据费用类型获取工作表名
        const worksheetName = getWorksheetNameByExpenseType(expenseType);
        if (!worksheetName) {
          console.log(`未找到费用类型"${expenseType}"对应的工作表`);
          return;
        }

        console.log(`费用类型"${expenseType}"对应的工作表: ${worksheetName}`);

        // 从对应工作表中查找服务费总额
        const serviceFeeTotal = getServiceFeeTotal(worksheetName);

        if (serviceFeeTotal !== null) {
          console.log(`找到${expenseType}的服务费总额: ${serviceFeeTotal}`);

          // 替换第六列的值
          sixthColCell.value = serviceFeeTotal;
          console.log(
            `替换第${rowIndex}行第6列: ${originalValue} -> ${serviceFeeTotal}`
          );
        } else {
          console.log(`未找到${expenseType}的服务费总额，保持原值`);
        }

        console.log(`===== 第${rowIndex}行 (${rowDescription}) 处理结束 =====`);
      };

      // 处理第九行（国内机票）- 服务费求和
      processExpenseTypeRow(9, "国内机票");

      // 处理第九行（国内机票）- 总计金额替换（原有逻辑）
      console.log(`===== 开始处理第九行国内机票总计金额替换（原有逻辑） =====`);

      if (summaryWorksheet) {
        const ninthRow = summaryWorksheet.getRow(9);
        if (ninthRow) {
          console.log(`检查结算单第九行数据...`);

          // 遍历第九行的每个单元格，查找"国内机票"
          let domesticFlightColIndex = -1;
          for (let col = 1; col <= 50; col++) {
            const cell = ninthRow.getCell(col);
            if (cell.value === "国内机票") {
              domesticFlightColIndex = col;
              console.log(`在第${col}列找到"国内机票"`);
              break;
            }
          }

          if (domesticFlightColIndex !== -1) {
            console.log(`发现国内机票数据，开始从拆分好的工作表中查找合计金额`);

            const domesticFlightWorksheet = newWorkbook.getWorksheet("国内机票");
            if (domesticFlightWorksheet) {
              console.log(`找到拆分好的国内机票工作表`);

              // 获取工作表的所有数据
              const domesticFlightData: any[][] = [];
              domesticFlightWorksheet.eachRow((row, rowNumber) => {
                const rowData: any[] = [];
                row.eachCell((cell, colNumber) => {
                  rowData[colNumber] = cell.value;
                });
                domesticFlightData[rowNumber] = rowData;
              });

              console.log(`国内机票工作表共${domesticFlightData.length}行数据`);

              // 查找合计行
              let totalAmount = null;
              for (let i = domesticFlightData.length - 1; i >= 0; i--) {
                const row = domesticFlightData[i];
                const hasTotalText = row.some(
                  cell => typeof cell === "string" && cell.includes("合计")
                );
                const hasNumbers = row.some(
                  cell =>
                    typeof cell === "number" ||
                    (typeof cell === "string" && /^\d+\.?\d*$/.test(cell))
                );

                if (hasTotalText || hasNumbers) {
                  console.log(`找到合计行（第${i + 1}行）:`, row);

                  for (let j = row.length - 1; j >= 0; j--) {
                    const cell = row[j];
                    if (typeof cell === "number") {
                      totalAmount = cell;
                      break;
                    } else if (
                      typeof cell === "string" &&
                      /^\d+\.?\d*$/.test(cell)
                    ) {
                      totalAmount = parseFloat(cell);
                      break;
                    }
                  }
                  break;
                }
              }

              if (totalAmount !== null) {
                console.log(`找到合计金额: ${totalAmount}`);

                // 查找国内机票后面的所有数字列，找到最后一个数字列作为总计金额
                let allNumberCols: { col: number; value: any }[] = [];
                for (let searchCol = domesticFlightColIndex + 1; searchCol <= 20; searchCol++) {
                  const searchCell = ninthRow.getCell(searchCol);
                  if (searchCell.value !== null && searchCell.value !== undefined) {
                    if (typeof searchCell.value === "number") {
                      allNumberCols.push({ col: searchCol, value: searchCell.value });
                      console.log(`找到数字列: 第${searchCol}列 = ${searchCell.value}`);
                    } else if (typeof searchCell.value === "string" && /^\d+\.?\d*$/.test(searchCell.value)) {
                      allNumberCols.push({ col: searchCol, value: parseFloat(searchCell.value) });
                      console.log(`找到数字字符串列: 第${searchCol}列 = ${searchCell.value}`);
                    }
                  }
                }

                let targetColIndex = -1;
                if (allNumberCols.length > 0) {
                  const lastNumberCol = allNumberCols[allNumberCols.length - 1];
                  targetColIndex = lastNumberCol.col;
                  console.log(`选择最后一个数字列作为总计金额: 第${targetColIndex}列 = ${lastNumberCol.value}`);
                }

                if (targetColIndex !== -1) {
                  const targetCell = ninthRow.getCell(targetColIndex);
                  const oldValue = targetCell.value;
                  targetCell.value = totalAmount;
                  console.log(`替换第九行第${targetColIndex}列（总计金额）: ${oldValue} -> ${totalAmount}`);
                } else {
                  console.log(`未找到第九行中的总计金额列，请检查数据结构`);
                }
              } else {
                console.log(`未在拆分好的国内机票工作表中找到有效的合计金额`);
              }
            } else {
              console.log(`新工作簿中没有找到国内机票工作表`);
            }
          } else {
            console.log(`第九行中未找到"国内机票"数据`);
          }
        } else {
          console.log(`结算单工作表中没有第九行数据`);
        }
      } else {
        console.log(`未找到结算单工作表`);
      }

      console.log(`===== 第九行国内机票总计金额替换处理结束（原有逻辑） =====`);

      // 处理第十一行（国内酒店）
      processExpenseTypeRow(11, "国内酒店");

      // 处理第十三行（国内火车）
      processExpenseTypeRow(13, "火车票");

      // 处理第十行（国际机票）
      processExpenseTypeRow(10, "国际机票");

      // 处理第十二行（国际酒店）
      processExpenseTypeRow(12, "国际酒店");

      console.log(`===== 所有费用类型服务费总计金额替换处理结束 =====`);

      // 恢复原有的总计金额计算逻辑
      // 处理第十行的国际机票总计金额替换
      console.log(`===== 开始处理第十行国际机票总计金额替换 =====`);

      // 检查结算单工作表的第十行是否包含"国际机票"
      if (summaryWorksheet) {
        const tenthRow = summaryWorksheet.getRow(10);
        if (tenthRow) {
          console.log(`检查结算单第十行数据...`);

          // 遍历第十行的每个单元格，查找"国际机票"
          let internationalFlightColIndex = -1;
          for (let col = 1; col <= 50; col++) {
            // 检查前50列
            const cell = tenthRow.getCell(col);
            if (cell.value === "国际机票") {
              internationalFlightColIndex = col;
              console.log(`在第${col}列找到"国际机票"`);
              break;
            }
          }

          if (internationalFlightColIndex !== -1) {
            console.log(`发现国际机票数据，开始从拆分好的工作表中查找合计金额`);

            // 从新创建的工作簿中查找国际机票工作表
            const internationalFlightWorksheet =
              newWorkbook.getWorksheet("国际机票");
            if (internationalFlightWorksheet) {
              console.log(`找到拆分好的国际机票工作表`);

              // 获取工作表的所有数据
              const internationalFlightData: any[][] = [];
              internationalFlightWorksheet.eachRow((row, rowNumber) => {
                const rowData: any[] = [];
                row.eachCell((cell, colNumber) => {
                  rowData[colNumber] = cell.value;
                });
                internationalFlightData[rowNumber] = rowData;
              });

              console.log(
                `国际机票工作表共${internationalFlightData.length}行数据`
              );

              // 查找合计行
              let totalAmount = null;
              for (let i = internationalFlightData.length - 1; i >= 0; i--) {
                const row = internationalFlightData[i];
                // 检查这一行是否包含"合计"或数字
                const hasTotalText = row.some(
                  cell => typeof cell === "string" && cell.includes("合计")
                );
                const hasNumbers = row.some(
                  cell =>
                    typeof cell === "number" ||
                    (typeof cell === "string" && /^\d+\.?\d*$/.test(cell))
                );

                if (hasTotalText || hasNumbers) {
                  console.log(`找到合计行（第${i + 1}行）:`, row);

                  // 查找合计行中的数字（最后一个数字是合计金额）
                  for (let j = row.length - 1; j >= 0; j--) {
                    const cell = row[j];
                    if (typeof cell === "number") {
                      totalAmount = cell;
                      break;
                    } else if (
                      typeof cell === "string" &&
                      /^\d+\.?\d*$/.test(cell)
                    ) {
                      totalAmount = parseFloat(cell);
                      break;
                    }
                  }
                  break;
                }
              }

              if (totalAmount !== null) {
                console.log(`找到合计金额: ${totalAmount}`);

                // 替换第十行中国际机票对应的总计金额
                console.log(
                  `国际机票在第${internationalFlightColIndex}列，开始查找总计金额列`
                );

                // 打印第十行的完整数据进行调试
                console.log(`第十行完整数据:`);
                for (let debugCol = 1; debugCol <= 20; debugCol++) {
                  const debugCell = tenthRow.getCell(debugCol);
                  if (
                    debugCell.value !== null &&
                    debugCell.value !== undefined
                  ) {
                    console.log(
                      `  第${debugCol}列: ${debugCell.value} (类型: ${typeof debugCell.value})`
                    );
                  }
                }

                // 查找国际机票后面的所有数字列，找到最后一个数字列作为总计金额
                let allNumberCols: { col: number; value: any }[] = [];
                for (
                  let searchCol = internationalFlightColIndex + 1;
                  searchCol <= 20;
                  searchCol++
                ) {
                  const searchCell = tenthRow.getCell(searchCol);
                  if (
                    searchCell.value !== null &&
                    searchCell.value !== undefined
                  ) {
                    // 检查是否是数字
                    if (typeof searchCell.value === "number") {
                      allNumberCols.push({
                        col: searchCol,
                        value: searchCell.value
                      });
                      console.log(
                        `找到数字列: 第${searchCol}列 = ${searchCell.value}`
                      );
                    } else if (
                      typeof searchCell.value === "string" &&
                      /^\d+\.?\d*$/.test(searchCell.value)
                    ) {
                      allNumberCols.push({
                        col: searchCol,
                        value: parseFloat(searchCell.value)
                      });
                      console.log(
                        `找到数字字符串列: 第${searchCol}列 = ${searchCell.value}`
                      );
                    }
                  }
                }

                // 使用最后一个数字列作为总计金额列
                let targetColIndex = -1;
                if (allNumberCols.length > 0) {
                  const lastNumberCol = allNumberCols[allNumberCols.length - 1];
                  targetColIndex = lastNumberCol.col;
                  console.log(
                    `选择最后一个数字列作为总计金额: 第${targetColIndex}列 = ${lastNumberCol.value}`
                  );
                }

                if (targetColIndex !== -1) {
                  const targetCell = tenthRow.getCell(targetColIndex);
                  const oldValue = targetCell.value;
                  targetCell.value = totalAmount;
                  console.log(
                    `替换第十行第${targetColIndex}列: ${oldValue} -> ${totalAmount}`
                  );
                } else {
                  console.log(`未找到第十行中的总计金额列，请检查数据结构`);
                }
              } else {
                console.log(`未在拆分好的国际机票工作表中找到有效的合计金额`);
              }
            } else {
              console.log(`新工作簿中没有找到国际机票工作表`);
            }
          } else {
            console.log(`第十行中未找到"国际机票"数据`);
          }
        } else {
          console.log(`结算单工作表中没有第十行数据`);
        }
      } else {
        console.log(`未找到结算单工作表`);
      }

      console.log(`===== 第十行国际机票总计金额替换处理结束 =====`);

      // 处理第十一行的国内酒店总计金额替换
      console.log(`===== 开始处理第十一行国内酒店总计金额替换 =====`);

      // 检查结算单工作表的第十一行是否包含"国内酒店"
      if (summaryWorksheet) {
        const eleventhRow = summaryWorksheet.getRow(11);
        if (eleventhRow) {
          console.log(`检查结算单第十一行数据...`);

          // 遍历第十一行的每个单元格，查找"国内酒店"
          let domesticHotelColIndex = -1;
          for (let col = 1; col <= 50; col++) {
            // 检查前50列
            const cell = eleventhRow.getCell(col);
            if (cell.value === "国内酒店") {
              domesticHotelColIndex = col;
              console.log(`在第${col}列找到"国内酒店"`);
              break;
            }
          }

          if (domesticHotelColIndex !== -1) {
            console.log(`发现国内酒店数据，开始从拆分好的工作表中查找合计金额`);

            // 从新创建的工作簿中查找国内酒店工作表
            const domesticHotelWorksheet = newWorkbook.getWorksheet("国内酒店");
            if (domesticHotelWorksheet) {
              console.log(`找到拆分好的国内酒店工作表`);

              // 获取工作表的所有数据
              const domesticHotelData: any[][] = [];
              domesticHotelWorksheet.eachRow((row, rowNumber) => {
                const rowData: any[] = [];
                row.eachCell((cell, colNumber) => {
                  rowData[colNumber] = cell.value;
                });
                domesticHotelData[rowNumber] = rowData;
              });

              console.log(`国内酒店工作表共${domesticHotelData.length}行数据`);

              // 查找合计行
              let totalAmount = null;
              for (let i = domesticHotelData.length - 1; i >= 0; i--) {
                const row = domesticHotelData[i];
                // 检查这一行是否包含"合计"或数字
                const hasTotalText = row.some(
                  cell => typeof cell === "string" && cell.includes("合计")
                );
                const hasNumbers = row.some(
                  cell =>
                    typeof cell === "number" ||
                    (typeof cell === "string" && /^\d+\.?\d*$/.test(cell))
                );

                if (hasTotalText || hasNumbers) {
                  console.log(`找到合计行（第${i + 1}行）:`, row);

                  // 查找合计行中的数字（最后一个数字是合计金额）
                  for (let j = row.length - 1; j >= 0; j--) {
                    const cell = row[j];
                    if (typeof cell === "number") {
                      totalAmount = cell;
                      break;
                    } else if (
                      typeof cell === "string" &&
                      /^\d+\.?\d*$/.test(cell)
                    ) {
                      totalAmount = parseFloat(cell);
                      break;
                    }
                  }
                  break;
                }
              }

              if (totalAmount !== null) {
                console.log(`找到合计金额: ${totalAmount}`);

                // 替换第十一行中国内酒店对应的总计金额
                console.log(
                  `国内酒店在第${domesticHotelColIndex}列，开始查找总计金额列`
                );

                // 打印第十一行的完整数据进行调试
                console.log(`第十一行完整数据:`);
                for (let debugCol = 1; debugCol <= 20; debugCol++) {
                  const debugCell = eleventhRow.getCell(debugCol);
                  if (
                    debugCell.value !== null &&
                    debugCell.value !== undefined
                  ) {
                    console.log(
                      `  第${debugCol}列: ${debugCell.value} (类型: ${typeof debugCell.value})`
                    );
                  }
                }

                // 查找国内酒店后面的所有数字列，找到最后一个数字列作为总计金额
                let allNumberCols: { col: number; value: any }[] = [];
                for (
                  let searchCol = domesticHotelColIndex + 1;
                  searchCol <= 20;
                  searchCol++
                ) {
                  const searchCell = eleventhRow.getCell(searchCol);
                  if (
                    searchCell.value !== null &&
                    searchCell.value !== undefined
                  ) {
                    // 检查是否是数字
                    if (typeof searchCell.value === "number") {
                      allNumberCols.push({
                        col: searchCol,
                        value: searchCell.value
                      });
                      console.log(
                        `找到数字列: 第${searchCol}列 = ${searchCell.value}`
                      );
                    } else if (
                      typeof searchCell.value === "string" &&
                      /^\d+\.?\d*$/.test(searchCell.value)
                    ) {
                      allNumberCols.push({
                        col: searchCol,
                        value: parseFloat(searchCell.value)
                      });
                      console.log(
                        `找到数字字符串列: 第${searchCol}列 = ${searchCell.value}`
                      );
                    }
                  }
                }

                // 使用最后一个数字列作为总计金额列
                let targetColIndex = -1;
                if (allNumberCols.length > 0) {
                  const lastNumberCol = allNumberCols[allNumberCols.length - 1];
                  targetColIndex = lastNumberCol.col;
                  console.log(
                    `选择最后一个数字列作为总计金额: 第${targetColIndex}列 = ${lastNumberCol.value}`
                  );
                }

                if (targetColIndex !== -1) {
                  const targetCell = eleventhRow.getCell(targetColIndex);
                  const oldValue = targetCell.value;
                  targetCell.value = totalAmount;
                  console.log(
                    `替换第十一行第${targetColIndex}列: ${oldValue} -> ${totalAmount}`
                  );
                } else {
                  console.log(`未找到第十一行中的总计金额列，请检查数据结构`);
                }
              } else {
                console.log(`未在拆分好的国内酒店工作表中找到有效的合计金额`);
              }
            } else {
              console.log(`新工作簿中没有找到国内酒店工作表`);
            }
          } else {
            console.log(`第十一行中未找到"国内酒店"数据`);
          }
        } else {
          console.log(`结算单工作表中没有第十一行数据`);
        }
      } else {
        console.log(`未找到结算单工作表`);
      }

      console.log(`===== 第十一行国内酒店总计金额替换处理结束 =====`);

      // 处理第十二行的国际酒店总计金额替换
      console.log(`===== 开始处理第十二行国际酒店总计金额替换 =====`);

      // 检查结算单工作表的第十二行是否包含"国际酒店"
      if (summaryWorksheet) {
        const twelfthRow = summaryWorksheet.getRow(12);
        if (twelfthRow) {
          console.log(`检查结算单第十二行数据...`);

          // 遍历第十二行的每个单元格，查找"国际酒店"
          let internationalHotelColIndex = -1;
          for (let col = 1; col <= 50; col++) {
            // 检查前50列
            const cell = twelfthRow.getCell(col);
            if (cell.value === "国际酒店") {
              internationalHotelColIndex = col;
              console.log(`在第${col}列找到"国际酒店"`);
              break;
            }
          }

          if (internationalHotelColIndex !== -1) {
            console.log(`发现国际酒店数据，开始从拆分好的工作表中查找合计金额`);

            // 从新创建的工作簿中查找国际酒店工作表
            const internationalHotelWorksheet =
              newWorkbook.getWorksheet("国际酒店");
            if (internationalHotelWorksheet) {
              console.log(`找到拆分好的国际酒店工作表`);

              // 获取工作表的所有数据
              const internationalHotelData: any[][] = [];
              internationalHotelWorksheet.eachRow((row, rowNumber) => {
                const rowData: any[] = [];
                row.eachCell((cell, colNumber) => {
                  rowData[colNumber] = cell.value;
                });
                internationalHotelData[rowNumber] = rowData;
              });

              console.log(
                `国际酒店工作表共${internationalHotelData.length}行数据`
              );

              // 查找合计行
              let totalAmount = null;
              for (let i = internationalHotelData.length - 1; i >= 0; i--) {
                const row = internationalHotelData[i];
                // 检查这一行是否包含"合计"或数字
                const hasTotalText = row.some(
                  cell => typeof cell === "string" && cell.includes("合计")
                );
                const hasNumbers = row.some(
                  cell =>
                    typeof cell === "number" ||
                    (typeof cell === "string" && /^\d+\.?\d*$/.test(cell))
                );

                if (hasTotalText || hasNumbers) {
                  console.log(`找到合计行（第${i + 1}行）:`, row);

                  // 查找合计行中的数字（最后一个数字是合计金额）
                  for (let j = row.length - 1; j >= 0; j--) {
                    const cell = row[j];
                    if (typeof cell === "number") {
                      totalAmount = cell;
                      break;
                    } else if (
                      typeof cell === "string" &&
                      /^\d+\.?\d*$/.test(cell)
                    ) {
                      totalAmount = parseFloat(cell);
                      break;
                    }
                  }
                  break;
                }
              }

              if (totalAmount !== null) {
                console.log(`找到合计金额: ${totalAmount}`);

                // 替换第十二行中国际酒店对应的总计金额
                console.log(
                  `国际酒店在第${internationalHotelColIndex}列，开始查找总计金额列`
                );

                // 打印第十二行的完整数据进行调试
                console.log(`第十二行完整数据:`);
                for (let debugCol = 1; debugCol <= 20; debugCol++) {
                  const debugCell = twelfthRow.getCell(debugCol);
                  if (
                    debugCell.value !== null &&
                    debugCell.value !== undefined
                  ) {
                    console.log(
                      `  第${debugCol}列: ${debugCell.value} (类型: ${typeof debugCell.value})`
                    );
                  }
                }

                // 查找国际酒店后面的所有数字列，找到最后一个数字列作为总计金额
                let allNumberCols: { col: number; value: any }[] = [];
                for (
                  let searchCol = internationalHotelColIndex + 1;
                  searchCol <= 20;
                  searchCol++
                ) {
                  const searchCell = twelfthRow.getCell(searchCol);
                  if (
                    searchCell.value !== null &&
                    searchCell.value !== undefined
                  ) {
                    // 检查是否是数字
                    if (typeof searchCell.value === "number") {
                      allNumberCols.push({
                        col: searchCol,
                        value: searchCell.value
                      });
                      console.log(
                        `找到数字列: 第${searchCol}列 = ${searchCell.value}`
                      );
                    } else if (
                      typeof searchCell.value === "string" &&
                      /^\d+\.?\d*$/.test(searchCell.value)
                    ) {
                      allNumberCols.push({
                        col: searchCol,
                        value: parseFloat(searchCell.value)
                      });
                      console.log(
                        `找到数字字符串列: 第${searchCol}列 = ${searchCell.value}`
                      );
                    }
                  }
                }

                // 使用最后一个数字列作为总计金额列
                let targetColIndex = -1;
                if (allNumberCols.length > 0) {
                  const lastNumberCol = allNumberCols[allNumberCols.length - 1];
                  targetColIndex = lastNumberCol.col;
                  console.log(
                    `选择最后一个数字列作为总计金额: 第${targetColIndex}列 = ${lastNumberCol.value}`
                  );
                }

                if (targetColIndex !== -1) {
                  const targetCell = twelfthRow.getCell(targetColIndex);
                  const oldValue = targetCell.value;
                  targetCell.value = totalAmount;
                  console.log(
                    `替换第十二行第${targetColIndex}列: ${oldValue} -> ${totalAmount}`
                  );
                } else {
                  console.log(`未找到第十二行中的总计金额列，请检查数据结构`);
                }
              } else {
                console.log(`未在拆分好的国际酒店工作表中找到有效的合计金额`);
              }
            } else {
              console.log(`新工作簿中没有找到国际酒店工作表`);
            }
          } else {
            console.log(`第十二行中未找到"国际酒店"数据`);
          }
        } else {
          console.log(`结算单工作表中没有第十二行数据`);
        }
      } else {
        console.log(`未找到结算单工作表`);
      }

      console.log(`===== 第十二行国际酒店总计金额替换处理结束 =====`);

      // 处理第十三行的国内火车总计金额替换
      console.log(`===== 开始处理第十三行国内火车总计金额替换 =====`);

      // 检查结算单工作表的第十三行是否包含"国内火车"
      if (summaryWorksheet) {
        const thirteenthRow = summaryWorksheet.getRow(13);
        if (thirteenthRow) {
          console.log(`检查结算单第十三行数据...`);

          // 遍历第十三行的每个单元格，查找"国内火车"
          let domesticTrainColIndex = -1;
          for (let col = 1; col <= 50; col++) {
            // 检查前50列
            const cell = thirteenthRow.getCell(col);
            if (cell.value === "国内火车") {
              domesticTrainColIndex = col;
              console.log(`在第${col}列找到"国内火车"`);
              break;
            }
          }

          if (domesticTrainColIndex !== -1) {
            console.log(`发现国内火车数据，开始从拆分好的工作表中查找合计金额`);

            // 从新创建的工作簿中查找火车票工作表
            const trainWorksheet = newWorkbook.getWorksheet("火车票");
            if (trainWorksheet) {
              console.log(`找到拆分好的火车票工作表`);

              // 获取工作表的所有数据
              const trainData: any[][] = [];
              trainWorksheet.eachRow((row, rowNumber) => {
                const rowData: any[] = [];
                row.eachCell((cell, colNumber) => {
                  rowData[colNumber] = cell.value;
                });
                trainData[rowNumber] = rowData;
              });

              console.log(`火车票工作表共${trainData.length}行数据`);

              // 查找合计行
              let totalAmount = null;
              for (let i = trainData.length - 1; i >= 0; i--) {
                const row = trainData[i];
                // 检查这一行是否包含"合计"或数字
                const hasTotalText = row.some(
                  cell => typeof cell === "string" && cell.includes("合计")
                );
                const hasNumbers = row.some(
                  cell =>
                    typeof cell === "number" ||
                    (typeof cell === "string" && /^\d+\.?\d*$/.test(cell))
                );

                if (hasTotalText || hasNumbers) {
                  console.log(`找到合计行（第${i + 1}行）:`, row);

                  // 查找合计行中的数字（最后一个数字是合计金额）
                  for (let j = row.length - 1; j >= 0; j--) {
                    const cell = row[j];
                    if (typeof cell === "number") {
                      totalAmount = cell;
                      break;
                    } else if (
                      typeof cell === "string" &&
                      /^\d+\.?\d*$/.test(cell)
                    ) {
                      totalAmount = parseFloat(cell);
                      break;
                    }
                  }
                  break;
                }
              }

              if (totalAmount !== null) {
                console.log(`找到合计金额: ${totalAmount}`);

                // 替换第十三行中国内火车对应的总计金额
                console.log(
                  `国内火车在第${domesticTrainColIndex}列，开始查找总计金额列`
                );

                // 打印第十三行的完整数据进行调试
                console.log(`第十三行完整数据:`);
                for (let debugCol = 1; debugCol <= 20; debugCol++) {
                  const debugCell = thirteenthRow.getCell(debugCol);
                  if (
                    debugCell.value !== null &&
                    debugCell.value !== undefined
                  ) {
                    console.log(
                      `  第${debugCol}列: ${debugCell.value} (类型: ${typeof debugCell.value})`
                    );
                  }
                }

                // 查找国内火车后面的所有数字列，找到最后一个数字列作为总计金额
                let allNumberCols: { col: number; value: any }[] = [];
                for (
                  let searchCol = domesticTrainColIndex + 1;
                  searchCol <= 20;
                  searchCol++
                ) {
                  const searchCell = thirteenthRow.getCell(searchCol);
                  if (
                    searchCell.value !== null &&
                    searchCell.value !== undefined
                  ) {
                    // 检查是否是数字
                    if (typeof searchCell.value === "number") {
                      allNumberCols.push({
                        col: searchCol,
                        value: searchCell.value
                      });
                      console.log(
                        `找到数字列: 第${searchCol}列 = ${searchCell.value}`
                      );
                    } else if (
                      typeof searchCell.value === "string" &&
                      /^\d+\.?\d*$/.test(searchCell.value)
                    ) {
                      allNumberCols.push({
                        col: searchCol,
                        value: parseFloat(searchCell.value)
                      });
                      console.log(
                        `找到数字字符串列: 第${searchCol}列 = ${searchCell.value}`
                      );
                    }
                  }
                }

                // 使用最后一个数字列作为总计金额列
                let targetColIndex = -1;
                if (allNumberCols.length > 0) {
                  const lastNumberCol = allNumberCols[allNumberCols.length - 1];
                  targetColIndex = lastNumberCol.col;
                  console.log(
                    `选择最后一个数字列作为总计金额: 第${targetColIndex}列 = ${lastNumberCol.value}`
                  );
                }

                if (targetColIndex !== -1) {
                  const targetCell = thirteenthRow.getCell(targetColIndex);
                  const oldValue = targetCell.value;
                  targetCell.value = totalAmount;
                  console.log(
                    `替换第十三行第${targetColIndex}列: ${oldValue} -> ${totalAmount}`
                  );
                } else {
                  console.log(`未找到第十三行中的总计金额列，请检查数据结构`);
                }
              } else {
                console.log(`未在拆分好的火车票工作表中找到有效的合计金额`);
              }
            } else {
              console.log(`新工作簿中没有找到火车票工作表`);
            }
          } else {
            console.log(`第十三行中未找到"国内火车"数据`);
          }
        } else {
          console.log(`结算单工作表中没有第十三行数据`);
        }
      } else {
        console.log(`未找到结算单工作表`);
      }

      console.log(`===== 第十三行国内火车总计金额替换处理结束 =====`);

      // 处理第九行到十三行的折叠逻辑（如果总计金额为0）
      console.log(`===== 开始处理第九行到十七行折叠逻辑 =====`);

      if (summaryWorksheet) {
        const rowsToProcess = [
          { rowNumber: 9, name: "国内机票" },
          { rowNumber: 10, name: "国际机票" },
          { rowNumber: 11, name: "国内酒店" },
          { rowNumber: 12, name: "国际酒店" },
          { rowNumber: 13, name: "国内火车" },
          { rowNumber: 14, name: "国内用车" },
          { rowNumber: 15, name: "国内外卖" },
          { rowNumber: 16, name: "商务卡" },
          { rowNumber: 17, name: "滞纳金" }
        ];

        for (const rowInfo of rowsToProcess) {
          const row = summaryWorksheet.getRow(rowInfo.rowNumber);
          if (row) {
            console.log(
              `检查第${rowInfo.rowNumber}行(${rowInfo.name})的总计金额...`
            );

            // 查找该行的总计金额（最后一个数字列）
            let totalAmount = null;
            let totalColIndex = -1;
            let allNumberCols: { col: number; value: any }[] = [];

            // 遍历该行的所有单元格，查找数字列
            for (let col = 1; col <= 50; col++) {
              const cell = row.getCell(col);
              if (cell.value !== null && cell.value !== undefined) {
                if (typeof cell.value === "number") {
                  allNumberCols.push({ col: col, value: cell.value });
                } else if (
                  typeof cell.value === "string" &&
                  /^\d+\.?\d*$/.test(cell.value)
                ) {
                  allNumberCols.push({
                    col: col,
                    value: parseFloat(cell.value)
                  });
                }
              }
            }

            // 使用最后一个数字作为总计金额
            if (allNumberCols.length > 0) {
              const lastNumberCol = allNumberCols[allNumberCols.length - 1];
              totalAmount = lastNumberCol.value;
              totalColIndex = lastNumberCol.col;
              console.log(
                `第${rowInfo.rowNumber}行总计金额: ${totalAmount} (第${totalColIndex}列)`
              );
            }

            // 如果总计金额为0，进行折叠处理
            if (totalAmount !== null && totalAmount === 0) {
              console.log(
                `第${rowInfo.rowNumber}行(${rowInfo.name})总计金额为0，开始隐藏处理`
              );

              // 隐藏整行 - 尝试多种方法

              // 方法1: 标准隐藏
              row.hidden = true;

              // 方法2: 设置极小行高
              row.height = 0; // 1磅，Excel可能的最小行高

              // 方法3: 使用outline层级隐藏
              row.outlineLevel = 1;

              console.log(
                `第${rowInfo.rowNumber}行隐藏处理完成 (hidden: ${row.hidden}, height: ${row.height})`
              );
            } else {
              console.log(
                `第${rowInfo.rowNumber}行总计金额不为0(${totalAmount})，无需隐藏`
              );
            }
          } else {
            console.log(`第${rowInfo.rowNumber}行不存在`);
          }
        }

        // 可选：如果有行被隐藏，可以调整行高或添加备注
        console.log(`第九行到十三行折叠逻辑处理完成`);
      } else {
        console.log(`未找到结算单工作表，无法进行折叠处理`);
      }

      console.log(`===== 第九行到十三行折叠逻辑处理结束 =====`);

      // 最终确认隐藏设置：在生成文件前再次确保隐藏属性生效
      console.log(`===== 最终确认隐藏设置 =====`);
      if (summaryWorksheet) {
        const rowsToHide = [9, 10, 11, 12, 13];

        // 设置工作表的行隐藏属性
        for (const rowNum of rowsToHide) {
          const row = summaryWorksheet.getRow(rowNum);
          if (row && row.hidden) {
            // 再次设置隐藏属性，确保在生成文件时生效
            row.hidden = true;
            row.height = 0;

            // 尝试通过设置行属性来强制隐藏
            if (row.model) {
              row.model.hidden = true;
            }

            console.log(
              `再次确认第${rowNum}行隐藏设置 (hidden: ${row.hidden})`
            );
          }
        }

        // 尝试直接操作工作表的行隐藏设置
        // 这是ExcelJS的内部方法，可能更有效
        try {
          // 遍历所有需要隐藏的行
          rowsToHide.forEach(rowNum => {
            const row = summaryWorksheet.getRow(rowNum);
            if (row && row.hidden) {
              // 强制设置行模型属性
              if (row.model) {
                row.model.hidden = true;
                row.model.height = 0;
              }
            }
          });
        } catch (error) {
          console.log(`内部隐藏设置方法失败: ${error}`);
        }
      }
      console.log(`===== 最终确认隐藏设置完成 =====`);

      // 生成Excel文件内容
      const excelBuffer = await newWorkbook.xlsx.writeBuffer();

      // 使用用户编辑的文件名
      const finalFileName = latestFileName.endsWith(".xlsx")
        ? latestFileName
        : `${latestFileName}.xlsx`;

      console.log(`使用文件名: ${finalFileName}`);
      zip.file(finalFileName, excelBuffer);
    }

    // 生成ZIP文件
    const zipBuffer = await zip.generateAsync({ type: "array" });

    // 下载ZIP文件
    const zipBlob = new Blob([new Uint8Array(zipBuffer)], {
      type: "application/zip"
    });
    const fileName = `账单分账_${uploadedFile.value?.name.replace(".xlsx", "").replace(".xls", "")}_${new Date().toISOString().slice(0, 10)}.zip`;
    saveAs(zipBlob, fileName);

    ElMessage.success(
      `成功为 ${groupInfo.length} 个公司生成Excel文件并打包为ZIP文件`
    );
    console.log(`生成完成: ${fileName}`);
  } catch (error) {
    console.error("生成文件失败:", error);
    ElMessage.error("生成文件失败");
  } finally {
    generating.value = false;
  }
};
</script>

<template>
  <div class="bill-split-container">
    <div class="bill-split-header">
      <h1>账单分账</h1>
      <p>上传Excel文件进行账单分账处理</p>
    </div>
    <div class="bill-split-content">
      <!-- 文件上传区域 -->
      <div class="upload-section">
        <el-upload
          class="upload-demo"
          drag
          :auto-upload="false"
          :show-file-list="false"
          :before-upload="beforeUpload"
          @change="handleFileChange"
          accept=".xlsx,.xls"
        >
          <el-icon class="el-icon--upload">
            <UploadFilled />
          </el-icon>
          <div class="el-upload__text">
            将Excel文件拖拽到此处，或<em>点击上传</em>
          </div>
          <template #tip>
            <div class="el-upload__tip">
              只能上传.xlsx或.xls文件，且不超过10MB
            </div>
          </template>
        </el-upload>

        <div v-if="uploadedFile" class="file-info">
          <p>已选择文件: {{ uploadedFile.name }}</p>
          <p>文件大小: {{ (uploadedFile.size / 1024 / 1024).toFixed(2) }} MB</p>
        </div>
      </div>

      <!-- 数据展示区域 -->
      <div v-if="showData && excelData.length > 0" class="data-section">
        <div class="data-header">
          <h3>分组信息 - 将生成以下文件</h3>
          <div class="header-buttons">
            <el-button
              type="success"
              :loading="generating"
              @click="generateGroupedExcelFiles"
              :disabled="!showData"
            >
              {{ generating ? "生成中..." : "生成分组Excel文件" }}
            </el-button>
          </div>
        </div>

        <div class="data-summary">
          <el-alert
            title="分组概览"
            type="info"
            :description="`检测到 ${getGroupCount()} 个分组，将生成 ${getGroupCount()} 个Excel文件`"
            show-icon
          />
        </div>

        <div class="data-table">
          <el-table :data="getGroupInfo()" border style="width: 100%">
            <el-table-column prop="groupName" label="公司名称" width="200" />
            <el-table-column label="酒店明细(国内)" width="150">
              <template #default="scope">
                <div v-if="scope.row.hotelInfo">
                  <div>{{ scope.row.hotelInfo.count }} 条</div>
                  <div class="text-gray-500 text-sm">
                    {{ scope.row.hotelInfo.rowRange }}
                  </div>
                </div>
                <div v-else class="text-gray-400">无数据</div>
              </template>
            </el-table-column>
            <el-table-column label="酒店明细(国际)" width="150">
              <template #default="scope">
                <div v-if="scope.row.internationalHotelInfo">
                  <div>{{ scope.row.internationalHotelInfo.count }} 条</div>
                  <div class="text-gray-500 text-sm">
                    {{ scope.row.internationalHotelInfo.rowRange }}
                  </div>
                </div>
                <div v-else class="text-gray-400">无数据</div>
              </template>
            </el-table-column>
            <el-table-column label="火车票明细" width="150">
              <template #default="scope">
                <div v-if="scope.row.trainInfo">
                  <div>{{ scope.row.trainInfo.count }} 条</div>
                  <div class="text-gray-500 text-sm">
                    {{ scope.row.trainInfo.rowRange }}
                  </div>
                </div>
                <div v-else class="text-gray-400">无数据</div>
              </template>
            </el-table-column>
            <el-table-column label="机票明细(国内)" width="150">
              <template #default="scope">
                <div v-if="scope.row.flightInfo">
                  <div>{{ scope.row.flightInfo.count }} 条</div>
                  <div class="text-gray-500 text-sm">
                    {{ scope.row.flightInfo.rowRange }}
                  </div>
                </div>
                <div v-else class="text-gray-400">无数据</div>
              </template>
            </el-table-column>
            <el-table-column label="机票明细(国际)" width="150">
              <template #default="scope">
                <div v-if="scope.row.internationalFlightInfo">
                  <div>{{ scope.row.internationalFlightInfo.count }} 条</div>
                  <div class="text-gray-500 text-sm">
                    {{ scope.row.internationalFlightInfo.rowRange }}
                  </div>
                </div>
                <div v-else class="text-gray-400">无数据</div>
              </template>
            </el-table-column>
            <el-table-column prop="totalCount" label="总数据条数" width="120" />
            <el-table-column label="生成文件名">
              <template #default="scope">
                <el-input
                  :model-value="scope.row.editableFileName"
                  @update:model-value="
                    value => updateFileName(scope.row.groupName, value)
                  "
                  placeholder="请输入文件名"
                  style="width: 100%"
                >
                  <template #suffix>.xlsx</template>
                </el-input>
              </template>
            </el-table-column>
          </el-table>
        </div>
      </div>

      <!-- 空状态 -->
      <div v-if="!showData" class="placeholder-content">
        <el-empty description="请上传Excel文件开始处理" />
      </div>
    </div>
  </div>
</template>

<style scoped>
.bill-split-container {
  background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
  position: relative;
  overflow: hidden;
  padding: 20px;
}

.bill-split-header {
  text-align: center;
  margin-bottom: 30px;
  padding: 20px;
  background: rgba(255, 255, 255, 0.9);
  border-radius: 8px;
  box-shadow: 0 2px 12px rgba(0, 0, 0, 0.1);
}

.bill-split-header h1 {
  color: #303133;
  margin: 0 0 10px 0;
  font-size: 28px;
}

.bill-split-header p {
  color: #606266;
  margin: 0;
  font-size: 16px;
}

.bill-split-content {
  background: rgba(255, 255, 255, 0.9);
  border-radius: 8px;
  box-shadow: 0 2px 12px rgba(0, 0, 0, 0.1);
  padding: 40px;
  min-height: 400px;
}

.upload-section {
  margin-bottom: 30px;
}

.file-info {
  margin-top: 20px;
  padding: 15px;
  background: #f5f7fa;
  border-radius: 6px;
  border-left: 4px solid #409eff;
}

.file-info p {
  margin: 5px 0;
  color: #606266;
}

.data-section {
  margin-top: 30px;
}

.data-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 20px;
}

.header-buttons {
  display: flex;
  gap: 10px;
}

.data-header h3 {
  margin: 0;
  color: #303133;
}

.data-summary {
  margin-bottom: 20px;
}

.data-table {
  background: #fff;
  border-radius: 6px;
  overflow: hidden;
}

.data-more {
  text-align: center;
  color: #909399;
  margin-top: 10px;
  font-size: 14px;
}

.placeholder-content {
  display: flex;
  align-items: center;
  justify-content: center;
  height: 200px;
}
</style>
