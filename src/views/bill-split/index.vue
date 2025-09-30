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
              departmentKeyword: "入住人部门"
            },
            {
              name: "火车票明细",
              key: "train",
              departmentKeyword: "乘车人部门"
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

// 获取分组信息
const getGroupInfo = () => {
  const allGroupInfo: any[] = [];

  // 处理酒店数据
  if (allSheetData.value.hotel && allSheetData.value.hotel.length > 0) {
    const hotelGroups = processSheetData(
      allSheetData.value.hotel,
      "酒店明细(国内)",
      "入住人部门",
      "hotel"
    );
    allGroupInfo.push(...hotelGroups);
  }

  // 处理火车票数据
  if (allSheetData.value.train && allSheetData.value.train.length > 0) {
    const trainGroups = processSheetData(
      allSheetData.value.train,
      "火车票明细",
      "乘车人部门",
      "train"
    );
    allGroupInfo.push(...trainGroups);
  }

  return allGroupInfo;
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
      return (
        row[departmentColumnIndex] &&
        row[departmentColumnIndex].toString().trim() !== ""
      );
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
    const firstPart = department.split("-")[0];
    if (!groups.has(firstPart)) {
      groups.set(firstPart, []);
    }
    groups.get(firstPart)!.push(item);
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

// 更新文件名
const updateFileName = (
  groupName: string,
  newFileName: string,
  sheetType?: string
) => {
  const existing = editableFileNames.value.find(
    item => item.groupName === groupName && item.sheetType === sheetType
  );
  if (existing) {
    existing.fileName = newFileName;
  } else {
    editableFileNames.value.push({
      groupName: groupName,
      fileName: newFileName,
      sheetType: sheetType
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

// 生成分组Excel文件并打包成ZIP
const generateGroupedExcelFiles = async () => {
  if (!originalWorkbook.value || Object.keys(allSheetData.value).length === 0) {
    ElMessage.error("请先上传并处理Excel文件");
    return;
  }

  generating.value = true;

  try {
    const groupInfo = getGroupInfo();
    console.log(`准备生成 ${groupInfo.length} 个Excel文件`);

    // 创建ZIP文件
    const zip = new JSZip();

    // 按分组名组织数据，每个分组可能包含多个工作表
    const groupedData = new Map<
      string,
      { sheetType: string; data: any[]; groupName: string }[]
    >();

    groupInfo.forEach(group => {
      if (!groupedData.has(group.groupName)) {
        groupedData.set(group.groupName, []);
      }
      groupedData.get(group.groupName)!.push({
        sheetType: group.sheetType,
        data: allSheetData.value[group.sheetType],
        groupName: group.groupName
      });
    });

    // 为每个分组生成Excel文件
    for (const [groupName, sheets] of groupedData.entries()) {
      console.log(
        `生成文件: ${groupName}.xlsx，包含 ${sheets.length} 个工作表`
      );

      // 创建新的工作簿
      const newWorkbook = new ExcelJS.Workbook();

      // 为每个工作表创建对应的工作表
      for (const sheetInfo of sheets) {
        const { sheetType, data } = sheetInfo;

        if (!data || data.length === 0) continue;

        const sheetName = sheetType === "hotel" ? "国内酒店" : "火车票";
        const newWorksheet = newWorkbook.addWorksheet(sheetName, {
          views: [{ showGridLines: true }]
        });

        // 设置默认行高
        newWorksheet.properties.defaultRowHeight = 40;

        // 获取该工作表的分组数据
        const departmentKeyword =
          sheetType === "hotel" ? "入住人部门" : "乘车人部门";
        const departmentColumnIndex = (data[2] as any[]).findIndex(
          (cell: any) => cell && cell.toString().includes(departmentKeyword)
        );

        if (departmentColumnIndex === -1) continue;

        // 筛选该分组的数据
        const groupData = data.slice(3).filter((row: any[], index: number) => {
          return (
            row[departmentColumnIndex] &&
            row[departmentColumnIndex].toString().split("-")[0] === groupName
          );
        });

        // 复制原始前三行
        const headerRows = data.slice(0, 3);

        // 写入数据到工作表
        const newData = [...headerRows];
        groupData.forEach(row => {
          newData.push(row);
        });

        console.log(`  工作表 ${sheetName}: ${newData.length} 行数据`);

        // 应用样式和格式
        await applyWorksheetStyling(newWorksheet, newData, departmentKeyword);
      }

      // 生成Excel文件内容
      const excelBuffer = await newWorkbook.xlsx.writeBuffer();

      // 添加到ZIP文件 - 使用用户编辑的文件名
      const savedFileName = editableFileNames.value.find(
        item =>
          item.groupName === groupName && item.sheetType === sheets[0].sheetType
      );
      const userFileName = savedFileName ? savedFileName.fileName : groupName;
      const finalFileName = userFileName.endsWith(".xlsx")
        ? userFileName
        : `${userFileName}.xlsx`;

      console.log(
        `使用文件名: ${finalFileName} (原始分组: ${groupName}, 用户输入: ${userFileName})`
      );
      zip.file(finalFileName, excelBuffer);
    }

    // 生成ZIP文件
    const zipBuffer = await zip.generateAsync({ type: "array" });

    // 下载ZIP文件
    const zipBlob = new Blob([new Uint8Array(zipBuffer)], {
      type: "application/zip"
    });
    const fileName = `国内酒店账单_${uploadedFile.value?.name.replace(".xlsx", "").replace(".xls", "")}_${new Date().toISOString().slice(0, 10)}.zip`;
    saveAs(zipBlob, fileName);

    ElMessage.success(
      `成功生成 ${groupInfo.length} 个Excel文件并打包为ZIP文件`
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
            <el-table-column prop="sheetName" label="工作表" width="120" />
            <el-table-column prop="groupName" label="分组名称" width="200" />
            <el-table-column prop="count" label="数据条数" width="120" />
            <el-table-column prop="rowRange" label="行号范围" width="150" />
            <el-table-column label="生成文件名">
              <template #default="scope">
                <el-input
                  :model-value="scope.row.editableFileName"
                  @update:model-value="
                    value =>
                      updateFileName(
                        scope.row.groupName,
                        value,
                        scope.row.sheetType
                      )
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
