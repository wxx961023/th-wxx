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
const originalWorkbook = ref<any>(null);
const loading = ref(false);
const showData = ref(false);
const generating = ref(false);

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
          // 查找名为 '酒店明细(国内)' 的工作表
          const targetSheetName = "酒店明细(国内)";
          const worksheet = workbook.getWorksheet(targetSheetName);

          if (!worksheet) {
            ElMessage.error(`未找到工作表 "${targetSheetName}"`);
            console.log(
              "可用的工作表:",
              workbook.worksheets.map(ws => ws.name)
            );
            loading.value = false;
            return;
          }

          // 读取数据为二维数组
          const jsonData: any[][] = [];
          worksheet.eachRow((row, rowNumber) => {
            const rowData: any[] = [];
            row.eachCell((cell, colNumber) => {
              rowData.push(cell.value);
            });
            jsonData.push(rowData);
          });

          excelData.value = jsonData;
          originalWorkbook.value = workbook; // 保存原始工作簿
          showData.value = true;
          loading.value = false;

          // 打印Excel文件信息
          console.log("Excel文件信息:");
          console.log("文件名:", file.name);
          console.log("文件大小:", file.size, "bytes");
          console.log("工作表数量:", workbook.worksheets.length);
          console.log(
            "所有工作表名称:",
            workbook.worksheets.map(ws => ws.name)
          );
          console.log("当前读取工作表:", targetSheetName);
          console.log("数据行数:", jsonData.length);
          console.log("数据列数:", (jsonData[0] as any[])?.length || 0);
          console.log("工作表内容:", jsonData);

          // 查找"入住人部门"所在的列
          if (jsonData.length > 2) {
            const headers = jsonData[0] as any[];
            const thirdRow = jsonData[2] as any[];

            // 在第三行中查找包含"入住人部门"的单元格
            const departmentColumnIndex = thirdRow.findIndex(
              (cell: any) => cell && cell.toString().includes("入住人部门")
            );

            if (departmentColumnIndex !== -1) {
              const departmentColumnName =
                headers[departmentColumnIndex] ||
                `第${departmentColumnIndex + 1}列`;
              const departmentData = jsonData.map((row, index) => ({
                行号: index + 1, // Excel行号（从1开始）
                入住人部门: row[departmentColumnIndex]
              }));

              console.log(`\n========== "入住人部门"列信息 ==========`);
              console.log(
                `找到"入住人部门"单元格位置: 第3行，第${departmentColumnIndex + 1}列`
              );
              console.log(`对应表头列名: ${departmentColumnName}`);
              console.log(`该列完整数据:`, departmentData);

              // 过滤掉空值，只显示有数据的行
              const validDepartmentData = departmentData.filter(
                item =>
                  item.入住人部门 && item.入住人部门.toString().trim() !== ""
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
                  const department = item.入住人部门.toString();
                  const firstPart = department.split("-")[0]; // 通过'-'拆分，取第一个部分

                  if (!groups.has(firstPart)) {
                    groups.set(firstPart, []);
                  }
                  groups.get(firstPart)!.push(item);
                });

                // 打印分组结果
                console.log(`共分为 ${groups.size} 组:`);

                groups.forEach((groupItems, groupName) => {
                  console.log(`\n--- 组名: ${groupName} ---`);
                  console.log(`包含数据条数: ${groupItems.length}`);
                  console.log(`具体数据:`);
                  groupItems.forEach(item => {
                    console.log(`  行${item.行号}: ${item.入住人部门}`);
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
                console.log("分组统计表:", groupStats);
              } else {
                console.log("\n没有第4行及以后的有效数据进行分组处理");
              }
            } else {
              console.log("第三行数据:", thirdRow);
              console.log("所有表头:", headers);
              throw new Error("在第三行中未找到'入住人部门'单元格");
            }
          }

          ElMessage.success(
            `成功读取工作表 "${targetSheetName}"！请在控制台查看详细信息`
          );
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

// 生成分组Excel文件并打包成ZIP
const generateGroupedExcelFiles = async () => {
  if (!originalWorkbook.value || excelData.value.length === 0) {
    ElMessage.error("请先上传并处理Excel文件");
    return;
  }

  generating.value = true;

  try {
    // 重新计算分组数据
    const departmentColumnIndex = (excelData.value[2] as any[]).findIndex(
      (cell: any) => cell && cell.toString().includes("入住人部门")
    );

    if (departmentColumnIndex === -1) {
      ElMessage.error("未找到入住人部门列");
      generating.value = false;
      return;
    }

    const validDataFromRow4 = excelData.value
      .slice(3)
      .filter((row: any[], index: number) => {
        return (
          row[departmentColumnIndex] &&
          row[departmentColumnIndex].toString().trim() !== ""
        );
      })
      .map((row: any[], index: number) => ({
        行号: index + 4,
        入住人部门: row[departmentColumnIndex],
        完整行数据: row
      }));

    // 分组处理
    const groups = new Map<string, typeof validDataFromRow4>();
    validDataFromRow4.forEach(item => {
      const department = item.入住人部门.toString();
      const firstPart = department.split("-")[0];
      if (!groups.has(firstPart)) {
        groups.set(firstPart, []);
      }
      groups.get(firstPart)!.push(item);
    });

    // 调试分组结果
    console.log(`有效数据总条数: ${validDataFromRow4.length}`);
    console.log(`分组数量: ${groups.size}`);
    groups.forEach((groupItems, groupName) => {
      console.log(`组名: ${groupName}, 数据条数: ${groupItems.length}`);
      console.log(
        `行号范围: ${Math.min(...groupItems.map(i => i.行号))}-${Math.max(...groupItems.map(i => i.行号))}`
      );
    });

    console.log(`准备生成 ${groups.size} 个Excel文件`);

    // 创建ZIP文件
    const zip = new JSZip();

    // 为每个组生成Excel文件
    for (const [groupName, groupItems] of groups.entries()) {
      console.log(
        `生成文件: ${groupName}.xlsx，数据条数: ${groupItems.length}`
      );
      console.log(
        `该组包含的行号:`,
        groupItems.map(item => item.行号)
      );

      // 创建新的工作簿
      const newWorkbook = new ExcelJS.Workbook();

      // 创建工作表
      const newWorksheet = newWorkbook.addWorksheet("酒店明细(国内)", {
        views: [{ showGridLines: true }]
      });

      // 设置默认行高
      newWorksheet.properties.defaultRowHeight = 40;

      // 复制原始前三行
      const headerRows = excelData.value.slice(0, 3);

      // 写入数据到工作表
      const newData = [...headerRows];
      groupItems.forEach(item => {
        newData.push(item.完整行数据);
      });

      console.log(
        `文件 ${groupName}.xlsx 总行数: ${newData.length} (前三行 + ${groupItems.length}条数据)`
      );

      // 写入数据
      newData.forEach((row, rowIndex) => {
        row.forEach((cellValue, colIndex) => {
          const cell = newWorksheet.getCell(rowIndex + 1, colIndex + 1);
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
              // 第一行：16号字体，加粗，微软雅黑
              cell.font = {
                name: "微软雅黑",
                size: 16,
                bold: true
              };
            } else if (rowIndex === 1) {
              // 第二行：11号字体，加粗，微软雅黑
              cell.font = {
                name: "微软雅黑",
                size: 11,
                bold: true
              };
            } else if (rowIndex === 2) {
              // 第三行：10号字体，加粗，微软雅黑
              cell.font = {
                name: "微软雅黑",
                size: 10,
                bold: true
              };
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
        for (
          let colIndex = 1;
          colIndex < newData[rowIndex].length;
          colIndex++
        ) {
          const currentValue = newData[rowIndex][colIndex];
          const previousValue = newData[rowIndex][startCol];

          if (
            currentValue &&
            previousValue &&
            currentValue.toString() === previousValue.toString()
          ) {
            // 相邻单元格值相等，继续检查下一个
            continue;
          } else {
            // 值不相等或为空，合并前面的相同值单元格
            if (colIndex - 1 > startCol) {
              // 有连续的相同值需要合并
              newWorksheet.mergeCells(
                rowIndex + 1,
                startCol + 1,
                rowIndex + 1,
                colIndex
              );

              // 设置合并后单元格的样式
              const mergedCell = newWorksheet.getCell(
                rowIndex + 1,
                startCol + 1
              );
              mergedCell.alignment = {
                horizontal: "center",
                vertical: "middle"
              };
            }
            startCol = colIndex;
          }
        }

        // 处理行末的最后一段相同值
        if (newData[rowIndex].length - 1 > startCol) {
          newWorksheet.mergeCells(
            rowIndex + 1,
            startCol + 1,
            rowIndex + 1,
            newData[rowIndex].length
          );

          // 设置合并后单元格的样式
          const mergedCell = newWorksheet.getCell(rowIndex + 1, startCol + 1);
          mergedCell.alignment = {
            horizontal: "center",
            vertical: "middle"
          };
        }
      }

      // 设置行高
      newWorksheet.eachRow((row, rowNumber) => {
        // 所有行都设置40磅行高
        row.height = 40; // ExcelJS中高度单位
      });

      // 自动调整列宽
      newWorksheet.columns.forEach((column, index) => {
        let maxLength = 0;
        column.eachCell((cell, rowNumber) => {
          if (cell.value) {
            const text = cell.value.toString();
            // 考虑中文字符，中文字符比英文字符宽
            const charWidth = text.split("").reduce((width, char) => {
              // 中文字符占2个宽度，英文字符占1个宽度
              return width + (char.charCodeAt(0) > 127 ? 2 : 1);
            }, 0);
            if (charWidth > maxLength) {
              maxLength = charWidth;
            }
          }
        });
        // 设置列宽，Excel中每个单位宽度大约对应一个英文字符
        // 中文字符需要更多空间，所以需要更大的系数
        column.width = Math.max(maxLength * 1.1, 15); // 最小宽度20
      });

      // 生成Excel文件内容
      const excelBuffer = await newWorkbook.xlsx.writeBuffer();

      // 添加到ZIP文件
      const fileName = groupName + ".xlsx";
      zip.file(fileName, excelBuffer);
    }

    // 生成ZIP文件
    const zipBuffer = await zip.generateAsync({ type: "array" });

    // 下载ZIP文件
    const zipBlob = new Blob([new Uint8Array(zipBuffer)], {
      type: "application/zip"
    });
    const fileName = `账单分账_${uploadedFile.value?.name.replace(".xlsx", "").replace(".xls", "")}_${new Date().toISOString().slice(0, 10)}.zip`;
    saveAs(zipBlob, fileName);

    ElMessage.success(`成功生成 ${groups.size} 个Excel文件并打包为ZIP文件`);
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
          <h3>酒店明细(国内) - 数据预览</h3>
          <div class="header-buttons">
            <el-button
              type="primary"
              @click="
                console.log('Excel文件详细信息:', {
                  fileName: uploadedFile?.name,
                  fileSize: uploadedFile?.size,
                  sheetName: '酒店明细(国内)',
                  rowCount: excelData.length,
                  columnCount: excelData[0]?.length || 0,
                  data: excelData
                })
              "
            >
              打印详细信息到控制台
            </el-button>
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
            title="数据概览"
            type="info"
            :description="`工作表 '酒店明细(国内)' 包含 ${excelData.length} 行数据，${excelData[0]?.length || 0} 列`"
            show-icon
          />
        </div>

        <div class="data-table">
          <el-table :data="excelData.slice(0, 10)" border style="width: 100%">
            <el-table-column
              v-for="(header, index) in excelData[0] || []"
              :key="index"
              :label="`列 ${index + 1}`"
              :prop="index.toString()"
            />
          </el-table>
          <p v-if="excelData.length > 10" class="data-more">
            显示前10行，共{{ excelData.length }}行数据
          </p>
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
