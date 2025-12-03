<script setup lang="ts">
import { ref } from "vue";
import { ElMessage } from "element-plus";
import { UploadFilled } from "@element-plus/icons-vue";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import JSZip from "jszip";

defineOptions({
  name: "FzjtBillSplit"
});

// 工作表配置
const sheetProcessors = [
  { name: "国内机票", key: "flight", personKeyword: "乘机人" },
  { name: "国内火车票", key: "train", personKeyword: "乘车人" },
  { name: "国内酒店", key: "hotel", personKeyword: "入住人" },
  { name: "通用产品", key: "general", personKeyword: "出差人" }
];

const uploadedFile = ref<File | null>(null);
const allSheetData = ref<
  Record<
    string,
    {
      headers: any[];
      data: any[][];
      personColIndex: number;
      productTypeColIndex?: number; // 仅通用产品工作表需要
      totalAmountColIndex?: number; // 总金额列索引
    }
  >
>({});
const loading = ref(false);
const showData = ref(false);
const generating = ref(false);

// 分组结果
interface PersonGroup {
  personName: string;
  sheetData: Record<string, any[][]>;
  totalCount: number;
  editableFileName: string;
}
const personGroups = ref<PersonGroup[]>([]);

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

      const sheetData: Record<
        string,
        {
          headers: any[];
          data: any[][];
          personColIndex: number;
          productTypeColIndex?: number;
          totalAmountColIndex?: number;
        }
      > = {};
      let totalSheets = 0;

      // 处理每个工作表
      for (const processor of sheetProcessors) {
        const worksheet = workbook.getWorksheet(processor.name);
        if (!worksheet) {
          console.log(`跳过不存在的工作表: ${processor.name}`);
          continue;
        }

        totalSheets++;
        const rows: any[][] = [];
        worksheet.eachRow(row => {
          const rowData: any[] = [];
          row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            rowData[colNumber - 1] = cell.value;
          });
          rows.push(rowData);
        });

        if (rows.length === 0) continue;

        // 第一行是表头
        const headers = rows[0];
        const personColIndex = headers.findIndex(
          h => h && h.toString().includes(processor.personKeyword)
        );

        if (personColIndex === -1) {
          console.warn(
            `工作表 ${processor.name} 未找到 "${processor.personKeyword}" 列`
          );
          continue;
        }

        // 对于通用产品工作表，查找"产品类型"列的索引
        let productTypeColIndex: number | undefined;
        if (processor.key === "general") {
          productTypeColIndex = headers.findIndex(
            h => h && h.toString().includes("产品类型")
          );
          if (productTypeColIndex === -1) {
            console.warn(`工作表 ${processor.name} 未找到 "产品类型" 列`);
            productTypeColIndex = undefined;
          } else {
            console.log(
              `${processor.name}: 产品类型列索引: ${productTypeColIndex}`
            );
          }
        }

        // 查找"总金额"列的索引
        let totalAmountColIndex: number | undefined;
        const totalAmountIdx = headers.findIndex(
          h => h && h.toString().includes("总金额")
        );
        if (totalAmountIdx !== -1) {
          totalAmountColIndex = totalAmountIdx;
          console.log(
            `${processor.name}: 总金额列索引: ${totalAmountColIndex}`
          );
        } else {
          console.warn(`工作表 ${processor.name} 未找到 "总金额" 列`);
        }

        sheetData[processor.key] = {
          headers,
          data: rows.slice(1), // 数据行（不含表头）
          personColIndex,
          productTypeColIndex,
          totalAmountColIndex
        };

        console.log(
          `${processor.name}: ${rows.length - 1} 条数据, 人员列索引: ${personColIndex}`
        );
      }

      if (totalSheets === 0) {
        ElMessage.error("未找到任何需要处理的工作表");
        loading.value = false;
        return;
      }

      allSheetData.value = sheetData;
      processGroups();
      showData.value = true;
      loading.value = false;
      ElMessage.success(`成功读取 ${totalSheets} 个工作表！`);
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

// 处理分组逻辑
const processGroups = () => {
  const groups = new Map<string, Record<string, any[][]>>();

  // 遍历每个工作表
  for (const processor of sheetProcessors) {
    const sheet = allSheetData.value[processor.key];
    if (!sheet) continue;

    // 遍历数据行
    for (const row of sheet.data) {
      const personName = row[sheet.personColIndex]?.toString().trim();
      if (!personName || personName === "") continue;

      if (!groups.has(personName)) {
        groups.set(personName, {});
      }

      const personData = groups.get(personName)!;
      if (!personData[processor.key]) {
        personData[processor.key] = [];
      }
      personData[processor.key].push(row);
    }
  }

  // 转换为数组
  personGroups.value = Array.from(groups.entries()).map(
    ([personName, sheetData]) => {
      let totalCount = 0;
      for (const key of Object.keys(sheetData)) {
        totalCount += sheetData[key].length;
      }
      return {
        personName,
        sheetData,
        totalCount,
        editableFileName: personName // 文件名仅使用人名，不包含日期范围
      };
    }
  );

  console.log(`共分为 ${personGroups.value.length} 组`);
};

// 重新排列行数据，使总金额列对齐到基准位置
const alignRowToTotalAmount = (
  row: any[],
  currentTotalAmountIdx: number | undefined,
  baseTotalAmountPosition: number,
  totalColumns: number
): any[] => {
  // 创建一个足够长的空数组
  const alignedRow = new Array(totalColumns).fill("");

  if (currentTotalAmountIdx === undefined) {
    // 如果当前工作表没有总金额列，直接填充原数据
    for (let i = 0; i < row.length && i < totalColumns; i++) {
      alignedRow[i] = row[i];
    }
    return alignedRow;
  }

  // 计算需要的偏移量
  const offset = baseTotalAmountPosition - currentTotalAmountIdx;

  if (offset === 0) {
    // 不需要偏移，直接填充
    for (let i = 0; i < row.length && i < totalColumns; i++) {
      alignedRow[i] = row[i];
    }
  } else if (offset > 0) {
    // 需要向右偏移：在总金额列之前插入空列
    for (let i = 0; i < row.length; i++) {
      if (i < currentTotalAmountIdx) {
        // 总金额列之前的列保持原位置
        alignedRow[i] = row[i];
      } else {
        // 总金额列及之后的列向右偏移
        const newIdx = i + offset;
        if (newIdx < totalColumns) {
          alignedRow[newIdx] = row[i];
        }
      }
    }
  } else {
    // offset < 0，需要向左偏移（较少见的情况）
    for (let i = 0; i < row.length; i++) {
      const newIdx = i + offset;
      if (newIdx >= 0 && newIdx < totalColumns) {
        alignedRow[newIdx] = row[i];
      }
    }
  }

  return alignedRow;
};

// 生成单个Excel文件
const generateExcelForPerson = async (group: PersonGroup): Promise<Blob> => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("账单明细");

  // 获取基准位置：国内机票的总金额列位置（加1是因为要加上"类别"列）
  const flightSheet = allSheetData.value["flight"];
  const baseTotalAmountPosition =
    flightSheet?.totalAmountColIndex !== undefined
      ? flightSheet.totalAmountColIndex + 1
      : undefined;

  // 计算最大列数（所有工作表中最大的列数 + 1（类别列））
  // 以国内机票的列数为基准，不需要额外扩展
  let maxColumns = 0;
  for (const processor of sheetProcessors) {
    const sheet = allSheetData.value[processor.key];
    if (sheet) {
      const sheetColumns = sheet.headers.length + 1; // +1 for 类别列
      maxColumns = Math.max(maxColumns, sheetColumns);
    }
  }

  console.log(
    `基准总金额列位置: ${baseTotalAmountPosition}, 最大列数: ${maxColumns}`
  );

  let currentRow = 1;
  let isFirstSheet = true; // 标记是否是第一个工作表
  let totalAmount = 0; // 用于累计总金额

  // 按顺序处理每个工作表的数据
  for (const processor of sheetProcessors) {
    const sheet = allSheetData.value[processor.key];
    const personRows = group.sheetData[processor.key];

    if (!sheet || !personRows || personRows.length === 0) continue;

    // 计算当前工作表总金额列的新位置（加1是因为类别列）
    const currentTotalAmountIdx =
      sheet.totalAmountColIndex !== undefined
        ? sheet.totalAmountColIndex + 1
        : undefined;

    // 对齐表头行
    const headerWithCategory = ["类别", ...sheet.headers];
    let alignedHeader: any[];
    if (baseTotalAmountPosition !== undefined) {
      alignedHeader = alignRowToTotalAmount(
        headerWithCategory,
        currentTotalAmountIdx,
        baseTotalAmountPosition, // baseTotalAmountPosition 已经包含了类别列的偏移
        maxColumns
      );
    } else {
      alignedHeader = headerWithCategory;
    }

    // 需求3：非第一个工作表的表头中，"总金额"列文字替换为空
    if (!isFirstSheet && baseTotalAmountPosition !== undefined) {
      alignedHeader[baseTotalAmountPosition] = "";
    }

    const headerRow = worksheet.addRow(alignedHeader);
    headerRow.height = 20; // 设置行高为20磅
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
    currentRow++;

    // 添加数据行（第一列插入类别值）
    for (const row of personRows) {
      let categoryValue: string;

      if (processor.key === "general") {
        // 通用产品：从"产品类型"列获取值
        const productTypeColIndex = sheet.productTypeColIndex;
        if (
          productTypeColIndex !== undefined &&
          row[productTypeColIndex] !== undefined
        ) {
          categoryValue = row[productTypeColIndex]?.toString() || "";
        } else {
          categoryValue = "通用产品";
        }
      } else {
        // 其他工作表：使用工作表名称作为类别
        categoryValue = processor.name;
      }

      const rowWithCategory = [categoryValue, ...row];
      let alignedRow: any[];
      if (baseTotalAmountPosition !== undefined) {
        alignedRow = alignRowToTotalAmount(
          rowWithCategory,
          currentTotalAmountIdx,
          baseTotalAmountPosition, // baseTotalAmountPosition 已经包含了类别列的偏移
          maxColumns
        );
      } else {
        alignedRow = rowWithCategory;
      }

      // 累计总金额
      if (baseTotalAmountPosition !== undefined) {
        const amountValue = alignedRow[baseTotalAmountPosition];
        if (amountValue !== undefined && amountValue !== "") {
          const numValue = parseFloat(amountValue?.toString() || "0");
          if (!isNaN(numValue)) {
            totalAmount += numValue;
          }
        }
      }

      const dataRow = worksheet.addRow(alignedRow);
      dataRow.height = 20; // 设置行高为20磅
      dataRow.eachCell(cell => {
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" }
        };
      });
      currentRow++;
    }

    // 需求1：移除空行分隔（删除原来的空行添加代码）
    isFirstSheet = false; // 标记已处理完第一个工作表
  }

  // 需求2：在所有数据末尾添加合计行
  if (baseTotalAmountPosition !== undefined) {
    const summaryRow = new Array(maxColumns).fill("");
    summaryRow[0] = "合计";
    summaryRow[baseTotalAmountPosition] = totalAmount.toFixed(2);

    const summaryRowExcel = worksheet.addRow(summaryRow);
    summaryRowExcel.height = 20; // 设置行高为20磅
    summaryRowExcel.eachCell(cell => {
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

  // 自动调整列宽
  worksheet.columns.forEach(column => {
    let maxLength = 10;
    column.eachCell?.({ includeEmpty: true }, cell => {
      const cellValue = cell.value?.toString() || "";
      maxLength = Math.max(maxLength, cellValue.length * 2);
    });
    column.width = Math.min(maxLength, 30);
  });

  const buffer = await workbook.xlsx.writeBuffer();
  return new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
};

// 生成并下载所有文件
const generateAllFiles = async () => {
  if (personGroups.value.length === 0) {
    ElMessage.warning("没有可导出的数据");
    return;
  }

  generating.value = true;

  try {
    if (personGroups.value.length === 1) {
      // 只有一个人，直接下载Excel
      const group = personGroups.value[0];
      const blob = await generateExcelForPerson(group);
      saveAs(blob, `${group.editableFileName}.xlsx`);
      ElMessage.success("文件生成成功！");
    } else {
      // 多个人，打包成ZIP
      const zip = new JSZip();

      for (const group of personGroups.value) {
        const blob = await generateExcelForPerson(group);
        zip.file(`${group.editableFileName}.xlsx`, blob);
      }

      const zipBlob = await zip.generateAsync({ type: "blob" });
      saveAs(zipBlob, `纺织集团账单拆分.zip`);
      ElMessage.success(`成功生成 ${personGroups.value.length} 个文件！`);
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
  personGroups.value[index].editableFileName = newName;
};
</script>

<template>
  <div class="fzjt-bill-split">
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

    <!-- 分组结果 -->
    <el-card v-if="showData && personGroups.length > 0" class="result-card">
      <template #header>
        <div class="card-header">
          <span>分组结果（共 {{ personGroups.length }} 人）</span>
          <el-button
            type="primary"
            :loading="generating"
            @click="generateAllFiles"
          >
            {{ generating ? "生成中..." : "生成并下载" }}
          </el-button>
        </div>
      </template>

      <el-table :data="personGroups" border stripe>
        <el-table-column prop="personName" label="人员姓名" width="120" />
        <el-table-column label="数据统计" width="300">
          <template #default="{ row }">
            <div class="data-stats">
              <el-tag v-if="row.sheetData.flight" size="small" type="primary">
                机票: {{ row.sheetData.flight.length }}条
              </el-tag>
              <el-tag v-if="row.sheetData.train" size="small" type="success">
                火车票: {{ row.sheetData.train.length }}条
              </el-tag>
              <el-tag v-if="row.sheetData.hotel" size="small" type="warning">
                酒店: {{ row.sheetData.hotel.length }}条
              </el-tag>
              <el-tag v-if="row.sheetData.general" size="small" type="info">
                通用: {{ row.sheetData.general.length }}条
              </el-tag>
            </div>
          </template>
        </el-table-column>
        <el-table-column prop="totalCount" label="总条数" width="100" />
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
      v-if="showData && personGroups.length === 0"
      description="未找到有效数据"
    />
  </div>
</template>

<style scoped>
.fzjt-bill-split {
  padding: 20px;
}

.upload-card,
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

.data-stats {
  display: flex;
  flex-wrap: wrap;
  gap: 5px;
}

:deep(.el-upload-dragger) {
  width: 100%;
}
</style>
