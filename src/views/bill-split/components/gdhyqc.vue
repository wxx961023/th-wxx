<script setup lang="ts">
import { ref } from "vue";
import { ElMessage } from "element-plus";
import { UploadFilled } from "@element-plus/icons-vue";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import JSZip from "jszip";

defineOptions({
  name: "GdhyqcBillSplit"
});

// 工作表名称
const SHEET_NAME = "机票明细(国内)";
// 表头行数（第1-3行为表头）
const HEADER_ROWS = 3;
// 数据起始行（从第4行开始）
const DATA_START_ROW = 4;
// 分组字段
const GROUP_FIELD = "开票单位";

const uploadedFile = ref<File | null>(null);
const sheetData = ref<{
  headers: any[][]; // 多行表头
  data: any[][];
  groupColIndex: number; // 开票单位列索引
} | null>(null);
const loading = ref(false);
const showData = ref(false);
const generating = ref(false);

// 分组结果
interface CompanyGroup {
  companyName: string;
  rows: any[][];
  totalCount: number;
  editableFileName: string;
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

const readFile = (file: File) => {
  loading.value = true;
  const reader = new FileReader();

  reader.onload = async e => {
    try {
      const buffer = e.target?.result as ArrayBuffer;
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);

      // 查找目标工作表
      const worksheet = workbook.getWorksheet(SHEET_NAME);
      if (!worksheet) {
        ElMessage.error(`未找到工作表: ${SHEET_NAME}`);
        loading.value = false;
        return;
      }

      // 读取多行表头（第1-3行）
      const headers: any[][] = [];
      for (let i = 1; i <= HEADER_ROWS; i++) {
        const row = worksheet.getRow(i);
        const headerRow: any[] = [];
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          headerRow[colNumber - 1] = cell.value;
        });
        headers.push(headerRow);
      }

      // 在表头中查找"开票单位"列（通常在最后一行表头中查找）
      const lastHeaderRow = headers[HEADER_ROWS - 1];
      let groupColIndex = -1;
      for (let i = 0; i < lastHeaderRow.length; i++) {
        const cellValue = lastHeaderRow[i]?.toString() || "";
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

      // 读取数据行（从第4行开始）
      const data: any[][] = [];
      const rowCount = worksheet.rowCount;
      for (let i = DATA_START_ROW; i <= rowCount; i++) {
        const row = worksheet.getRow(i);
        const rowData: any[] = [];
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          rowData[colNumber - 1] = cell.value;
        });
        // 过滤空行
        if (
          rowData.some(
            cell => cell !== null && cell !== undefined && cell !== ""
          )
        ) {
          data.push(rowData);
        }
      }

      sheetData.value = {
        headers,
        data,
        groupColIndex
      };

      console.log(`读取到 ${data.length} 条数据`);

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

// 生成单个Excel文件
const generateExcelForCompany = async (group: CompanyGroup): Promise<Blob> => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("机票明细");

  if (!sheetData.value) {
    throw new Error("数据未加载");
  }

  const { headers } = sheetData.value;

  // 添加多行表头
  for (const headerRow of headers) {
    const row = worksheet.addRow(headerRow);
    row.height = 20; // 设置行高为20磅
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

  // 添加数据行
  for (const rowData of group.rows) {
    const row = worksheet.addRow(rowData);
    row.height = 20; // 设置行高为20磅
    row.eachCell(cell => {
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
    column.width = Math.min(maxLength + 2, 50);
  });

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
            {{ generating ? "生成中..." : "生成并下载" }}
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
