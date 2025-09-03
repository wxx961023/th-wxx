<script setup lang="ts">
import { ref, computed } from "vue";
import {
  ElMessage,
  ElUpload,
  ElButton,
  ElCard,
  ElProgress,
  ElTable,
  ElTableColumn,
  ElTag,
  ElInput,
  ElDivider,
  ElSpace,
  ElAlert,
  ElIcon
} from "element-plus";
import {
  Upload,
  Download,
  Delete,
  FolderOpened,
  DocumentCopy,
  Warning
} from "@element-plus/icons-vue";
import type { UploadProps } from "element-plus";
import * as pdfjsLib from "pdfjs-dist";
import JSZip from "jszip";
import { saveAs } from "file-saver";

// 设置PDF.js的worker路径，使用本地worker文件
pdfjsLib.GlobalWorkerOptions.workerSrc = "/pdf.worker.min.mjs";

defineOptions({
  name: "PdfBatchRename"
});

// 文件处理状态
interface FileItem {
  id: string;
  originalName: string;
  newName: string;
  status: "pending" | "processing" | "success" | "error";
  progress: number;
  errorMessage?: string;
  companyName?: string;
  expenseType?: string;
  amount?: string;
  file?: File;
}

// 批量重命名规则
interface BatchRenameRule {
  prefix: string;
  suffix: string;
  useSequence: boolean;
  sequenceStart: number;
  sequencePadding: number;
}

// 响应式数据
const fileList = ref<FileItem[]>([]);
const isProcessing = ref(false);
const isExtracting = ref(false);
const zipFile = ref<File | null>(null);
const batchRenameRule = ref<BatchRenameRule>({
  prefix: "",
  suffix: "",
  useSequence: false,
  sequenceStart: 1,
  sequencePadding: 2
});

// 计算属性：检测重复文件名
const duplicateFileNames = computed(() => {
  const nameCount = new Map<string, number>();
  const duplicates = new Set<string>();

  // 统计每个文件名出现的次数
  fileList.value.forEach(item => {
    if (item.newName && item.newName.trim()) {
      const name = item.newName.toLowerCase(); // 不区分大小写
      nameCount.set(name, (nameCount.get(name) || 0) + 1);
    }
  });

  // 找出重复的文件名
  nameCount.forEach((count, name) => {
    if (count > 1) {
      duplicates.add(name);
    }
  });

  return duplicates;
});

// 检查指定文件是否有重复文件名
const isFileNameDuplicate = (fileName: string): boolean => {
  if (!fileName || !fileName.trim()) return false;
  return duplicateFileNames.value.has(fileName.toLowerCase());
};

// 费用类型映射规则
const expenseTypeRules = [
  { keywords: ["代订火车票"], type: "火车费" },
  { keywords: ["代订退票费"], type: "退票费" },
  { keywords: ["代订机票"], type: "机票费" },
  { keywords: ["服务费"], type: "服务费" }
];

// 从文件名提取公司名称
const extractCompanyName = (filename: string): string => {
  const parts = filename.split("_");
  if (parts.length >= 3) {
    return parts[2]; // 第三个下划线分隔的部分
  }
  return "";
};

// 从PDF内容确定费用类型
const determineExpenseType = (pdfText: string): string => {
  for (const rule of expenseTypeRules) {
    if (rule.keywords.some(keyword => pdfText.includes(keyword))) {
      return rule.type;
    }
  }
  return "机票费"; // 默认类型
};

// 从PDF内容提取金额
const extractAmount = (pdfText: string): string => {
  console.log("pdfText", pdfText);

  // 使用您建议的过滤规则：提取开票人之前的内容
  let targetText = pdfText;
  const invoicerMatch = pdfText.match(/(.*?)王欣欣[：:]?/);
  if (invoicerMatch) {
    targetText = invoicerMatch[1];
    console.log("提取到开票人之前的内容:", targetText);
  }

  // 直接在原文本中查找所有可能的金额模式
  const allAmounts = [];

  // 模式1: 查找所有 ¥ 后面的数字（包括空格分隔的）
  const yuanMatches = targetText.match(/¥\s*[\d\s\.]+/g);
  if (yuanMatches) {
    console.log("找到¥符号匹配:", yuanMatches);
    yuanMatches.forEach(match => {
      // 提取并清理数字
      const numberPart = match.replace(/¥\s*/, "").replace(/\s+/g, "");
      const amount = parseFloat(numberPart);
      if (!isNaN(amount) && amount > 0 && amount < 1000000) {
        allAmounts.push({
          value: amount,
          text: numberPart,
          source: `¥符号: ${match}`
        });
        console.log(`提取金额: ${numberPart} 来源: ${match}`);
      }
    });
  }

  // 模式2: 查找大写金额后的数字
  const chineseAmountMatches = targetText.match(
    /[壹贰叁肆伍陆柒捌玖拾佰仟万亿圆整]+\s*¥?\s*[\d\s\.]+/g
  );
  if (chineseAmountMatches) {
    console.log("找到大写金额匹配:", chineseAmountMatches);
    chineseAmountMatches.forEach(match => {
      const numberMatch = match.match(/[\d\s\.]+$/);
      if (numberMatch) {
        const numberPart = numberMatch[0].replace(/\s+/g, "");
        const amount = parseFloat(numberPart);
        if (!isNaN(amount) && amount > 0 && amount < 1000000) {
          allAmounts.push({
            value: amount,
            text: numberPart,
            source: `大写金额: ${match}`
          });
          console.log(`提取金额: ${numberPart} 来源: ${match}`);
        }
      }
    });
  }

  // 模式3: 查找价税合计相关
  const totalMatches = targetText.match(/价税合计[^¥]*¥?\s*[\d\s\.]+/g);
  if (totalMatches) {
    console.log("找到价税合计匹配:", totalMatches);
    totalMatches.forEach(match => {
      const numberMatch = match.match(/[\d\s\.]+$/);
      if (numberMatch) {
        const numberPart = numberMatch[0].replace(/\s+/g, "");
        const amount = parseFloat(numberPart);
        if (!isNaN(amount) && amount > 0 && amount < 1000000) {
          allAmounts.push({
            value: amount,
            text: numberPart,
            source: `价税合计: ${match}`
          });
          console.log(`提取金额: ${numberPart} 来源: ${match}`);
        }
      }
    });
  }

  console.log("所有找到的金额:", allAmounts);

  // 选择最大的金额作为总金额
  if (allAmounts.length > 0) {
    const maxAmount = allAmounts.reduce((max, current) =>
      current.value > max.value ? current : max
    );
    console.log(`选择最大金额: ${maxAmount.text} (${maxAmount.source})`);
    return maxAmount.text;
  }

  console.log("未找到匹配的金额");
  return "0";
};

// 清理文件名中的非法字符
const sanitizeFileName = (filename: string): string => {
  return filename.replace(/[<>:"/\\|?*]/g, "_");
};

// 验证文件是否为PDF格式
const isPdfFile = (file: File): boolean => {
  return (
    file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf")
  );
};

// 处理ZIP文件解压
const extractZipFile = async (zipFileData: File): Promise<FileItem[]> => {
  try {
    isExtracting.value = true;
    const zip = new JSZip();
    const zipContent = await zip.loadAsync(zipFileData);
    const extractedFiles: FileItem[] = [];

    // 遍历ZIP文件中的所有文件
    for (const [filename, zipEntry] of Object.entries(zipContent.files)) {
      // 跳过文件夹
      if (zipEntry.dir) continue;

      // 只处理PDF文件
      if (!filename.toLowerCase().endsWith(".pdf")) {
        console.log(`跳过非PDF文件: ${filename}`);
        continue;
      }

      try {
        // 获取文件内容
        const fileData = await zipEntry.async("blob");
        const file = new File([fileData], filename, {
          type: "application/pdf"
        });

        // 验证是否为有效的PDF文件
        if (isPdfFile(file)) {
          const fileItem: FileItem = {
            id:
              Date.now().toString() +
              Math.random().toString(36).substring(2, 11),
            originalName: filename,
            newName: "",
            status: "pending",
            progress: 0,
            file: file
          };
          extractedFiles.push(fileItem);
        }
      } catch (error) {
        console.error(`处理文件 ${filename} 时出错:`, error);
      }
    }

    return extractedFiles;
  } catch (error) {
    console.error("解压ZIP文件失败:", error);
    throw new Error("ZIP文件解压失败");
  } finally {
    isExtracting.value = false;
  }
};

// 处理单个PDF文件
const processPdfFile = async (fileItem: FileItem): Promise<void> => {
  try {
    fileItem.status = "processing";
    fileItem.progress = 10;

    if (!fileItem.file) {
      throw new Error("文件不存在");
    }

    // 读取PDF文件内容
    const arrayBuffer = await fileItem.file.arrayBuffer();
    fileItem.progress = 30;

    // 使用PDF.js解析PDF内容
    const loadingTask = pdfjsLib.getDocument({
      data: arrayBuffer,
      useWorkerFetch: false,
      isEvalSupported: false,
      useSystemFonts: true
    });
    const pdf = await loadingTask.promise;
    fileItem.progress = 50;

    let pdfText = "";
    // 提取所有页面的文本
    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const textContent = await page.getTextContent();
      const pageText = textContent.items.map((item: any) => item.str).join(" ");
      pdfText += pageText + " ";
    }
    fileItem.progress = 70;

    // 提取信息
    const companyName = extractCompanyName(fileItem.originalName);
    const expenseType = determineExpenseType(pdfText);
    const amount = extractAmount(pdfText);

    fileItem.companyName = companyName;
    fileItem.expenseType = expenseType;
    fileItem.amount = amount;
    fileItem.progress = 90;

    // 生成新文件名
    const newFileName = sanitizeFileName(
      `${companyName}_${expenseType}_${amount}.pdf`
    );
    fileItem.newName = newFileName;
    fileItem.progress = 100;
    fileItem.status = "success";
  } catch (error) {
    console.log("error: ", error);
    fileItem.status = "error";
    fileItem.errorMessage = error instanceof Error ? error.message : "处理失败";
    fileItem.progress = 0;
  }
};

// 文件上传前的检查
const beforeUpload: UploadProps["beforeUpload"] = file => {
  const isPdf = file.type === "application/pdf";
  const isZip =
    file.type === "application/zip" ||
    file.type === "application/x-zip-compressed" ||
    file.name.toLowerCase().endsWith(".zip");
  const isLt50M = file.size / 1024 / 1024 < 50; // 增加到50MB以支持ZIP文件

  if (!isPdf && !isZip) {
    ElMessage.error("只能上传PDF文件或ZIP压缩包!");
    return false;
  }
  if (!isLt50M) {
    ElMessage.error("文件大小不能超过50MB!");
    return false;
  }
  return true;
};

// 文件选择处理
const handleFileChange: UploadProps["onChange"] = async uploadFile => {
  if (uploadFile.raw) {
    const file = uploadFile.raw;
    const isZip =
      file.type === "application/zip" ||
      file.type === "application/x-zip-compressed" ||
      file.name.toLowerCase().endsWith(".zip");

    if (isZip) {
      // 处理ZIP文件
      try {
        zipFile.value = file;
        ElMessage.info("正在解压ZIP文件，请稍候...");
        const extractedFiles = await extractZipFile(file);

        if (extractedFiles.length === 0) {
          ElMessage.warning("ZIP文件中没有找到PDF文件");
          return;
        }

        fileList.value.push(...extractedFiles);
        ElMessage.success(
          `成功从ZIP文件中提取了 ${extractedFiles.length} 个PDF文件`
        );
      } catch (error) {
        ElMessage.error(
          "ZIP文件处理失败: " +
            (error instanceof Error ? error.message : "未知错误")
        );
      }
    } else {
      // 处理单个PDF文件
      const fileItem: FileItem = {
        id: Date.now().toString() + Math.random().toString(36).substring(2, 11),
        originalName: uploadFile.name,
        newName: "",
        status: "pending",
        progress: 0,
        file: uploadFile.raw
      };
      fileList.value.push(fileItem);
    }
  }
};

// 批量处理所有文件
const processAllFiles = async () => {
  if (fileList.value.length === 0) {
    ElMessage.warning("请先选择PDF文件");
    return;
  }

  isProcessing.value = true;

  try {
    // 并发处理所有文件
    const promises = fileList.value
      .filter(item => item.status === "pending")
      .map(item => processPdfFile(item));

    await Promise.all(promises);

    const successCount = fileList.value.filter(
      item => item.status === "success"
    ).length;
    const errorCount = fileList.value.filter(
      item => item.status === "error"
    ).length;

    if (errorCount === 0) {
      ElMessage.success(`成功处理 ${successCount} 个文件`);
    } else {
      ElMessage.warning(
        `处理完成：成功 ${successCount} 个，失败 ${errorCount} 个`
      );
    }
  } catch (error) {
    ElMessage.error("批量处理失败");
  } finally {
    isProcessing.value = false;
  }
};

// 下载重命名后的文件
const downloadFile = (fileItem: FileItem) => {
  if (!fileItem.file || fileItem.status !== "success") {
    ElMessage.error("文件未准备好");
    return;
  }

  const url = URL.createObjectURL(fileItem.file);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileItem.newName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
};

// 删除文件
const removeFile = (fileItem: FileItem) => {
  const index = fileList.value.findIndex(item => item.id === fileItem.id);
  if (index > -1) {
    fileList.value.splice(index, 1);
  }
};

// 清空所有文件
const clearAllFiles = () => {
  fileList.value = [];
  zipFile.value = null;
};

// 应用批量重命名规则
const applyBatchRename = () => {
  if (fileList.value.length === 0) {
    ElMessage.warning("没有文件可以重命名");
    return;
  }

  fileList.value.forEach((fileItem, index) => {
    let newName = fileItem.originalName;

    // 移除原始扩展名
    const nameWithoutExt = newName.replace(/\.pdf$/i, "");

    // 应用前缀
    if (batchRenameRule.value.prefix) {
      newName = batchRenameRule.value.prefix + nameWithoutExt;
    } else {
      newName = nameWithoutExt;
    }

    // 应用序号
    if (batchRenameRule.value.useSequence) {
      const sequenceNumber = (batchRenameRule.value.sequenceStart + index)
        .toString()
        .padStart(batchRenameRule.value.sequencePadding, "0");
      newName = newName + "_" + sequenceNumber;
    }

    // 应用后缀
    if (batchRenameRule.value.suffix) {
      newName = newName + "_" + batchRenameRule.value.suffix;
    }

    // 添加PDF扩展名
    newName = sanitizeFileName(newName + ".pdf");
    fileItem.newName = newName;
  });

  ElMessage.success("批量重命名规则已应用");
};

// 验证文件名是否有重复
const validateFileNames = (): boolean => {
  const names = fileList.value.map(item => item.newName).filter(name => name);
  const uniqueNames = new Set(names.map(name => name.toLowerCase()));

  if (names.length !== uniqueNames.size) {
    ElMessage.error("存在重复的文件名，请检查重命名规则");
    return false;
  }

  return true;
};

// 下载重新打包的ZIP文件
const downloadRenamedZip = async () => {
  if (fileList.value.length === 0) {
    ElMessage.error("没有文件可以下载");
    return;
  }

  // 验证所有文件都有新文件名
  const unnamedFiles = fileList.value.filter(item => !item.newName);
  if (unnamedFiles.length > 0) {
    ElMessage.error("请先为所有文件设置新文件名");
    return;
  }

  // 验证文件名不重复
  if (!validateFileNames()) {
    return;
  }

  try {
    ElMessage.info("正在打包文件，请稍候...");
    const zip = new JSZip();

    // 将所有文件添加到ZIP中
    for (const fileItem of fileList.value) {
      if (fileItem.file && fileItem.newName) {
        zip.file(fileItem.newName, fileItem.file);
      }
    }

    // 生成ZIP文件
    const zipBlob = await zip.generateAsync({ type: "blob" });

    // 下载文件
    const originalZipName = zipFile.value?.name || "files.zip";
    const newZipName = originalZipName.replace(/\.zip$/i, "_renamed.zip");
    saveAs(zipBlob, newZipName);

    ElMessage.success("文件打包下载成功");
  } catch (error) {
    console.error("打包下载失败:", error);
    ElMessage.error("文件打包失败");
  }
};

// 获取状态标签类型
const getStatusTagType = (status: string) => {
  switch (status) {
    case "success":
      return "success";
    case "error":
      return "danger";
    case "processing":
      return "warning";
    default:
      return "info";
  }
};

// 获取状态文本
const getStatusText = (status: string) => {
  switch (status) {
    case "pending":
      return "待处理";
    case "processing":
      return "处理中";
    case "success":
      return "成功";
    case "error":
      return "失败";
    default:
      return "未知";
  }
};

// 获取表格行的样式类名
const getRowClassName = ({ row }: { row: FileItem }) => {
  if (row.newName && isFileNameDuplicate(row.newName)) {
    return "duplicate-filename-row";
  }
  return "";
};
</script>

<template>
  <div class="pdf-rename-container p-6">
    <el-card class="mb-6">
      <template #header>
        <div class="flex justify-between items-center">
          <h2 class="text-xl font-bold">PDF批量重命名工具</h2>
          <div class="space-x-2">
            <el-button
              type="primary"
              :icon="Upload"
              @click="processAllFiles"
              :loading="isProcessing"
              :disabled="fileList.length === 0"
            >
              {{ isProcessing ? "处理中..." : "开始处理" }}
            </el-button>
            <el-button
              type="success"
              :icon="Download"
              @click="downloadRenamedZip"
              :disabled="
                fileList.length === 0 || fileList.some(f => !f.newName)
              "
            >
              下载重命名后的ZIP
            </el-button>
            <el-button
              type="danger"
              :icon="Delete"
              @click="clearAllFiles"
              :disabled="fileList.length === 0"
            >
              清空列表
            </el-button>
          </div>
        </div>
      </template>

      <div class="upload-section mb-6">
        <el-upload
          class="upload-demo"
          drag
          multiple
          :auto-upload="false"
          :before-upload="beforeUpload"
          :on-change="handleFileChange"
          :show-file-list="false"
          accept=".pdf,.zip"
        >
          <div class="upload-content text-center py-8">
            <el-icon class="el-icon--upload text-4xl mb-4">
              <FolderOpened v-if="isExtracting" />
              <Upload v-else />
            </el-icon>
            <div class="el-upload__text text-lg">
              <span v-if="isExtracting">正在解压ZIP文件...</span>
              <span v-else
                >将ZIP压缩包或PDF文件拖拽到此处，或<em>点击选择文件</em></span
              >
            </div>
            <div class="el-upload__tip text-sm text-gray-500 mt-2">
              支持上传ZIP压缩包（包含多个PDF）或单个PDF文件，文件大小不超过50MB
            </div>
          </div>
        </el-upload>
      </div>

      <!-- 批量重命名规则设置 -->
      <div class="batch-rename-section mb-6" v-if="fileList.length > 0">
        <el-divider content-position="left"> </el-divider>
      </div>

      <div class="file-list-section" v-if="fileList.length > 0">
        <h3 class="text-lg font-semibold mb-4">
          文件列表 ({{ fileList.length }})
        </h3>
        <el-table
          :data="fileList"
          style="width: 100%"
          stripe
          :row-class-name="getRowClassName"
        >
          <el-table-column prop="originalName" label="原文件名" min-width="200">
            <template #default="{ row }">
              <div class="truncate" :title="row.originalName">
                {{ row.originalName }}
              </div>
            </template>
          </el-table-column>

          <el-table-column prop="expenseType" label="费用类型" width="100">
            <template #default="{ row }">
              <span v-if="row.expenseType">{{ row.expenseType }}</span>
              <span v-else class="text-gray-400">-</span>
            </template>
          </el-table-column>

          <el-table-column prop="amount" label="金额" width="100">
            <template #default="{ row }">
              <span v-if="row.amount">{{ row.amount }}</span>
              <span v-else class="text-gray-400">-</span>
            </template>
          </el-table-column>

          <el-table-column prop="status" label="状态" width="80">
            <template #default="{ row }">
              <el-tag :type="getStatusTagType(row.status)" size="small">
                {{ getStatusText(row.status) }}
              </el-tag>
            </template>
          </el-table-column>

          <el-table-column label="进度" width="70">
            <template #default="{ row }">
              <el-progress
                v-if="row.status === 'processing'"
                :percentage="row.progress"
                :stroke-width="6"
                size="small"
              />
              <span v-else-if="row.status === 'success'" class="text-green-500"
                >完成</span
              >
              <span v-else-if="row.status === 'error'" class="text-red-500"
                >失败</span
              >
              <span v-else class="text-gray-400">待处理</span>
            </template>
          </el-table-column>

          <el-table-column prop="newName" label="新文件名" min-width="320">
            <template #default="{ row }">
              <div class="flex items-center gap-2">
                <el-input
                  v-model="row.newName"
                  placeholder="输入新文件名（不含扩展名）"
                  clearable
                  :class="{
                    'duplicate-filename-input': isFileNameDuplicate(row.newName)
                  }"
                  @input="
                    (value: string) => {
                      if (value && !value.endsWith('.pdf')) {
                        row.newName = sanitizeFileName(value + '.pdf');
                      }
                    }
                  "
                >
                  <template #suffix>
                    <span class="text-gray-400">.pdf</span>
                  </template>
                </el-input>
                <el-icon
                  v-if="isFileNameDuplicate(row.newName)"
                  class="text-red-500 flex-shrink-0"
                  :title="'文件名重复'"
                >
                  <Warning />
                </el-icon>
              </div>
            </template>
          </el-table-column>

          <el-table-column label="操作" width="200" fixed="right">
            <template #default="{ row }">
              <div class="space-x-2">
                <el-button
                  v-if="row.status === 'success'"
                  type="primary"
                  size="small"
                  :icon="Download"
                  @click="downloadFile(row)"
                >
                  下载
                </el-button>
                <el-button
                  type="danger"
                  size="small"
                  :icon="Delete"
                  @click="removeFile(row)"
                >
                  删除
                </el-button>
              </div>
            </template>
          </el-table-column>
        </el-table>
      </div>
    </el-card>
  </div>
</template>

<style scoped>
.pdf-rename-container {
  max-width: 1600px;
  margin: 0 auto;
}

.upload-demo :deep(.el-upload-dragger) {
  border: 2px dashed #d9d9d9;
  border-radius: 6px;
  width: 100%;
  height: auto;
  text-align: center;
  background-color: #fafafa;
  transition: border-color 0.3s;
}

.upload-demo :deep(.el-upload-dragger:hover) {
  border-color: #409eff;
}

.upload-content {
  padding: 2rem;
}

.truncate {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.space-x-2 > * + * {
  margin-left: 0.5rem;
}

.space-y-1 > * + * {
  margin-top: 0.25rem;
}

.space-y-2 > * + * {
  margin-top: 0.5rem;
}

/* 重复文件名行的样式 */
:deep(.duplicate-filename-row) {
  background-color: #fef2f2 !important;
  border-left: 4px solid #ef4444;
}

:deep(.duplicate-filename-row:hover) {
  background-color: #fee2e2 !important;
}

/* 重复文件名输入框的样式 */
.duplicate-filename-input :deep(.el-input__wrapper) {
  border-color: #ef4444;
  box-shadow: 0 0 0 1px #ef4444 inset;
}

.duplicate-filename-input :deep(.el-input__wrapper:hover) {
  border-color: #dc2626;
  box-shadow: 0 0 0 1px #dc2626 inset;
}

.duplicate-filename-input :deep(.el-input__wrapper.is-focus) {
  border-color: #ef4444;
  box-shadow: 0 0 0 1px #ef4444 inset;
}
</style>
