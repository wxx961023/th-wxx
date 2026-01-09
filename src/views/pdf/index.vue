<script setup lang="ts">
import { ref, nextTick, computed, onMounted, onUnmounted } from "vue";
import {
  ElMessage,
  ElUpload,
  ElButton,
  ElCard,
  ElProgress,
  ElTag,
  ElAlert,
  ElIcon,
  ElTable,
  ElTableColumn,
  ElInput
} from "element-plus";
import {
  Upload,
  Delete,
  Download
} from "@element-plus/icons-vue";
import type { UploadProps } from "element-plus";
import {
  parsePDFFile,
  printPDFContent,
  generateNewFileName,
  extractPDFsFromZip,
  createRenamedZip,
  type ParsedPDFContent
} from "@/utils/pdf-parser";
import { saveAs } from "file-saver";

defineOptions({
  name: "PdfParser"
});

// 文件处理状态
interface FileParseItem {
  id: string;
  fileName: string;
  originalFileName: string; // 保存原始文件名
  file: File;
  status: "pending" | "parsing" | "success" | "error";
  progress: number;
  errorMessage?: string;
  parsedContent?: ParsedPDFContent;
  newFileName?: string; // 生成的新文件名
  zipPath?: string; // 如果来自 ZIP，记录在 ZIP 中的路径
  isEditing?: boolean; // 是否正在编辑文件名
}

// 响应式数据
const fileList = ref<FileParseItem[]>([]);
const isParsing = ref(false);
const previewDialogVisible = ref(false);
const previewFileUrl = ref("");
const currentPreviewFileName = ref("");
const editInputRefs = ref<Map<string, any>>(new Map()); // 存储编辑输入框的引用
const sourceZipFileName = ref(""); // 记录上传的 ZIP 文件名

// 计算总体解析进度
const overallProgress = computed(() => {
  if (fileList.value.length === 0) return 0;

  const totalProgress = fileList.value.reduce((sum, item) => {
    if (item.status === 'success') return sum + 100;
    if (item.status === 'error') return sum + 0;
    return sum + (item.progress || 0);
  }, 0);

  return Math.floor(totalProgress / fileList.value.length);
});

// 计算成功和失败数量
const successCount = computed(() => {
  return fileList.value.filter(item => item.status === 'success').length;
});

const errorCount = computed(() => {
  return fileList.value.filter(item => item.status === 'error').length;
});

// 计算表格高度，避免页面出现滚动条
const tableHeight = computed(() => {
  const viewportHeight = window.innerHeight;

  // 计算各个区域的高度
  // 页面外边距: p-6 = 24px * 2 = 48px
  // 卡片头部（标题 + 按钮）: 约 80px
  // 上传区域: 约 180px
  // 文件列表标题: 约 50px
  // 提示信息: 约 60px
  // 进度条区域: 约 50px (如果显示)
  // 卡片内边距: 约 20px
  // 额外预留: 130px
  const reservedHeight = 48 + 80 + 180 + 50 + 60 + 50 + 20 + 250; // 约 618px

  return Math.max(300, viewportHeight - reservedHeight); // 最小高度300px
});

// 文件上传前的检查
const beforeUpload: UploadProps["beforeUpload"] = file => {
  const isPdf = file.type === "application/pdf";
  const isZip = file.type === "application/zip" || file.type === "application/x-zip-compressed";
  const isValidFile = isPdf || isZip;
  const isLt50M = file.size / 1024 / 1024 < 50;

  if (!isValidFile) {
    ElMessage.error("只能上传PDF或ZIP文件!");
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
    const fileName = uploadFile.name;

    // 检查是否是 ZIP 文件
    const isZip = file.type === "application/zip" ||
                  file.type === "application/x-zip-compressed" ||
                  fileName.toLowerCase().endsWith('.zip');

    if (isZip) {
      try {
        ElMessage.info(`正在解压 ${fileName}...`);

        // 记录 ZIP 文件名（去掉 .zip 后缀）
        const zipName = fileName.replace(/\.zip$/i, '');
        sourceZipFileName.value = zipName;

        // 从 ZIP 中提取 PDF 文件
        const pdfFiles = await extractPDFsFromZip(file, {
          onProgress: (_current, _total, message) => {
            console.log(`[ZIP 解压] ${message}`);
          }
        });

        // 将提取的 PDF 文件添加到列表
        let addedCount = 0;
        for (const pdfFile of pdfFiles) {
          const fileItem: FileParseItem = {
            id: Date.now().toString() + Math.random().toString(36).substring(2, 11),
            fileName: pdfFile.name,
            originalFileName: pdfFile.name,
            file: pdfFile.file,
            status: "pending",
            progress: 0,
            zipPath: pdfFile.path // 记录 ZIP 中的路径
          };
          fileList.value.push(fileItem);
          addedCount++;
        }

        ElMessage.success(`从 ${fileName} 中提取了 ${addedCount} 个 PDF 文件`);
      } catch (error) {
        console.error("解压 ZIP 文件失败:", error);
        ElMessage.error(`解压失败: ${error.message}`);
      }
    } else {
      // 普通的 PDF 文件
      const fileItem: FileParseItem = {
        id: Date.now().toString() + Math.random().toString(36).substring(2, 11),
        fileName: fileName,
        originalFileName: fileName,
        file: file,
        status: "pending",
        progress: 0
      };
      fileList.value.push(fileItem);
      ElMessage.success(`已添加文件: ${fileName}`);
    }
  }
};

// 解析单个PDF文件
const parseFile = async (fileItem: FileParseItem): Promise<void> => {
  const startTime = Date.now();
  console.log(`\n开始解析文件: ${fileItem.fileName} (ID: ${fileItem.id})`);

  try {
    fileItem.status = "parsing";
    fileItem.progress = 10;

    console.log(`[${fileItem.fileName}] 文件信息:`, {
      size: `${(fileItem.file.size / 1024 / 1024).toFixed(2)}MB`,
      type: fileItem.file.type,
      lastModified: new Date(fileItem.file.lastModified).toISOString()
    });

    fileItem.progress = 30;

    // 调用PDF解析工具
    const parsedContent = await parsePDFFile(fileItem.file, {
      includeSeparator: true,
      debugMode: false, // 固定为 false
      onProgress: (current, total) => {
        const baseProgress = 30;
        const progressRange = 60; // 30% 到 90%
        const currentProgress = baseProgress + (current / total) * progressRange;
        fileItem.progress = Math.min(Math.floor(currentProgress), 90);
        console.log(`[${fileItem.fileName}] 解析进度: ${current}/${total} 页 (${fileItem.progress}%)`);
      }
    });

    fileItem.progress = 90;
    fileItem.parsedContent = parsedContent;

    // 打印解析结果到控制台
    printPDFContent(parsedContent, {
      printFullText: true, // 打印完整文本
      printPageText: true, // 打印每页的文本
      maxPreviewLength: 500 // 预览长度
    });

    // 生成新文件名
    const newFileName = generateNewFileName(parsedContent, fileItem.originalFileName);
    if (newFileName) {
      fileItem.newFileName = newFileName;
      fileItem.fileName = newFileName; // 更新显示的文件名
      fileItem.progress = 100;
      fileItem.status = "success";

      const endTime = Date.now();
      console.log(`[${fileItem.fileName}] 解析成功，总耗时: ${endTime - startTime}ms`);

      ElMessage.success(`文件 "${fileItem.fileName}" 解析成功!`);
    } else {
      // 如果没有提取到姓名，标记为失败
      fileItem.newFileName = fileItem.originalFileName;
      fileItem.fileName = fileItem.originalFileName;
      fileItem.status = "error";
      fileItem.errorMessage = "未提取到姓名信息";
      fileItem.progress = 0;
      console.log(`未提取到姓名，标记为失败: ${fileItem.originalFileName}`);
      ElMessage.warning(`文件 "${fileItem.originalFileName}" 未提取到姓名信息`);
    }
  } catch (error) {
    const endTime = Date.now();
    console.error(`[${fileItem.fileName}] 解析失败，耗时: ${endTime - startTime}ms`);
    console.error(`[${fileItem.fileName}] 错误详情:`, {
      name: error.name,
      message: error.message,
      stack: error.stack
    });

    fileItem.status = "error";
    fileItem.errorMessage = error instanceof Error ? error.message : "解析失败";
    fileItem.progress = 0;

    ElMessage.error(`文件 "${fileItem.fileName}" 解析失败: ${fileItem.errorMessage}`);
  }
};

// 解析所有文件
const parseAllFiles = async () => {
  if (fileList.value.length === 0) {
    ElMessage.warning("请先选择PDF文件");
    return;
  }

  const pendingFiles = fileList.value.filter(item => item.status === "pending");

  if (pendingFiles.length === 0) {
    ElMessage.info("没有待解析的文件");
    return;
  }

  isParsing.value = true;

  try {
    console.log(`\n=== 开始批量解析 ${pendingFiles.length} 个文件 ===`);
    const startTime = Date.now();

    // 依次解析每个文件
    for (let i = 0; i < pendingFiles.length; i++) {
      const fileItem = pendingFiles[i];
      console.log(`\n[${i + 1}/${pendingFiles.length}] 处理文件: ${fileItem.fileName}`);

      await parseFile(fileItem);

      // 等待一小段时间,让控制台输出有时间刷新
      await new Promise(resolve => setTimeout(resolve, 100));
    }

    const endTime = Date.now();
    const processingTime = endTime - startTime;
    console.log(`\n=== 批量解析完成，总耗时: ${processingTime}ms ===`);

    const successCount = fileList.value.filter(item => item.status === "success").length;
    const errorCount = fileList.value.filter(item => item.status === "error").length;

    console.log("解析结果统计:", {
      success: successCount,
      error: errorCount,
      total: fileList.value.length,
      processingTime: `${processingTime}ms`
    });

    if (errorCount === 0) {
      ElMessage.success(`成功解析 ${successCount} 个文件`);
    } else {
      ElMessage.warning(`解析完成：成功 ${successCount} 个，失败 ${errorCount} 个`);
    }
  } catch (error) {
    console.error("批量解析过程中发生异常:", error);
    ElMessage.error("批量解析失败: " + (error.message || "未知错误"));
  } finally {
    isParsing.value = false;
  }
};

// 删除文件
const removeFile = (fileItem: FileParseItem) => {
  const index = fileList.value.findIndex(item => item.id === fileItem.id);
  if (index > -1) {
    fileList.value.splice(index, 1);
  }
};

// 清空所有文件
const clearAllFiles = () => {
  fileList.value = [];
};

// 下载重命名后的文件
const downloadRenamedFile = (fileItem: FileParseItem) => {
  if (!fileItem.newFileName) {
    ElMessage.warning("该文件尚未生成新文件名");
    return;
  }

  // 创建 Blob 对象
  const blob = new Blob([fileItem.file], { type: "application/pdf" });

  // 创建下载链接
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileItem.newFileName;

  // 触发下载
  document.body.appendChild(link);
  link.click();

  // 清理
  document.body.removeChild(link);
  URL.revokeObjectURL(url);

  ElMessage.success(`已下载: ${fileItem.newFileName}`);
};

// 批量导出所有重命名后的文件为 ZIP
const exportAllRenamedFiles = async () => {
  // 筛选出成功解析且有新文件名的文件
  const successFiles = fileList.value.filter(item =>
    item.status === "success" && item.newFileName
  );

  if (successFiles.length === 0) {
    ElMessage.warning("没有已成功解析的文件可以导出");
    return;
  }

  try {
    ElMessage.info(`正在打包 ${successFiles.length} 个文件...`);

    // 创建 ZIP
    const JSZip = (await import("jszip")).default;
    const zip = new JSZip();

    // 添加文件到 ZIP，保持原有目录结构
    for (const item of successFiles) {
      if (item.zipPath) {
        // 如果来自 ZIP，保持原有路径结构
        // zipPath 格式如: "folder/subfolder/file.pdf"
        // 我们需要提取目录路径，然后替换文件名为新文件名
        const pathParts = item.zipPath.split('/');

        if (pathParts.length > 1) {
          // 有子目录，保持目录结构
          const dirPath = pathParts.slice(0, -1).join('/');
          const fullPath = `${dirPath}/${item.newFileName}`;
          zip.file(fullPath, item.file);
          console.log(`导出文件: ${fullPath}`);
        } else {
          // 根目录下的文件
          zip.file(item.newFileName!, item.file);
          console.log(`导出文件: ${item.newFileName}`);
        }
      } else {
        // 直接上传的 PDF 文件（不是来自 ZIP）
        zip.file(item.newFileName!, item.file);
        console.log(`导出文件: ${item.newFileName}`);
      }
    }

    // 生成 ZIP 文件
    const zipBlob = await zip.generateAsync({
      type: "blob",
      compression: "DEFLATE",
      compressionOptions: {
        level: 6
      }
    });

    // 确定导出的 ZIP 文件名
    let exportFileName = "重命名文件.zip";
    if (sourceZipFileName.value) {
      // 如果来自 ZIP，使用原 ZIP 文件名
      exportFileName = `${sourceZipFileName.value}.zip`;
    }

    // 下载 ZIP 文件
    saveAs(zipBlob, exportFileName);

    ElMessage.success(`成功导出 ${successFiles.length} 个文件`);
  } catch (error) {
    console.error("导出失败:", error);
    ElMessage.error(`导出失败: ${error.message}`);
  }
};

// 获取状态标签类型
const getStatusTagType = (status: string) => {
  switch (status) {
    case "success":
      return "success";
    case "error":
      return "danger";
    case "parsing":
      return "warning";
    default:
      return "info";
  }
};

// 获取状态文本
const getStatusText = (status: string) => {
  switch (status) {
    case "pending":
      return "待解析";
    case "parsing":
      return "解析中";
    case "success":
      return "成功";
    case "error":
      return "失败";
    default:
      return "未知";
  }
};

// 开始编辑文件名
const startEditFileName = (fileItem: FileParseItem) => {
  fileItem.isEditing = true;
  // 自动聚焦输入框
  nextTick(() => {
    const inputRef = editInputRefs.value.get(fileItem.id);
    if (inputRef) {
      inputRef.focus();

      // 查找第一个下划线的位置
      const text = fileItem.newFileName || "";
      const underscoreIndex = text.indexOf('_');

      if (underscoreIndex !== -1) {
        // 找到下划线，将光标定位到下划线前面
        inputRef.setSelectionRange(underscoreIndex, underscoreIndex);
      } else {
        // 没有下划线，选中所有文本
        inputRef.select();
      }
    }
  });
};

// 确认编辑文件名
const confirmEditFileName = (fileItem: FileParseItem) => {
  fileItem.isEditing = false;
  if (fileItem.newFileName) {
    fileItem.fileName = fileItem.newFileName;
    ElMessage.success("文件名已更新");
  }
};

// 取消编辑文件名
const cancelEditFileName = (fileItem: FileParseItem) => {
  fileItem.isEditing = false;
  // 恢复原来的文件名
  fileItem.newFileName = fileItem.fileName;
};

// 预览 PDF 文件
const previewPdf = (fileItem: FileParseItem) => {
  // 创建 Blob URL
  const blob = new Blob([fileItem.file], { type: "application/pdf" });
  const blobUrl = URL.createObjectURL(blob);

  // 添加参数来控制 PDF 显示（隐藏缩略图等）
  // #toolbar=0 只显示内容，不显示工具栏
  // #navpanes=0 隐藏左侧缩略图导航窗格
  previewFileUrl.value = `${blobUrl}#toolbar=0&navpanes=0&scrollbar=0`;

  currentPreviewFileName.value = fileItem.originalFileName;
  previewDialogVisible.value = true;
};

// 关闭预览对话框
const closePreviewDialog = () => {
  previewDialogVisible.value = false;
  // 释放 Blob URL
  if (previewFileUrl.value) {
    URL.revokeObjectURL(previewFileUrl.value);
    previewFileUrl.value = "";
  }
};

// 监听 ESC 键关闭预览
const handleKeydown = (event: KeyboardEvent) => {
  if (event.key === 'Escape' && previewDialogVisible.value) {
    closePreviewDialog();
  }
};

// 组件挂载时添加键盘监听
onMounted(() => {
  document.addEventListener('keydown', handleKeydown);
});

// 组件卸载时移除键盘监听
onUnmounted(() => {
  document.removeEventListener('keydown', handleKeydown);
});
</script>

<template>
  <div class="pdf-parser-container p-6">
    <el-card>
      <template #header>
        <div class="flex justify-between items-center">
          <h2 class="text-xl font-bold">铁路电子客票PDF名字修改</h2>
          <div class="flex items-center space-x-4">
            <div class="space-x-2">
              <el-button
                type="primary"
                :icon="Upload"
                @click="parseAllFiles"
                :loading="isParsing"
                :disabled="fileList.length === 0"
              >
                {{ isParsing ? "解析中..." : "开始解析" }}
              </el-button>
              <el-button
                type="success"
                :icon="Download"
                @click="exportAllRenamedFiles"
                :disabled="fileList.filter(item => item.status === 'success' && item.newFileName).length === 0"
              >
                批量导出
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
        </div>
      </template>

      <!-- 文件上传区域 -->
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
              <div class="upload-content text-center py-4">
                <el-icon class="el-icon--upload text-2xl mb-2">
                  <Upload />
                </el-icon>
                <div class="el-upload__text text-base">
                  将PDF文件或ZIP文件拖拽到此处，或<em>点击选择文件</em>
                </div>
                <div class="el-upload__tip text-xs text-gray-500 mt-1">
                  支持上传PDF和ZIP文件（支持嵌套ZIP），文件大小不超过50MB
                </div>
              </div>
            </el-upload>
          </div>

          <!-- 文件列表 -->
          <div class="file-list-section" v-if="fileList.length > 0">
            <h3 class="text-lg font-semibold mb-4">
              文件列表 ({{ fileList.length }})
            </h3>

            <el-alert
              title="提示"
              type="info"
              :closable="false"
              class="mb-4"
            >
              <template #default>
                <div>解析后的内容将打印在浏览器控制台中，请按 F12 打开开发者工具查看</div>
              </template>
            </el-alert>

            <!-- 总体解析进度条 -->
            <div v-if="isParsing || overallProgress > 0" class="mb-4">
              <div class="flex justify-between items-center mb-2">
                <span class="text-sm font-medium">
                  总体解析进度 ({{ successCount + errorCount }}/{{ fileList.length }}, 成功 {{ successCount }}, 失败 {{ errorCount }})
                </span>
                <span class="text-sm text-gray-600">{{ overallProgress }}%</span>
              </div>
              <el-progress
                :percentage="overallProgress"
                :status="overallProgress === 100 ? 'success' : undefined"
                :stroke-width="20"
              />
            </div>

            <!-- 表格展示文件列表 -->
            <el-table
              :data="fileList"
              stripe
              border
              style="width: 100%"
              :height="tableHeight"
              :max-height="tableHeight"
            >
              <el-table-column type="index" label="序号" width="60" align="center" />
              <el-table-column prop="originalFileName" label="原始文件名" width="350">
                <template #default="{ row }">
                  <span
                    class="text-blue-600 cursor-pointer hover:underline"
                    @click="previewPdf(row)"
                    title="点击预览 PDF"
                  >
                    {{ row.originalFileName }}
                  </span>
                </template>
              </el-table-column>
              <el-table-column label="新文件名" width="350">
                <template #default="{ row }">
                  <div v-if="row.newFileName">
                    <div v-if="row.isEditing" class="flex items-center space-x-2">
                      <el-input
                        :ref="(el: any) => editInputRefs.set(row.id, el)"
                        v-model="row.newFileName"
                        size="small"
                        placeholder="请输入新文件名"
                        @blur="confirmEditFileName(row)"
                        @keyup.enter="confirmEditFileName(row)"
                        @keyup.esc="cancelEditFileName(row)"
                      />
                    </div>
                    <div
                      v-else
                      :class="row.status === 'error' ? 'text-red-600' : (row.newFileName === row.originalFileName ? 'text-gray-600' : 'text-green-600')"
                      class="font-medium cursor-pointer hover:underline"
                      @click="startEditFileName(row)"
                      title="点击编辑"
                    >
                      {{ row.newFileName }}
                    </div>
                  </div>
                  <span v-else class="text-gray-400">-</span>
                </template>
              </el-table-column>
              <el-table-column label="状态" width="100">
                <template #default="{ row }">
                  <el-tag :type="getStatusTagType(row.status)" size="small">
                    {{ getStatusText(row.status) }}
                  </el-tag>
                </template>
              </el-table-column>
              <el-table-column label="进度" width="100">
                <template #default="{ row }">
                  <el-progress
                    v-if="row.status === 'parsing'"
                    :percentage="row.progress"
                    :stroke-width="6"
                    size="small"
                  />
                  <span v-else-if="row.status === 'success'">100%</span>
                  <span v-else>-</span>
                </template>
              </el-table-column>
              <el-table-column label="操作" width="250" fixed="right">
                <template #default="{ row }">
                  <div class="flex items-center space-x-2">
                    <el-button
                      v-if="row.status === 'success' && row.newFileName"
                      type="success"
                      size="small"
                      :icon="Download"
                      @click="downloadRenamedFile(row)"
                    >
                      下载重命名
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

    <!-- PDF 预览侧边栏 -->
    <Transition name="slide-fade">
      <div v-if="previewDialogVisible" class="pdf-preview-sidebar">
        <div class="pdf-preview-header">
          <span class="pdf-preview-title">{{ currentPreviewFileName }}</span>
          <el-button
            type="danger"
            :icon="Delete"
            circle
            size="small"
            @click="closePreviewDialog"
          />
        </div>
        <div class="pdf-preview-container">
          <iframe
            v-if="previewFileUrl"
            :src="previewFileUrl"
            frameborder="0"
          ></iframe>
        </div>
      </div>
    </Transition>
  </div>
</template>

<style>
/* 侧边栏过渡动画 */
.slide-fade-enter-active {
  transition: all 0.3s ease-out;
}

.slide-fade-leave-active {
  transition: all 0.3s ease-in;
}

.slide-fade-enter-from {
  transform: translateX(100%);
}

.slide-fade-leave-to {
  transform: translateX(100%);
}

.slide-fade-enter-to,
.slide-fade-leave-from {
  transform: translateX(0);
}
</style>

<style scoped>
.pdf-parser-container {
  max-width: 1400px;
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
  padding: 1rem;
}

.space-x-2 > * + * {
  margin-left: 0.5rem;
}

.space-y-2 > * + * {
  margin-top: 0.5rem;
}

.space-y-3 > * + * {
  margin-top: 0.75rem;
}

.page-content pre {
  margin: 0;
  font-family: 'Courier New', Courier, monospace;
  font-size: 13px;
  line-height: 1.5;
}

/* PDF 预览侧边栏 */
.pdf-preview-sidebar {
  position: fixed;
  top: 0;
  right: 0;
  width: 40%;
  height: 100vh;
  background: white;
  box-shadow: -2px 0 8px rgba(0, 0, 0, 0.15);
  z-index: 1000;
  display: flex;
  flex-direction: column;
}

.pdf-preview-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 16px 20px;
  border-bottom: 1px solid #e4e7ed;
  background: #f5f7fa;
}

.pdf-preview-title {
  font-size: 16px;
  font-weight: 600;
  color: #303133;
  flex: 1;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.pdf-preview-container {
  flex: 1;
  overflow: hidden;
  background: #525659;
}

.pdf-preview-container iframe {
  width: 100%;
  height: 100%;
  border: none;
}
</style>
