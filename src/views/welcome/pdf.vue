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
  Warning,
  Edit,
  Check,
  Close,
  Refresh,
  View
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

// 分组文件项
interface GroupFileItem {
  id: string;
  originalName: string;
  companyName: string;
  file: File;
}

// 公司分组信息
interface CompanyGroup {
  companyName: string;
  files: GroupFileItem[];
  folderName: string;
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

// 分组功能相关的响应式数据
const groupFileList = ref<GroupFileItem[]>([]);
const isGroupExtracting = ref(false);
const groupZipFile = ref<File | null>(null);
const companyGroups = ref<CompanyGroup[]>([]);

// 文件夹名称编辑状态管理
const editingFolderNames = ref<Map<string, boolean>>(new Map());
const tempFolderNames = ref<Map<string, string>>(new Map());

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
  { keywords: ["代订机票款"], type: "机票费" },
  { keywords: ["代订机票费"], type: "机票费" },
  { keywords: ["代订服务费"], type: "服务费" },
  { keywords: ["代订接车费"], type: "接车费" },
  { keywords: ["代订签证费"], type: "签证费" },
  { keywords: ["代订住宿费"], type: "住宿费" },
  { keywords: ["代订酒店费"], type: "酒店费" },
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

// 从文件名提取公司名称（用于分组功能）
const extractCompanyNameForGroup = (filename: string): string => {
  const parts = filename.split("_");
  if (parts.length >= 1) {
    return parts[0]; // 第一个下划线之前的部分
  }
  return "";
};

// 标准化公司名称（处理特殊字符）
const normalizeCompanyName = (companyName: string): string => {
  return companyName
    .replace(/（/g, "(")
    .replace(/）/g, ")")
    .replace(/【/g, "[")
    .replace(/】/g, "]")
    .trim();
};

// 清理文件夹名称中的非法字符
const sanitizeFolderName = (folderName: string): string => {
  return folderName.replace(/[<>:"/\\|?*]/g, "_");
};

// 格式化金额显示
const formatAmount = (amount: string): string => {
  if (!amount || amount === "0") return "0";

  // 如果金额包含小数点
  if (amount.includes(".")) {
    const parts = amount.split(".");
    const integerPart = parts[0];
    const decimalPart = parts[1] || "";

    // 检查小数部分是否全为0
    if (decimalPart && Number(decimalPart) > 0) {
      // 有有效小数位，保留所有小数位
      return amount;
    } else {
      // 小数部分全为0，只返回整数部分
      return integerPart;
    }
  }

  // 没有小数点，直接返回
  return amount;
};

// 获取上一个月份的中文显示
const getPreviousMonthText = (): string => {
  const now = new Date();
  const currentMonth = now.getMonth(); // 0-11，0表示1月，11表示12月

  // 计算上一个月的月份数字（1-12）
  let previousMonthNumber: number;
  if (currentMonth === 0) {
    // 如果当前是1月（getMonth()=0），上一个月是12月
    previousMonthNumber = 12;
  } else {
    // 其他情况，上一个月就是当前月份数字（getMonth()+1-1 = getMonth()）
    // 例如：当前9月(getMonth()=8)，上一个月是8月
    previousMonthNumber = currentMonth;
  }

  return `${previousMonthNumber}月`;
};

// 从PDF内容确定费用类型
const determineExpenseType = (pdfText: string): string => {
  console.log("从PDF内容确定费用类型: ", pdfText);

  // 清理文本：去除多余空格、换行符等
  const cleanText = pdfText
    .replace(/\s+/g, " ") // 将多个空格替换为单个空格
    .replace(/\n/g, " ") // 将换行符替换为空格
    .trim();

  // 创建无空格版本用于精确匹配
  const noSpaceText = cleanText.replace(/\s/g, "");

  console.log("清理后的文本:", cleanText);
  console.log("无空格文本:", noSpaceText);

  for (const rule of expenseTypeRules) {
    // 策略1: 直接匹配（原有逻辑）
    if (rule.keywords.some(keyword => cleanText.includes(keyword))) {
      console.log(
        `匹配成功 - 直接匹配: ${rule.type}, 关键词: ${rule.keywords.find(k => cleanText.includes(k))}`
      );
      return rule.type;
    }

    // 策略2: 无空格匹配（处理空格分割问题）
    if (
      rule.keywords.some(keyword => {
        const noSpaceKeyword = keyword.replace(/\s/g, "");
        return noSpaceText.includes(noSpaceKeyword);
      })
    ) {
      const matchedKeyword = rule.keywords.find(keyword => {
        const noSpaceKeyword = keyword.replace(/\s/g, "");
        return noSpaceText.includes(noSpaceKeyword);
      });
      console.log(
        `匹配成功 - 无空格匹配: ${rule.type}, 关键词: ${matchedKeyword}`
      );
      return rule.type;
    }

    // 策略3: 模糊匹配（允许关键词字符间有空格）
    if (
      rule.keywords.some(keyword => {
        // 将关键词转换为正则表达式，允许字符间有空格
        const fuzzyPattern = keyword.split("").join("\\s*");
        const regex = new RegExp(fuzzyPattern, "i");
        return regex.test(cleanText);
      })
    ) {
      const matchedKeyword = rule.keywords.find(keyword => {
        const fuzzyPattern = keyword.split("").join("\\s*");
        const regex = new RegExp(fuzzyPattern, "i");
        return regex.test(cleanText);
      });
      console.log(
        `匹配成功 - 模糊匹配: ${rule.type}, 关键词: ${matchedKeyword}`
      );
      return rule.type;
    }
  }

  console.log("未找到匹配的费用类型，返回默认值");
  return "未命名"; // 默认类型
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

    // 使用格式化函数处理金额显示
    const formattedAmount = formatAmount(amount);
    fileItem.amount = formattedAmount;
    fileItem.progress = 90;

    // 生成新文件名，使用格式化后的金额
    const newFileName = sanitizeFileName(
      `${companyName}_${expenseType}${formattedAmount}.pdf`
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
        if (!fileItem.newName.includes(".pdf")) {
          fileItem.newName = fileItem.newName + ".pdf";
        }
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

// ========== 分组功能相关函数 ==========

// 处理分组ZIP文件解压
const extractGroupZipFile = async (
  zipFileData: File
): Promise<GroupFileItem[]> => {
  try {
    isGroupExtracting.value = true;
    const zip = new JSZip();
    const zipContent = await zip.loadAsync(zipFileData);
    const extractedFiles: GroupFileItem[] = [];

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
          const companyName = extractCompanyNameForGroup(filename);
          const fileItem: GroupFileItem = {
            id:
              Date.now().toString() +
              Math.random().toString(36).substring(2, 11),
            originalName: filename,
            companyName: normalizeCompanyName(companyName),
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
    isGroupExtracting.value = false;
  }
};

// 分组文件上传前的检查
const beforeGroupUpload: UploadProps["beforeUpload"] = file => {
  const isZip =
    file.type === "application/zip" ||
    file.type === "application/x-zip-compressed" ||
    file.name.toLowerCase().endsWith(".zip");
  const isLt50M = file.size / 1024 / 1024 < 50;

  if (!isZip) {
    ElMessage.error("只能上传ZIP压缩包!");
    return false;
  }
  if (!isLt50M) {
    ElMessage.error("文件大小不能超过50MB!");
    return false;
  }
  return true;
};

// 分组文件选择处理
const handleGroupFileChange: UploadProps["onChange"] = async uploadFile => {
  if (uploadFile.raw) {
    const file = uploadFile.raw;

    try {
      groupZipFile.value = file;
      ElMessage.info("正在解压ZIP文件，请稍候...");
      const extractedFiles = await extractGroupZipFile(file);

      if (extractedFiles.length === 0) {
        ElMessage.warning("ZIP文件中没有找到PDF文件");
        return;
      }

      groupFileList.value = extractedFiles;
      generateCompanyGroups();
      ElMessage.success(
        `成功从ZIP文件中提取了 ${extractedFiles.length} 个PDF文件`
      );
    } catch (error) {
      ElMessage.error(
        "ZIP文件处理失败: " +
          (error instanceof Error ? error.message : "未知错误")
      );
    }
  }
};

// 生成公司分组
const generateCompanyGroups = () => {
  const groupMap = new Map<string, GroupFileItem[]>();

  // 按公司名称分组
  groupFileList.value.forEach(file => {
    const companyName = file.companyName || "未知公司";
    if (!groupMap.has(companyName)) {
      groupMap.set(companyName, []);
    }
    groupMap.get(companyName)!.push(file);
  });
  const fileTime = new Date().toLocaleString();
  // 生成分组信息
  const previousMonth = getPreviousMonthText();
  companyGroups.value = Array.from(groupMap.entries()).map(
    ([companyName, files]) => ({
      companyName,
      files,
      folderName: sanitizeFolderName(`${companyName}${previousMonth}发票`)
    })
  );
};

// 按公司分组下载ZIP文件
const downloadGroupedZip = async () => {
  if (companyGroups.value.length === 0) {
    ElMessage.error("没有文件可以下载");
    return;
  }

  try {
    ElMessage.info("正在按公司分组打包文件，请稍候...");
    const zip = new JSZip();

    // 按公司分组处理文件
    for (const group of companyGroups.value) {
      if (group.files.length >= 2) {
        // 有2个或以上文件的公司，创建文件夹
        const folder = zip.folder(group.folderName);
        if (folder) {
          for (const file of group.files) {
            folder.file(file.originalName, file.file);
          }
        }
      } else if (group.files.length === 1) {
        // 只有1个文件的公司，直接放在根目录
        const file = group.files[0];
        zip.file(file.originalName, file.file);
      }
    }

    // 生成ZIP文件
    const zipBlob = await zip.generateAsync({ type: "blob" });

    // 下载文件
    const originalZipName = `发票${getPreviousMonthText()}.zip`;
    // const newZipName = originalZipName.replace(/\.zip$/i, "_grouped.zip");
    saveAs(zipBlob, originalZipName);

    ElMessage.success("按公司分组打包下载成功");
  } catch (error) {
    console.error("分组打包下载失败:", error);
    ElMessage.error("分组打包失败");
  }
};

// 清空分组文件
const clearGroupFiles = () => {
  groupFileList.value = [];
  companyGroups.value = [];
  groupZipFile.value = null;
  editingFolderNames.value.clear();
  tempFolderNames.value.clear();
};

// ========== 文件夹名称编辑功能 ==========

// 开始编辑文件夹名称
const startEditFolderName = (
  companyName: string,
  currentFolderName: string
) => {
  editingFolderNames.value.set(companyName, true);
  tempFolderNames.value.set(companyName, currentFolderName);
};

// 取消编辑文件夹名称
const cancelEditFolderName = (companyName: string) => {
  editingFolderNames.value.set(companyName, false);
  tempFolderNames.value.delete(companyName);
};

// 保存文件夹名称
const saveFolderName = (companyName: string) => {
  const newFolderName = tempFolderNames.value.get(companyName);
  if (!newFolderName || !newFolderName.trim()) {
    ElMessage.error("文件夹名称不能为空");
    return;
  }

  // 清理文件名中的非法字符
  const sanitizedName = sanitizeFolderName(newFolderName.trim());

  // 检查是否有重复的文件夹名称
  const isDuplicate = companyGroups.value.some(
    group =>
      group.companyName !== companyName && group.folderName === sanitizedName
  );

  if (isDuplicate) {
    ElMessage.error("文件夹名称重复，请使用其他名称");
    return;
  }

  // 更新对应公司分组的文件夹名称
  const groupIndex = companyGroups.value.findIndex(
    group => group.companyName === companyName
  );
  if (groupIndex !== -1) {
    companyGroups.value[groupIndex].folderName = sanitizedName;
    editingFolderNames.value.set(companyName, false);
    tempFolderNames.value.delete(companyName);
    ElMessage.success("文件夹名称已更新");
  }
};

// 重置为默认文件夹名称
const resetToDefaultFolderName = (companyName: string) => {
  const previousMonth = getPreviousMonthText();
  const defaultFolderName = sanitizeFolderName(
    `${companyName}${previousMonth}发票`
  );

  const groupIndex = companyGroups.value.findIndex(
    group => group.companyName === companyName
  );
  if (groupIndex !== -1) {
    companyGroups.value[groupIndex].folderName = defaultFolderName;
    ElMessage.success("已恢复为默认文件夹名称");
  }
};

// 检查是否正在编辑
const isEditingFolder = (companyName: string): boolean => {
  return editingFolderNames.value.get(companyName) || false;
};

// 获取临时文件夹名称
const getTempFolderName = (companyName: string): string => {
  return tempFolderNames.value.get(companyName) || "";
};

// ========== PDF预览功能 ==========

// 预览PDF文件
const previewPdfFile = (fileItem: FileItem) => {
  if (!fileItem.file) {
    ElMessage.error("文件不存在，无法预览");
    return;
  }

  try {
    // 创建临时URL
    const fileUrl = URL.createObjectURL(fileItem.file);

    // 在新窗口中打开PDF预览
    const previewWindow = window.open(
      fileUrl,
      "_blank",
      "width=1000,height=800,scrollbars=yes,resizable=yes"
    );

    if (!previewWindow) {
      ElMessage.error("无法打开预览窗口，请检查浏览器弹窗设置");
      URL.revokeObjectURL(fileUrl);
      return;
    }

    // 监听窗口关闭事件，释放内存资源
    const checkClosed = setInterval(() => {
      if (previewWindow.closed) {
        URL.revokeObjectURL(fileUrl);
        clearInterval(checkClosed);
      }
    }, 1000);

    // 设置超时释放资源（防止内存泄漏）
    setTimeout(() => {
      if (!previewWindow.closed) {
        URL.revokeObjectURL(fileUrl);
      }
      clearInterval(checkClosed);
    }, 300000); // 5分钟后自动释放
  } catch (error) {
    console.error("PDF预览失败:", error);
    ElMessage.error("PDF文件预览失败，文件可能已损坏");
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
          <div class="upload-content text-center py-4">
            <el-icon class="el-icon--upload text-2xl mb-2">
              <FolderOpened v-if="isExtracting" />
              <Upload v-else />
            </el-icon>
            <div class="el-upload__text text-base">
              <span v-if="isExtracting">正在解压ZIP文件...</span>
              <span v-else
                >将ZIP压缩包或PDF文件拖拽到此处，或<em>点击选择文件</em></span
              >
            </div>
            <div class="el-upload__tip text-xs text-gray-500 mt-1">
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
              <div class="flex items-center space-x-2">
                <el-button
                  type="text"
                  class="pdf-preview-link p-0 text-left"
                  @click="previewPdfFile(row)"
                  :title="`点击预览: ${row.originalName}`"
                >
                  <div class="flex items-center space-x-1">
                    <el-icon class="text-blue-500">
                      <View />
                    </el-icon>
                    <span
                      class="truncate text-blue-600 hover:text-blue-800 underline"
                    >
                      {{ row.originalName }}
                    </span>
                  </div>
                </el-button>
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
              <span v-if="row.amount">{{ formatAmount(row.amount) }}</span>
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
                  v-model="row.newName.split('.pdf')[0]"
                  placeholder="输入新文件名（不含扩展名）"
                  clearable
                  :class="{
                    'duplicate-filename-input': isFileNameDuplicate(row.newName)
                  }"
                  @input="
                    (value: string) => {
                      if (value && !value.endsWith('.pdf')) {
                        row.newName = sanitizeFileName(value);
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

    <!-- 分组功能区域 -->
    <el-card class="mb-6">
      <template #header>
        <div class="flex justify-between items-center">
          <h2 class="text-xl font-bold">PDF按公司分组打包工具</h2>
          <div class="space-x-2">
            <el-button
              type="success"
              :icon="Download"
              @click="downloadGroupedZip"
              :disabled="companyGroups.length === 0"
            >
              按公司分组下载
            </el-button>
            <el-button
              type="danger"
              :icon="Delete"
              @click="clearGroupFiles"
              :disabled="groupFileList.length === 0"
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
          :auto-upload="false"
          :before-upload="beforeGroupUpload"
          :on-change="handleGroupFileChange"
          :show-file-list="false"
          accept=".zip"
        >
          <div class="upload-content text-center py-4">
            <el-icon class="el-icon--upload text-2xl mb-2">
              <FolderOpened v-if="isGroupExtracting" />
              <Upload v-else />
            </el-icon>
            <div class="el-upload__text text-base">
              <span v-if="isGroupExtracting">正在解压ZIP文件...</span>
              <span v-else
                >将包含PDF文件的ZIP压缩包拖拽到此处，或<em
                  >点击选择文件</em
                ></span
              >
            </div>
            <div class="el-upload__tip text-xs text-gray-500 mt-1">
              支持上传ZIP压缩包，将按公司名称自动分组，文件大小不超过50MB
            </div>
          </div>
        </el-upload>
      </div>

      <!-- 公司分组列表 -->
      <div class="group-list-section" v-if="companyGroups.length > 0">
        <h3 class="text-lg font-semibold mb-4">
          公司分组列表 ({{ companyGroups.length }} 个公司，共
          {{ groupFileList.length }} 个文件)
        </h3>

        <div class="space-y-4">
          <el-card
            v-for="group in companyGroups"
            :key="group.companyName"
            class="group-card"
            shadow="hover"
          >
            <template #header>
              <div class="flex justify-between items-center">
                <div class="flex items-center space-x-2">
                  <el-tag
                    :type="group.files.length >= 2 ? 'success' : 'info'"
                    size="small"
                  >
                    {{ group.files.length >= 2 ? "创建文件夹" : "根目录" }}
                  </el-tag>
                  <span class="font-medium">{{ group.companyName }}</span>
                  <span class="text-gray-500"
                    >({{ group.files.length }} 个文件)</span
                  >
                </div>
                <div
                  v-if="group.files.length >= 2"
                  class="flex items-center space-x-2"
                >
                  <!-- 非编辑状态 -->
                  <div
                    v-if="!isEditingFolder(group.companyName)"
                    class="flex items-center space-x-2"
                  >
                    <span class="text-sm text-gray-600"
                      >文件夹：{{ group.folderName }}</span
                    >
                    <el-button
                      type="text"
                      size="small"
                      :icon="Edit"
                      @click="
                        startEditFolderName(group.companyName, group.folderName)
                      "
                      class="text-blue-500 hover:text-blue-700"
                      title="编辑文件夹名称"
                    />
                    <el-button
                      type="text"
                      size="small"
                      :icon="Refresh"
                      @click="resetToDefaultFolderName(group.companyName)"
                      class="text-green-500 hover:text-green-700"
                      title="恢复默认名称"
                    />
                  </div>

                  <!-- 编辑状态 -->
                  <div v-else class="flex items-center space-x-2">
                    <span class="text-sm text-gray-600">文件夹：</span>
                    <el-input
                      :model-value="getTempFolderName(group.companyName)"
                      @update:model-value="
                        value => tempFolderNames.set(group.companyName, value)
                      "
                      size="small"
                      class="w-48"
                      placeholder="输入文件夹名称"
                      @keyup.enter="saveFolderName(group.companyName)"
                      @keyup.esc="cancelEditFolderName(group.companyName)"
                    />
                    <el-button
                      type="text"
                      size="small"
                      :icon="Check"
                      @click="saveFolderName(group.companyName)"
                      class="text-green-500 hover:text-green-700"
                      title="保存"
                    />
                    <el-button
                      type="text"
                      size="small"
                      :icon="Close"
                      @click="cancelEditFolderName(group.companyName)"
                      class="text-red-500 hover:text-red-700"
                      title="取消"
                    />
                  </div>
                </div>
              </div>
            </template>

            <div class="file-list">
              <div
                v-for="file in group.files"
                :key="file.id"
                class="flex items-center justify-between py-2 border-b border-gray-100 last:border-b-0"
              >
                <div class="flex items-center space-x-2">
                  <el-icon class="text-red-500">
                    <DocumentCopy />
                  </el-icon>
                  <span class="text-sm">{{ file.originalName }}</span>
                </div>
              </div>
            </div>
          </el-card>
        </div>
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
  padding: 1rem;
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

/* 分组功能样式 */
.group-card {
  margin-bottom: 1rem;
}

.group-card :deep(.el-card__header) {
  padding: 12px 16px;
  background-color: #f8f9fa;
}

.group-card :deep(.el-card__body) {
  padding: 12px 16px;
}

.file-list {
  max-height: 200px;
  overflow-y: auto;
}

.space-y-4 > * + * {
  margin-top: 1rem;
}

/* 文件夹名称编辑功能样式 */
.group-card .el-button--text {
  padding: 2px 4px;
  min-height: auto;
}

.group-card .el-button--text:hover {
  background-color: transparent;
}

.group-card .el-input--small {
  font-size: 12px;
}

.group-card .el-input--small .el-input__wrapper {
  padding: 2px 8px;
}

/* PDF预览功能样式 */
.pdf-preview-link {
  max-width: 100%;
  justify-content: flex-start !important;
  text-align: left !important;
}

.pdf-preview-link:hover {
  background-color: transparent !important;
}

.pdf-preview-link .truncate {
  max-width: 160px;
  display: inline-block;
}

.pdf-preview-link:hover .truncate {
  color: #1d4ed8 !important;
}
</style>
