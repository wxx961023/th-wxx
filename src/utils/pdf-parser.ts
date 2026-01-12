/**
 * PDF解析工具模块 - 使用 pdfjs-dist
 * 提供PDF文本提取功能
 *
 * 参考了 src/views/welcome/pdf.vue 的实现方式
 * pdfjs-dist 是 Mozilla 官方维护的 PDF 解析库
 * 稳定可靠,在浏览器环境中运行良好
 */

import * as pdfjsLib from "pdfjs-dist";
import JSZip from "jszip";

// Promise.withResolvers polyfill for Win7 compatibility
if (typeof Promise !== "undefined" && !Promise.withResolvers) {
  console.log("[pdf-parser] 添加 Promise.withResolvers polyfill");
  (Promise as any).withResolvers = function <T>() {
    let resolve: (value: T | PromiseLike<T>) => void;
    let reject: (reason?: any) => void;
    const promise = new Promise<T>((res, rej) => {
      resolve = res;
      reject = rej;
    });
    return { promise, resolve, reject };
  };
}

// 检查是否支持 Promise.withResolvers
const supportsPromiseWithResolvers =
  typeof Promise !== "undefined" && typeof Promise.withResolvers === "function";

// Win7 或不支持 Promise.withResolvers 时禁用 worker
// PDF.js worker 在不支持 Promise.withResolvers 的环境下会报错
const shouldDisableWorker =
  !supportsPromiseWithResolvers ||
  navigator.userAgent.includes("Windows NT 6.1") ||
  navigator.userAgent.includes("Windows 7");

// 导出：是否禁用了 worker
export const isWorkerDisabled = shouldDisableWorker;

if (shouldDisableWorker) {
  console.warn("[pdf-parser] 禁用 PDF.js worker 以兼容 Win7 环境");
  // 不设置 workerSrc，让 PDF.js 在主线程运行
  // 设置为 undefined 而不是空字符串，避免 "No workerSrc specified" 错误
  delete (pdfjsLib.GlobalWorkerOptions as any).workerSrc;
} else {
  // 设置PDF.js的worker路径，使用本地worker文件
  pdfjsLib.GlobalWorkerOptions.workerSrc = "/pdf.worker.min.mjs";
}

// 导出 PDF.js 配置选项，用于支持中文字体
// 这对于正确提取中文、日文、韩文等 CJK 字符非常重要
export const pdfjsDocumentOptions = {
  // 尝试使用本地 cMap 文件
  // 如果 public/cmaps/ 目录不存在，中文可能无法正确提取
  // 可以运行: node copy-cmaps.js 来复制 cMap 文件
  cMapUrl: "/cmaps/",
  cMapPacked: true,
  // 使用系统字体，避免字体加载问题
  useSystemFonts: true,
  // 禁用自动 fetch 以提高性能
  disableAutoFetch: true,
  // 禁用流式处理
  disableStream: true,
  // 禁用 eval（安全考虑）
  isEvalSupported: false,
  // Win7 兼容：禁用 worker fetch
  useWorkerFetch: !shouldDisableWorker
};

// 检测 cMap 是否可用的函数
export async function checkCMapAvailability(): Promise<boolean> {
  try {
    const response = await fetch("/cmaps/Adobe-GB1-UCS2.bcmap", { method: "HEAD" });
    return response.ok;
  } catch {
    return false;
  }
}

/**
 * PDF解析结果接口
 */
export interface ParsedPDFContent {
  fileName: string;
  totalPages: number;
  pages: PageContent[];
  fullText: string;
}

/**
 * 页面内容接口
 */
export interface PageContent {
  pageNumber: number;
  text: string;
  textLength: number;
}

/**
 * PDF解析选项接口
 */
export interface PDFParserOptions {
  includeSeparator?: boolean; // 是否在页面间添加分隔符
  maxPages?: number; // 最大处理页数,undefined表示处理全部
  onProgress?: (current: number, total: number) => void; // 进度回调
  debugMode?: boolean; // 调试模式：输出所有原始文本项
}

/**
 * 从File对象解析PDF内容
 * 参考了 src/views/welcome/pdf.vue 的 processPdfFile 函数实现
 *
 * @param file - PDF文件对象
 * @param options - 解析选项
 * @returns Promise<ParsedPDFContent> - 解析后的PDF内容
 */
export async function parsePDFFile(
  file: File,
  options: PDFParserOptions = {}
): Promise<ParsedPDFContent> {
  const {
    includeSeparator = true,
    maxPages,
    onProgress,
    debugMode = false
  } = options;

  console.log("=== 开始解析PDF文件 (使用 pdfjs-dist) ===");
  console.log("文件名:", file.name);
  console.log("文件大小:", `${(file.size / 1024 / 1024).toFixed(2)}MB`);
  console.log("文件类型:", file.type);

  try {
    // 1. 读取文件为ArrayBuffer
    console.log("\n[步骤 1/3] 读取文件到 ArrayBuffer...");
    const readStartTime = Date.now();
    const arrayBuffer = await file.arrayBuffer();
    const readEndTime = Date.now();
    console.log(`✓ 文件读取完成，耗时: ${readEndTime - readStartTime}ms`);
    console.log(`ArrayBuffer 大小: ${arrayBuffer.byteLength} bytes`);

    // 2. 加载PDF文档
    console.log("\n[步骤 2/3] 加载PDF文档...");
    const loadStartTime = Date.now();

    // 使用PDF.js解析PDF内容 - 配置了 cMap 支持中文字体
    const loadingTask = pdfjsLib.getDocument({
      data: arrayBuffer,
      ...pdfjsDocumentOptions
    });

    const pdf = await loadingTask.promise;
    const loadEndTime = Date.now();
    console.log(`✓ PDF文档加载完成，耗时: ${loadEndTime - loadStartTime}ms`);
    console.log(`PDF信息:`, {
      numPages: pdf.numPages,
      fingerprint: pdf.fingerprints || "N/A"
    });

    // 3. 提取所有页面的文本内容
    console.log("\n[步骤 3/3] 提取文本内容...");
    const extractStartTime = Date.now();

    const pages: PageContent[] = [];
    const separator = includeSeparator ? "\n\n--- 页面分隔 ---\n\n" : "\n\n";
    let fullText = "";

    // 确定实际要处理的页数
    const totalPages = pdf.numPages;
    const pagesToProcess = maxPages
      ? Math.min(maxPages, totalPages)
      : totalPages;

    console.log(`计划处理页数: ${pagesToProcess}/${totalPages}`);

    // 提取所有页面的文本 - 改进的文本提取逻辑
    for (let i = 1; i <= pagesToProcess; i++) {
      console.log(`处理第 ${i}/${totalPages} 页...`);
      try {
        const page = await pdf.getPage(i);

        // 使用更详细的选项获取文本内容
        const textContent = await page.getTextContent();

        // 调试模式：输出所有原始文本项
        if (debugMode) {
          console.log(`  [调试] 第 ${i} 页原始文本项 (共 ${textContent.items.length} 项):`);
          textContent.items.forEach((item: any, index: number) => {
            console.log(`    [${index}] str="${item.str}" width=${item.width} transform=`, item.transform);
          });
        }

        // 改进的文本提取：保留更多布局信息
        let lastY = -1;
        const textItems: string[] = [];

        textContent.items.forEach((item: any) => {
          const transform = item.transform;
          const y = transform ? transform[5] : 0; // Y坐标

          // 如果Y坐标变化明显，说明换行了
          if (lastY !== -1 && Math.abs(y - lastY) > 5) {
            textItems.push("\n"); // 添加换行
          }

          // 添加文本内容（不过滤任何内容）
          if (item.str !== undefined) {
            textItems.push(item.str);
          }

          // 检查是否有明显的水平间距（可能是单词间隔）
          if (item.width && item.width > 10) {
            // textItems.push(" "); // 可选：添加空格
          }

          lastY = y;
        });

        // 组合所有文本项
        const pageText = textItems.join("");

        fullText += (i > 1 ? separator : "") + pageText;

        pages.push({
          pageNumber: i,
          text: pageText,
          textLength: pageText.length
        });

        console.log(`  ✓ 第 ${i} 页处理完成，文本长度: ${pageText.length} 字符`);
        console.log(`  - 文本项数量: ${textContent.items.length}`);

        // 触发进度回调
        if (onProgress) {
          onProgress(i, totalPages);
        }
      } catch (pageError) {
        console.error(`  ✗ 第 ${i} 页处理失败:`, pageError);
        // 继续处理其他页面
      }
    }

    const extractEndTime = Date.now();
    console.log(`\n✓ 文本提取完成，总耗时: ${extractEndTime - extractStartTime}ms`);
    console.log(`提取的文本总长度: ${fullText.length}`);
    console.log(`文本预览:`, fullText.substring(0, 200) + (fullText.length > 200 ? "..." : ""));

    // 4. 返回解析结果
    const result: ParsedPDFContent = {
      fileName: file.name,
      totalPages,
      pages,
      fullText
    };

    console.log("=== PDF解析完成 ===\n");

    return result;
  } catch (error) {
    console.error("\n✗ PDF解析失败:", error);
    console.error("错误详情:", {
      name: error.name,
      message: error.message,
      stack: error.stack
    });
    throw error;
  }
}

/**
 * 从ArrayBuffer解析PDF内容
 *
 * @param arrayBuffer - PDF文件的ArrayBuffer
 * @param fileName - 文件名(用于标识)
 * @param options - 解析选项
 * @returns Promise<ParsedPDFContent> - 解析后的PDF内容
 */
export async function parsePDFArrayBuffer(
  arrayBuffer: ArrayBuffer,
  fileName: string = "unknown.pdf",
  options: PDFParserOptions = {}
): Promise<ParsedPDFContent> {
  const {
    includeSeparator = true,
    maxPages,
    onProgress,
    debugMode = false
  } = options;

  console.log("=== 开始解析PDF ArrayBuffer (使用 pdfjs-dist) ===");
  console.log("文件名:", fileName);
  console.log("ArrayBuffer 大小:", `${arrayBuffer.byteLength} bytes`);

  try {
    // 加载PDF文档
    console.log("\n[步骤 1/2] 加载PDF文档...");
    const loadStartTime = Date.now();

    const loadingTask = pdfjsLib.getDocument({
      data: arrayBuffer,
      useWorkerFetch: false,
      isEvalSupported: false,
      useSystemFonts: true,
      disableAutoFetch: true,
      disableStream: true
    });

    const pdf = await loadingTask.promise;
    const loadEndTime = Date.now();
    console.log(`✓ PDF文档加载完成，耗时: ${loadEndTime - loadStartTime}ms`);
    console.log(`总页数: ${pdf.numPages}`);

    // 提取所有页面的文本内容
    console.log("\n[步骤 2/2] 提取文本内容...");
    const extractStartTime = Date.now();

    const pages: PageContent[] = [];
    const separator = includeSeparator ? "\n\n--- 页面分隔 ---\n\n" : "\n\n";
    let fullText = "";

    // 确定实际要处理的页数
    const totalPages = pdf.numPages;
    const pagesToProcess = maxPages
      ? Math.min(maxPages, totalPages)
      : totalPages;

    console.log(`计划处理页数: ${pagesToProcess}/${totalPages}`);

    for (let i = 1; i <= pagesToProcess; i++) {
      try {
        const page = await pdf.getPage(i);
        const textContent = await page.getTextContent();

        // 调试模式：输出所有原始文本项
        if (debugMode) {
          console.log(`  [调试] 第 ${i} 页原始文本项 (共 ${textContent.items.length} 项):`);
          textContent.items.forEach((item: any, index: number) => {
            console.log(`    [${index}] str="${item.str}" width=${item.width} transform=`, item.transform);
          });
        }

        // 改进的文本提取：保留更多布局信息
        let lastY = -1;
        const textItems: string[] = [];

        textContent.items.forEach((item: any) => {
          const transform = item.transform;
          const y = transform ? transform[5] : 0; // Y坐标

          // 如果Y坐标变化明显，说明换行了
          if (lastY !== -1 && Math.abs(y - lastY) > 5) {
            textItems.push("\n"); // 添加换行
          }

          // 添加文本内容（不过滤任何内容）
          if (item.str !== undefined) {
            textItems.push(item.str);
          }

          lastY = y;
        });

        // 组合所有文本项
        const pageText = textItems.join("");

        fullText += (i > 1 ? separator : "") + pageText;

        pages.push({
          pageNumber: i,
          text: pageText,
          textLength: pageText.length
        });

        console.log(`✓ 第 ${i} 页处理完成，文本长度: ${pageText.length} 字符`);

        if (onProgress) {
          onProgress(i, totalPages);
        }
      } catch (pageError) {
        console.error(`✗ 第 ${i} 页处理失败:`, pageError);
      }
    }

    const extractEndTime = Date.now();
    console.log(`\n✓ 文本提取完成，总耗时: ${extractEndTime - extractStartTime}ms`);
    console.log(`总文本长度: ${fullText.length} 字符`);

    // 返回解析结果
    const result: ParsedPDFContent = {
      fileName,
      totalPages,
      pages,
      fullText
    };

    console.log("=== PDF解析完成 ===\n");

    return result;
  } catch (error) {
    console.error("\n✗ PDF解析失败:", error);
    throw error;
  }
}

/**
 * 在控制台打印PDF解析结果
 *
 * @param result - PDF解析结果
 * @param options - 打印选项
 */
export function printPDFContent(
  result: ParsedPDFContent,
  options: {
    printFullText?: boolean; // 是否打印完整文本
    printPageText?: boolean; // 是否打印每页的文本
    maxPreviewLength?: number; // 预览文本的最大长度
  } = {}
): void {
  const {
    printFullText = false,
    printPageText = true,
    maxPreviewLength = 500
  } = options;

  console.log("\n" + "=".repeat(80));
  console.log("PDF解析结果 (使用 pdfjs-dist)");
  console.log("=".repeat(80));
  console.log(`文件名: ${result.fileName}`);
  console.log(`总页数: ${result.totalPages}`);
  console.log(`成功提取页数: ${result.pages.length}`);
  console.log(`总文本长度: ${result.fullText.length} 字符`);
  console.log("-".repeat(80));

  // 打印每页信息
  if (result.pages.length > 0) {
    console.log("\n页面信息:");
    result.pages.forEach((page) => {
      console.log(`\n[页面 ${page.pageNumber}]`);
      console.log(`  文本长度: ${page.textLength} 字符`);

      if (printPageText && page.text) {
        const previewText = page.text.length > maxPreviewLength
          ? page.text.substring(0, maxPreviewLength) + "... (截断)"
          : page.text;
        console.log(`  文本内容:\n${previewText}`);
      }
    });
  }

  // 打印完整文本
  if (printFullText) {
    console.log("\n" + "-".repeat(80));
    console.log("完整文本内容:");
    console.log("-".repeat(80));
    console.log(result.fullText);
  }

  console.log("\n" + "=".repeat(80));
  console.log("打印完成");
  console.log("=".repeat(80) + "\n");
}

/**
 * 从PDF文本中提取姓名
 * 匹配模式：身份证号后跟姓名，如 "4202221988****5775 肖烨" 或 "2114031985****843X 王森"
 *
 * @param text - PDF文本内容
 * @returns 提取到的姓名，未找到返回null
 */
export function extractName(text: string): string | null {
  // 先清洗文本：去掉所有空格和换行符
  const cleanedText = text.replace(/[\s\n\r\t]+/g, '');

  // 匹配模式：18位身份证 + 4-6个星号 + 3-4位字符（可包含数字和X） + 中文姓名（2-4个字符）
  // 姓名后可以是：
  // 1. 数字、下划线、英文字母
  // 2. "电子客票号"（完整的电子客票号文本）
  // 例如：4202221988****5775肖烨电子客票号 或 2114031985****843X王森26119110010000193505
  const namePattern = /\d{17}[\dXx]\*{4,6}[\dXx]{3,4}([\u4e00-\u9fa5]{2,4})(?=[0-9_a-zA-Z]|电子客票号|$)/;

  // 先尝试精确匹配
  let match = cleanedText.match(namePattern);

  // 如果精确匹配失败，尝试更宽松的模式
  if (!match) {
    // 模式：数字开头，包含星号，最后是中文姓名，姓名后必须是数字、下划线、字母、"电子客票号"或结尾
    const loosePattern = /\d+\*{4,}[\dXx]+([\u4e00-\u9fa5]{2,4})(?=[0-9_a-zA-Z]|电子客票号|$)/;
    match = cleanedText.match(loosePattern);
  }

  if (match && match[1]) {
    console.log(`✓ 提取到姓名: ${match[1]}`);
    return match[1];
  }

  console.log("✗ 未找到姓名信息");
  return null;
}

/**
 * 从PDF文本中提取电子客票号
 * 匹配模式：电子客票号:6580074086121798365302025
 *
 * @param text - PDF文本内容
 * @returns 提取到的票号，未找到返回null
 */
export function extractTicketNumber(text: string): string | null {
  // 匹配模式：电子客票号 + 冒号（中英文） + 数字
  const ticketPattern = /电子客票号\s*[:：]\s*(\d+)/;
  let match = text.match(ticketPattern);

  // 如果精确匹配失败，尝试更宽松的模式
  if (!match) {
    const loosePattern = /客票号\s*[:：]\s*(\d+)/;
    match = text.match(loosePattern);
  }

  if (match && match[1]) {
    console.log(`✓ 提取到票号: ${match[1]}`);
    return match[1];
  }

  console.log("✗ 未找到票号信息");
  return null;
}

/**
 * 从文件名中提取后缀部分
 * 例如：18812330_26329166851000023784.pdf -> 26329166851000023784
 *
 * @param fileName - 原始文件名
 * @returns 提取到的后缀，未找到返回null
 */
export function extractFileSuffix(fileName: string): string | null {
  // 移除.pdf扩展名
  const nameWithoutExt = fileName.replace(/\.pdf$/i, "");

  // 尝试匹配模式：数字_数字.pdf 或 数字.pdf
  // 提取最后一个下划线后的数字部分
  const underscoreIndex = nameWithoutExt.lastIndexOf("_");
  if (underscoreIndex > 0) {
    const suffix = nameWithoutExt.substring(underscoreIndex + 1);
    if (/^\d+$/.test(suffix)) {
      console.log(`✓ 提取到文件后缀: ${suffix}`);
      return suffix;
    }
  }

  // 如果没有下划线，尝试直接使用文件名（去除数字前缀）
  const numericSuffixMatch = nameWithoutExt.match(/\d+/);
  if (numericSuffixMatch) {
    const suffix = numericSuffixMatch[0];
    console.log(`✓ 提取到文件后缀: ${suffix}`);
    return suffix;
  }

  console.log("✗ 未找到文件后缀");
  return null;
}

/**
 * 生成新的PDF文件名
 * 格式：姓名_后缀.pdf
 *
 * @param parsedContent - PDF解析结果
 * @param originalFileName - 原始文件名
 * @returns 新文件名，提取失败返回null
 */
export function generateNewFileName(
  parsedContent: ParsedPDFContent,
  originalFileName: string
): string | null {
  console.log(`\n开始生成新文件名...`);
  console.log(`原始文件名: ${originalFileName}`);

  // 提取姓名
  const name = extractName(parsedContent.fullText);
  if (!name) {
    console.log("生成失败：无法提取姓名");
    return null;
  }

  // 提取文件后缀
  const suffix = extractFileSuffix(originalFileName);
  if (!suffix) {
    console.log("生成失败：无法提取文件后缀");
    return null;
  }

  const newFileName = `${name}_${suffix}.pdf`;
  console.log(`✓ 生成新文件名: ${newFileName}`);

  return newFileName;
}

/**
 * 获取 PDF.js 版本信息
 * 用于调试
 */
export function getPDFJSInfo(): void {
  console.log("PDF.js 信息:", {
    version: pdfjsLib.version || "未知",
    workerSrc: pdfjsLib.GlobalWorkerOptions.workerSrc || "未设置"
  });
}

// ==================== ZIP 文件处理 ====================

/**
 * ZIP 中的文件项
 */
export interface ZipFileItem {
  path: string; // 文件在 ZIP 中的路径
  name: string; // 文件名
  file: File; // 文件对象
}

/**
 * 从 ZIP 文件中提取所有 PDF 文件（支持嵌套 ZIP）
 *
 * @param zipFile - ZIP 文件对象
 * @param options - 可选项
 * @returns Promise<ZipFileItem[]> 提取的 PDF 文件列表
 */
export async function extractPDFsFromZip(
  zipFile: File,
  options: {
    onProgress?: (current: number, total: number, message: string) => void;
    maxDepth?: number; // 最大嵌套深度，默认 3
  } = {}
): Promise<ZipFileItem[]> {
  const { onProgress, maxDepth = 3 } = options;
  const pdfFiles: ZipFileItem[] = [];
  let processedCount = 0;

  /**
   * 递归处理 ZIP 文件
   */
  async function processZip(file: File, basePath: string = "", depth: number = 0): Promise<void> {
    if (depth > maxDepth) {
      console.warn(`达到最大嵌套深度 ${maxDepth}，跳过: ${file.name}`);
      return;
    }

    try {
      const zip = new JSZip();
      const zipContent = await zip.loadAsync(file);

      const files = Object.keys(zipContent.files);

      for (let i = 0; i < files.length; i++) {
        const filePath = files[i];
        const zipEntry = zipContent.files[filePath];

        // 跳过目录
        if (zipEntry.dir) {
          continue;
        }

        // 通知进度
        if (onProgress) {
          onProgress(processedCount + 1, -1, `处理: ${filePath}`);
        }

        // 检查是否是 ZIP 文件（嵌套 ZIP）
        if (filePath.toLowerCase().endsWith('.zip')) {
          console.log(`发现嵌套 ZIP: ${filePath}`);

          // 提取嵌套的 ZIP 文件
          const zipBlob = await zipEntry.async('blob');
          const nestedZipFile = new File([zipBlob], zipEntry.name, {
            type: 'application/zip'
          });

          // 递归处理嵌套的 ZIP，保持完整的 ZIP 文件名作为路径的一部分
          // 这样导出时可以保持原有的 ZIP 文件结构
          await processZip(nestedZipFile, `${basePath}${filePath}/`, depth + 1);
          processedCount++;
          continue;
        }

        // 检查是否是 PDF 文件
        if (filePath.toLowerCase().endsWith('.pdf')) {
          console.log(`找到 PDF: ${filePath}`);

          // 提取 PDF 文件
          const pdfBlob = await zipEntry.async('blob');
          const pdfFile = new File([pdfBlob], zipEntry.name, {
            type: 'application/pdf'
          });

          pdfFiles.push({
            path: `${basePath}${filePath}`,
            name: zipEntry.name,
            file: pdfFile
          });

          console.log(`✓ 提取 PDF: ${zipEntry.name} (${basePath}${filePath})`);
          processedCount++;
        }
      }
    } catch (error) {
      console.error(`处理 ZIP 文件失败 (${file.name}):`, error);
      throw new Error(`处理 ZIP 文件失败: ${error.message}`);
    }
  }

  await processZip(zipFile);

  console.log(`\n从 ZIP 中提取了 ${pdfFiles.length} 个 PDF 文件`);
  return pdfFiles;
}

/**
 * 创建包含重命名后文件的 ZIP
 *
 * @param fileItems - 文件项列表（包含新文件名）
 * @param zipName - 生成的 ZIP 文件名
 * @returns Promise<Blob> ZIP 文件的 Blob 对象
 */
export async function createRenamedZip(
  fileItems: Array<{
    file: File;
    newFileName: string;
  }>,
  zipName: string = "renamed_files.zip"
): Promise<Blob> {
  console.log(`\n开始创建 ZIP 文件: ${zipName}`);
  console.log(`文件数量: ${fileItems.length}`);

  const zip = new JSZip();

  // 添加所有文件到 ZIP
  for (let i = 0; i < fileItems.length; i++) {
    const item = fileItems[i];
    if (item.newFileName) {
      console.log(`添加文件到 ZIP: ${item.newFileName}`);
      zip.file(item.newFileName, item.file);
    }
  }

  // 生成 ZIP 文件
  console.log("正在生成 ZIP 文件...");
  const zipBlob = await zip.generateAsync({
    type: "blob",
    compression: "DEFLATE",
    compressionOptions: {
      level: 6
    }
  });

  console.log(`✓ ZIP 文件创建完成: ${(zipBlob.size / 1024 / 1024).toFixed(2)} MB`);

  return zipBlob;
}
