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

// 始终启用 Worker，将 PDF 解析放到独立线程，避免主线程阻塞导致页面卡死
pdfjsLib.GlobalWorkerOptions.workerSrc = "/pdf.worker.min.mjs";

// 导出 PDF.js 配置选项，用于支持中文字体
export const pdfjsDocumentOptions = {
  cMapUrl: "/cmaps/",
  cMapPacked: true,
  useSystemFonts: true,
  disableAutoFetch: true,
  disableStream: true,
  isEvalSupported: false,
  useWorkerFetch: true,
  useWorker: true
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

  let loadingTask: any = null;
  let pdf: any = null;

  try {
    // 1. 读取文件为ArrayBuffer
    const arrayBuffer = await file.arrayBuffer();

    // 2. 加载PDF文档（Worker 线程处理）
    loadingTask = pdfjsLib.getDocument({
      data: new Uint8Array(arrayBuffer),
      ...pdfjsDocumentOptions
    });

    pdf = await loadingTask.promise;

    // 3. 提取所有页面的文本内容
    const pages: PageContent[] = [];
    const separator = includeSeparator ? "\n\n--- 页面分隔 ---\n\n" : "\n\n";
    let fullText = "";

    const totalPages = pdf.numPages;
    const pagesToProcess = maxPages
      ? Math.min(maxPages, totalPages)
      : totalPages;

    for (let i = 1; i <= pagesToProcess; i++) {
      try {
        const page = await pdf.getPage(i);
        const textContent = await page.getTextContent();

        let lastY = -1;
        const textItems: string[] = [];

        textContent.items.forEach((item: any) => {
          const transform = item.transform;
          const y = transform ? transform[5] : 0;

          if (lastY !== -1 && Math.abs(y - lastY) > 5) {
            textItems.push("\n");
          }

          if (item.str !== undefined) {
            textItems.push(item.str);
          }

          lastY = y;
        });

        const pageText = textItems.join("");

        fullText += (i > 1 ? separator : "") + pageText;

        pages.push({
          pageNumber: i,
          text: pageText,
          textLength: pageText.length
        });

        // 清理页面资源
        page.cleanup();

        if (onProgress) {
          onProgress(i, totalPages);
        }
      } catch {
        // 继续处理其他页面
      }
    }

    const result: ParsedPDFContent = {
      fileName: file.name,
      totalPages,
      pages,
      fullText
    };

    return result;
  } catch (error) {
    throw error;
  } finally {
    // 清理 PDF 文档资源，释放内存
    if (pdf) {
      try {
        await pdf.destroy();
      } catch {
        // 忽略清理错误
      }
    }
    // 注意：loadingTask 无需手动 terminate，pdf.destroy() 会处理
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

  let pdf: any = null;

  try {
    // 加载PDF文档
    const loadingTask = pdfjsLib.getDocument({
      data: new Uint8Array(arrayBuffer),
      useWorkerFetch: false,
      isEvalSupported: false,
      useSystemFonts: true,
      disableAutoFetch: true,
      disableStream: true
    });

    pdf = await loadingTask.promise;

    // 提取所有页面的文本内容
    const pages: PageContent[] = [];
    const separator = includeSeparator ? "\n\n--- 页面分隔 ---\n\n" : "\n\n";
    let fullText = "";

    const totalPages = pdf.numPages;
    const pagesToProcess = maxPages
      ? Math.min(maxPages, totalPages)
      : totalPages;

    for (let i = 1; i <= pagesToProcess; i++) {
      try {
        const page = await pdf.getPage(i);
        const textContent = await page.getTextContent();

        let lastY = -1;
        const textItems: string[] = [];

        textContent.items.forEach((item: any) => {
          const transform = item.transform;
          const y = transform ? transform[5] : 0;

          if (lastY !== -1 && Math.abs(y - lastY) > 5) {
            textItems.push("\n");
          }

          if (item.str !== undefined) {
            textItems.push(item.str);
          }

          lastY = y;
        });

        const pageText = textItems.join("");

        fullText += (i > 1 ? separator : "") + pageText;

        pages.push({
          pageNumber: i,
          text: pageText,
          textLength: pageText.length
        });

        page.cleanup();

        if (onProgress) {
          onProgress(i, totalPages);
        }
      } catch {
        // 继续处理其他页面
      }
    }

    const result: ParsedPDFContent = {
      fileName,
      totalPages,
      pages,
      fullText
    };

    return result;
  } catch (error) {
    throw error;
  } finally {
    if (pdf) {
      try {
        await pdf.destroy();
      } catch {
        // 忽略清理错误
      }
    }
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
  // 该函数已不再输出到控制台，保留接口以兼容调用方
  void result;
  void options;
}

/**
 * 从PDF文本中提取姓名
 * 匹配模式：
 * 1. 身份证号后跟姓名，如 "4202221988****5775 肖烨" 或 "2114031985****843X 王森"
 * 2. 护照号后跟姓名，如 "H043388** 林茵楠"
 *
 * @param text - PDF文本内容
 * @returns 提取到的姓名，未找到返回null
 */
export function extractName(text: string): string | null {
  // 先清洗文本：去掉所有空格和换行符
  const cleanedText = text.replace(/[\s\n\r\t]+/g, '');

  // 模式1：18位身份证 + 4-6个星号 + 3-4位字符（可包含数字和X） + 中文姓名（2-4个字符）
  const idCardPattern = /\d{17}[\dXx]\*{4,6}[\dXx]{3,4}([\u4e00-\u9fa5]{2,4})(?=[0-9_a-zA-Z]|电子客票号|$)/;

  let match = cleanedText.match(idCardPattern);
  if (match && match[1]) {
    return match[1];
  }

  // 模式2：护照号格式 - 字母开头 + 数字 + 星号 + 中文姓名
  const passportPattern = /[A-Za-z]\d{5,}\*{2,}([\u4e00-\u9fa5]{2,4})(?=[0-9_a-zA-Z]|电子客票号|$)/;

  match = cleanedText.match(passportPattern);
  if (match && match[1]) {
    return match[1];
  }

  // 模式3：宽松模式 - 数字开头，包含星号，最后是中文姓名
  const loosePattern = /\d+\*{4,}[\dXx]+([\u4e00-\u9fa5]{2,4})(?=[0-9_a-zA-Z]|电子客票号|$)/;

  match = cleanedText.match(loosePattern);
  if (match && match[1]) {
    return match[1];
  }

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
    return match[1];
  }

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
      return suffix;
    }
  }

  // 如果没有下划线，尝试直接使用文件名（去除数字前缀）
  const numericSuffixMatch = nameWithoutExt.match(/\d+/);
  if (numericSuffixMatch) {
    return numericSuffixMatch[0];
  }

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
  // 提取姓名
  const name = extractName(parsedContent.fullText);
  if (!name) {
    return null;
  }

  // 提取文件后缀
  const suffix = extractFileSuffix(originalFileName);
  if (!suffix) {
    return null;
  }

  return `${name}_${suffix}.pdf`;
}

/**
 * 获取 PDF.js 版本信息
 * 用于调试
 */
export function getPDFJSInfo(): void {
  // 保留接口以兼容调用方，不再输出到控制台
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
          // 提取嵌套的 ZIP 文件
          const zipBlob = await zipEntry.async('blob');
          const nestedZipFile = new File([zipBlob], zipEntry.name, {
            type: 'application/zip'
          });

          // 递归处理嵌套的 ZIP
          await processZip(nestedZipFile, `${basePath}${filePath}/`, depth + 1);
          processedCount++;
          continue;
        }

        // 检查是否是 PDF 文件
        if (filePath.toLowerCase().endsWith('.pdf')) {
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

          processedCount++;
        }
      }
    } catch (error) {
      throw new Error(`处理 ZIP 文件失败: ${error.message}`);
    }
  }

  await processZip(zipFile);

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
  const zip = new JSZip();

  // 添加所有文件到 ZIP
  for (let i = 0; i < fileItems.length; i++) {
    const item = fileItems[i];
    if (item.newFileName) {
      zip.file(item.newFileName, item.file);
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

  return zipBlob;
}
