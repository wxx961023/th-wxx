<template>
  <div class="bill-split-container">
    <div class="upload-section">
      <el-upload
        class="upload-dragger"
        drag
        :auto-upload="false"
        :on-change="handleFileChange"
        :before-upload="beforeUpload"
        accept=".xlsx,.xls"
        :show-file-list="false"
      >
        <el-icon class="el-icon--upload"><upload-filled /></el-icon>
        <div class="el-upload__text">
          将Excel文件拖到此处，或<em>点击上传</em>
        </div>
        <template #tip>
          <div class="el-upload__tip">
            只能上传 xlsx/xls 文件，且不超过 10MB
          </div>
        </template>
      </el-upload>
    </div>

    <!-- PDF上传区域 - 仅对戴德梁行显示 -->
    <div v-if="uploadedFile" class="pdf-upload-section">
      <el-card class="pdf-upload-card">
        <template #header>
          <div class="card-header">
            <span>PDF文件上传（印刷序号提取）</span>
          </div>
        </template>

        <el-upload
          class="pdf-uploader"
          accept=".pdf,.zip"
          :http-request="noopRequest"
          :on-change="handlePdfFileChange"
          :show-file-list="true"
          :multiple="true"
          :limit="10"
          :on-remove="handlePdfRemove"
          :auto-upload="false"
          drag
        >
          <el-icon class="el-icon--upload">
            <upload-filled />
          </el-icon>
          <div class="el-upload__text">
            将PDF文件或ZIP压缩包拖到此处，或<em>点击上传</em>
          </div>
          <template #tip>
            <div class="el-upload__tip">
              支持上传PDF文件或ZIP压缩包（ZIP包可包含多层文件夹中的PDF文件），用于提取印刷序号(发票号码)和备注信息
            </div>
          </template>
        </el-upload>

        <!-- PDF提取结果预览 -->
        <div v-if="pdfData.length > 0" class="pdf-data-preview">
          <el-divider content-position="left">
            <span>PDF提取结果预览（{{ pdfData.length }}条记录）</span>
          </el-divider>
          <el-table :data="pdfData" border stripe max-height="400">
            <el-table-column type="index" label="序号" width="60" align="center" />
            <el-table-column prop="ticketNumber" label="电子客票号" width="150" />
            <el-table-column prop="invoiceNumber" label="印刷序号(发票号码)" width="220" />
            <el-table-column prop="remark" label="备注" />
            <el-table-column prop="pageNum" label="页码" width="80" />
            <el-table-column prop="confidence" label="置信度" width="100">
              <template #default="{ row }">
                <el-tag :type="row.confidence > 0.8 ? 'success' : row.confidence > 0.6 ? 'warning' : 'danger'">
                  {{ (row.confidence * 100).toFixed(1) }}%
                </el-tag>
              </template>
            </el-table-column>
          </el-table>
        </div>

        <!-- PDF处理状态 -->
        <div v-if="pdfLoading" class="pdf-loading">
          <el-icon class="is-loading">
            <loading />
          </el-icon>
          <p>正在解析PDF文件...</p>
        </div>
      </el-card>
    </div>

    <!-- 数据展示区域 -->
    <div v-if="showData && getGroupInfo().length > 0" class="data-section">
      <div class="data-header">
        <h3>乘机人部门拆分 - 按公司分组信息</h3>
        <div class="header-buttons">
          <el-button
            type="success"
            :loading="generating"
            @click="generateGroupedExcelFiles"
            :disabled="!showData"
          >
            {{ generating ? "生成中..." : "生成拆分Excel文件" }}
          </el-button>
        </div>
      </div>

      <div class="data-summary">
        <el-alert
          title="分组概览"
          type="info"
          :description="`检测到 ${getGroupCount()} 个公司，将生成一个包含 ${getGroupCount()} 个工作表的Excel文件。点击公司名称可查看详细数据。`"
          show-icon
        />
      </div>

      <div class="data-table">
        <el-table :data="getGroupInfo()" border style="width: 100%">
          <el-table-column prop="groupName" label="公司名称" width="300">
            <template #default="scope">
              <div
                class="company-name"
                :class="{ 'selected': selectedCompany === scope.row.groupName }"
                @click="handleCompanyClick(scope.row.groupName)"
              >
                {{ scope.row.groupName }}
                <span v-if="scope.row.flightInfo" class="ml-2 text-sm text-gray-500">
                  ({{ scope.row.flightInfo.count }}条)
                </span>
              </div>
            </template>
          </el-table-column>
          <el-table-column label="机票数据" width="150">
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

        <!-- 详细数据表格 -->
        <div v-if="selectedCompany" class="detail-table mt-6">
          <h3 class="mb-4 text-lg font-semibold">
            {{ selectedCompany }} - 详细数据
            <span class="text-sm text-gray-500 ml-2">
              (共 {{ getSelectedCompanyDetails().length }} 行)
            </span>
          </h3>
          <el-table
            :data="getSelectedCompanyDetails()"
            border
            style="width: 100%"
            max-height="400"
            stripe
          >
            <el-table-column
              type="index"
              label="序号"
              width="60"
              :index="(index) => index + 1"
            />
            <el-table-column
              v-for="(header, index) in getSelectedCompanyDetails()[0] || []"
              :key="index"
              :label="String(header || `列${index + 1}`)"
              :width="150"
              show-overflow-tooltip
            >
              <template #default="scope">
                {{ scope.row[index] || '' }}
              </template>
            </el-table-column>
          </el-table>
        </div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref } from "vue";
import { ElMessage } from "element-plus";
import { UploadFilled, Loading } from "@element-plus/icons-vue";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { cushmanWakefieldConfig } from "../companyConfig";
import * as pdfjsLib from "pdfjs-dist";
import extractInvoiceInfo from "./extractInvoiceInfo";
import { GlobalWorkerOptions } from "pdfjs-dist";
import JSZip from "jszip";

defineOptions({
  name: "DdlxBillSplit"
});

const uploadedFile = ref<File | null>(null);
const allSheetData = ref<Record<string, any[]>>({});
const loading = ref(false);
const showData = ref(false);
const generating = ref(false);

// 存储每个公司的详细数据
const companyDetails = ref<Record<string, any[]>>({});

// 当前选中的公司
const selectedCompany = ref<string>("");

// PDF相关状态
const uploadedPdfFiles = ref<File[]>([]);
const pdfData = ref<any[]>([]);
const pdfLoading = ref(false);
const pdfProcessingCount = ref(0);

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
          console.log('=== Excel文件加载成功 ===');
          console.log('所有工作表:', workbook.worksheets.map(ws => ws.name));

          // 更灵活的列匹配规则
          const targetColumnPatterns = [
            "乘机人部门（全路径）",
          ];

          // 动态查找包含部门信息的工作表
          const availableSheets: any[] = [];
          workbook.worksheets.forEach(worksheet => {
            // 读取第一行数据来检查是否包含目标列
            const firstRow: any[] = [];
            worksheet.getRow(1).eachCell((cell, colNumber) => {
              firstRow.push(cell.value);
            });

            console.log(`工作表 "${worksheet.name}" 的第一行数据:`, firstRow);

            let hasTargetColumn = false;
            let matchedPattern = "";

            // 检查是否有匹配的列
            for (const pattern of targetColumnPatterns) {
              if (firstRow.some(cell => cell && cell.toString().includes(pattern))) {
                hasTargetColumn = true;
                matchedPattern = pattern;
                break;
              }
            }

            if (hasTargetColumn) {
              // 使用工作表名称作为key，这样更容易识别
              const sheetKey = worksheet.name;
              availableSheets.push({
                name: worksheet.name,
                key: sheetKey,
                departmentKeyword: matchedPattern
              });
              console.log(`找到匹配的工作表: ${worksheet.name}, key: ${sheetKey}, 匹配模式: ${matchedPattern}`);
            } else {
              console.log(`工作表 "${worksheet.name}" 未找到匹配的部门列`);
            }
          });

          console.log(`总共找到 ${availableSheets.length} 个包含目标列的工作表`);

          const sheetData: Record<string, any[]> = {};
          let processedSheets = 0;
          let totalSheets = availableSheets.length;

          if (totalSheets === 0) {
            ElMessage.error("未找到任何包含部门信息的工作表");
            console.log(
              "可用的工作表:",
              workbook.worksheets.map(ws => ws.name)
            );
            console.log(
              "查找的列模式:",
              targetColumnPatterns
            );
            loading.value = false;
            return;
          }

          // 处理每个工作表
          availableSheets.forEach(processor => {
            const worksheet = workbook.getWorksheet(processor.name);
            if (!worksheet) {
              console.log(`跳过不存在的工作表: ${processor.name}`);
              return;
            }

            console.log(
              `\n========== 处理工作表: ${processor.name} ==========`
            );

            // 读取数据为二维数组，确保读取完整的行数据
            const jsonData: any[][] = [];
            worksheet.eachRow((row, rowNumber) => {
              const rowData: any[] = [];

              // 获取工作表的实际列数
              const columnCount = worksheet.columnCount;

              // 确保读取所有列，包括空单元格
              for (let colIndex = 1; colIndex <= columnCount; colIndex++) {
                const cell = row.getCell(colIndex);
                rowData.push(cell.value);
              }

              jsonData.push(rowData);
            });

            sheetData[processor.key] = jsonData;

            console.log(`${processor.name} - 数据行数:`, jsonData.length);
            console.log(`${processor.name} - 工作表列数:`, worksheet.columnCount);
            console.log(`${processor.name} - 第一行列数:`, (jsonData[0] as any[])?.length || 0);
            if (jsonData.length > 1) {
              console.log(`${processor.name} - 第二行列数:`, (jsonData[1] as any[])?.length || 0);
              console.log(`${processor.name} - 第一行数据:`, jsonData[0]);
            }

            processedSheets++;

            // 当所有工作表都处理完成后显示结果
            if (processedSheets === totalSheets) {
              allSheetData.value = sheetData;

              // 处理所有工作表的数据，生成分组信息
              processAllSheetData(sheetData, availableSheets);

              showData.value = true;
              loading.value = false;

              console.log('=== 文件读取完成 ===');
              console.log('allSheetData.value:', allSheetData.value);
              console.log('可用的工作表键:', Object.keys(sheetData));

              ElMessage.success(
                `成功读取 ${totalSheets} 个工作表！请在控制台查看详细信息`
              );
            }
          });
        })
        .catch(error => {
          console.error("读取Excel文件失败:", error);
          ElMessage.error("读取Excel文件失败，请检查文件格式是否正确");
          loading.value = false;
        });
    } catch (error) {
      console.error("文件处理失败:", error);
      ElMessage.error("文件处理失败");
      loading.value = false;
    }
  };

  reader.readAsArrayBuffer(file);
};

// 处理所有工作表数据
const processAllSheetData = (sheetData: Record<string, any[]>, availableSheets: any[]) => {
  console.log('=== 开始处理所有工作表数据 ===');

  Object.entries(sheetData).forEach(([sheetKey, data]) => {
    if (!data || data.length === 0) return;

    console.log(`=== 处理工作表: ${sheetKey} ===`);

    // 查找部门列
    const headers = data[0] as any[];
    const departmentColumnIndex = headers.findIndex(
      (cell: any) => cell && cell.toString().includes("乘机人部门")
    );

    if (departmentColumnIndex === -1) {
      console.log(`工作表 ${sheetKey} 中未找到部门列，跳过`);
      return;
    }

    // 过滤有效数据，但保留完整的行数据以避免列错位
    const validData = data
      .slice(1)
      .filter((row: any[], rowIndex) => {
        const departmentValue = row[departmentColumnIndex];
        if (!departmentValue) return false;

        const departmentText = departmentValue.toString();

        // 过滤掉合计行、总计行等非数据行
        const summaryKeywords = [
          "合计", "总计", "小计", "汇总", "count", "Count", "COUNT", "总数", "张数", "金额"
        ];
        const isSummaryRow = summaryKeywords.some(keyword =>
          departmentText.includes(keyword)
        );

        // 过滤掉纯数字
        const isPureNumber = /^\d+$/.test(departmentText);

        // 过滤掉空值或特殊字符
        const isEmptyOrSpecial =
          departmentText.trim() === "" ||
          /^[\-_=+]+$/.test(departmentText.trim()) ||
          departmentText.length < 2;

        if (isSummaryRow || isPureNumber || isEmptyOrSpecial) {
          console.log(`跳过行 ${rowIndex + 2}: 部门信息="${departmentText}" (类型: ${
            isSummaryRow ? "合计行" : isPureNumber ? "纯数字" : "空值/特殊字符"
          })`);
          return false;
        }

        return true;
      })
      .map((row: any[], rowIndex) => {
        // 确保保留完整的行数据，包括空单元格
        const completeRow = [...row]; // 创建副本以避免修改原数据
        return {
          部门信息: row[departmentColumnIndex],
          完整行数据: completeRow,
          原始行号: rowIndex + 2 // 从第2行开始计数
        };
      });

    // 根据部门信息分组
    const groups = new Map<string, any[]>();
    validData.forEach(item => {
      const fullPath = item.部门信息.toString();

      // 提取公司名称
      let companyName: string;
      if (fullPath.includes("-")) {
        companyName = fullPath.split("-")[0].trim();
      } else {
        companyName = fullPath.trim();
      }

      if (companyName) {
        if (!groups.has(companyName)) {
          groups.set(companyName, []);
        }
        groups.get(companyName)!.push(item);
      }
    });

    // 存储每个公司的详细数据
    groups.forEach((items, companyName) => {
      console.log(`=== 处理公司 ${companyName} 的详细数据 ===`);
      if (!companyDetails.value[companyName]) {
        companyDetails.value[companyName] = [];
      }

      // 获取标准表头
      const { standardHeaders, columnMapping } = mapColumnsToStandard(data[0]);

      // 数据转换函数：处理特殊的列转换逻辑
      const transformRowDataForDetails = (originalRow: any[], standardHeader: string, itemIndex: number) => {
        const originalColIndex = columnMapping[standardHeader];

        if (originalColIndex !== undefined) {
          let value = originalRow[originalColIndex] || '';

          // 特殊处理逻辑
          if (standardHeader === "承运人") {
            // 承运人 = 票号 "-" 分割【0】
            const ticketNumberColIndex = columnMapping["电子客票号"];
            if (ticketNumberColIndex !== undefined) {
              const ticketNumber = originalRow[ticketNumberColIndex] || '';
              if (ticketNumber && typeof ticketNumber === 'string') {
                value = ticketNumber.split('-')[0] || value;
              }
            }
          } else if (standardHeader === "航程") {
            // 航程 = 出发城市-到达城市 来拼接
            const departureCityIndex = columnMapping["出发城市"];
            const arrivalCityIndex = columnMapping["到达城市"];

            console.log(`🔍 数据转换航程调试 - 原始行${itemIndex}:`);
            console.log(`  出发城市映射索引: ${departureCityIndex}`);
            console.log(`  到达城市映射索引: ${arrivalCityIndex}`);

            if (departureCityIndex !== undefined && arrivalCityIndex !== undefined) {
              const departureCity = originalRow[departureCityIndex] || '';
              const arrivalCity = originalRow[arrivalCityIndex] || '';
              console.log(`  出发城市原值: "${departureCity}"`);
              console.log(`  到达城市原值: "${arrivalCity}"`);

              if (departureCity && arrivalCity) {
                value = `${departureCity}-${arrivalCity}`;
                console.log(`  ✅ 生成航程: "${value}"`);
              } else {
                value = departureCity || arrivalCity || '';
                console.log(`  ⚠️ 部分城市为空，生成航程: "${value}"`);
              }
            } else {
              value = '';
              console.log(`  ❌ 未找到出发城市或到达城市列映射`);
            }
          }

          // 处理金额列的格式：在表格显示时保留两位小数，空值赋值为0
          if (standardHeader === "票价" || standardHeader === "燃油附加费" || standardHeader === "民航发展基金" ||
              standardHeader === "保险费" || standardHeader === "改签费" || standardHeader === "退票费" ||
              standardHeader === "小计" || standardHeader === "保险" || standardHeader === "服务费" ||
              standardHeader === "实收" || standardHeader === "机票计税价格（票价+燃油附加费）" || standardHeader === "机票增值税" ||
              standardHeader === "机票不含税金额" || standardHeader === "WD上填列Airfare数" || standardHeader === "代理商服务费增值税" ||
              standardHeader === "代理商不含税服务金额" || standardHeader === "机票增值税+服务费税额" || standardHeader === "Airfare+服务费不含税" ||
              standardHeader === "Checking") {
            const numValue = parseFloat(String(value || '').replace(/,/g, ''));
            if (!isNaN(numValue)) {
              value = numValue.toFixed(2);
            } else {
              value = '0.00'; // 空值或无效值赋值为0
            }
          }

          // 专门调试乘机人列
          if (standardHeader === "乘机人") {
            console.log(`=== 乘机人数据转换调试 ===`);
            console.log(`原始行索引: ${itemIndex}`);
            console.log(`乘机人映射列索引: ${originalColIndex}`);
            console.log(`原始行数据长度: ${originalRow.length}`);
            console.log(`原始行数据:`, originalRow);
            console.log(`乘机人原始值: "${originalRow[originalColIndex]}"`);
            console.log(`转换后值: "${value}"`);
            console.log(`=== 乘机人数据转换调试结束 ===`);
          }

          return value;
        }

        // 特殊处理未映射的列
        if (standardHeader === "序号") {
          return (itemIndex + 1).toString();
        } else if (standardHeader === "部门") {
          // 部门信息从部门列获取
          if (departmentColumnIndex !== -1) {
            return originalRow[departmentColumnIndex] || '';
          }
        } else if (standardHeader === "国际/国内") {
          return "国内";
        } else if (standardHeader === "机票计税价格（票价+燃油附加费）") {
          // 机票计税价格 = 票价 + 燃油附加费
          const ticketPriceIndex = columnMapping["票价"];
          const fuelFeeIndex = columnMapping["燃油附加费"];

          if (ticketPriceIndex !== undefined && fuelFeeIndex !== undefined) {
            const ticketPrice = parseFloat(String(originalRow[ticketPriceIndex] || '').replace(/,/g, '')) || 0;
            const fuelFee = parseFloat(String(originalRow[fuelFeeIndex] || '').replace(/,/g, '')) || 0;
            const taxPrice = ticketPrice + fuelFee;
            return taxPrice.toFixed(2);
          }
          return "0.00";
        } else if (standardHeader === "机票增值税") {
          // 机票增值税 = IF(OR(E3="",I3<>"国内"),0,ROUND(L3/1.09*0.09,2)+ROUND(M3/1.09*0.09,2))
          // E列是出票日期, I列是国际/国内, L列是票价, M列是燃油附加费
          const recordDateIndex = columnMapping["出票日期"];
          const domesticIndex = columnMapping["国际/国内"];
          const ticketPriceIndex = columnMapping["票价"];
          const fuelFeeIndex = columnMapping["燃油附加费"];

          if (recordDateIndex !== undefined && domesticIndex !== undefined &&
              ticketPriceIndex !== undefined && fuelFeeIndex !== undefined) {
            const recordDate = originalRow[recordDateIndex] || '';
            const domestic = originalRow[domesticIndex] || '';
            const ticketPrice = parseFloat(String(originalRow[ticketPriceIndex] || '').replace(/,/g, '')) || 0;
            const fuelFee = parseFloat(String(originalRow[fuelFeeIndex] || '').replace(/,/g, '')) || 0;

            // IF(OR(E3="",I3<>"国内"),0,ROUND(L3/1.09*0.09,2)+ROUND(M3/1.09*0.09,2))
            if (!recordDate || domestic !== "国内") {
              return "0.00";
            } else {
              const ticketTax = Math.round(ticketPrice / 1.09 * 0.09 * 100) / 100;
              const fuelTax = Math.round(fuelFee / 1.09 * 0.09 * 100) / 100;
              const totalTax = ticketTax + fuelTax;
              return totalTax.toFixed(2);
            }
          }
          return "0.00";
        } else if (standardHeader === "机票不含税金额") {
          // 机票不含税金额 = Y3-Z3 (机票计税价格 - 机票增值税)
          const taxPriceIndex = columnMapping["机票计税价格（票价+燃油附加费）"];
          const taxIndex = columnMapping["机票增值税"];

          if (taxPriceIndex !== undefined && taxIndex !== undefined) {
            const taxPrice = parseFloat(String(originalRow[taxPriceIndex] || '').replace(/,/g, '')) || 0;
            const tax = parseFloat(String(originalRow[taxIndex] || '').replace(/,/g, '')) || 0;
            const noTaxAmount = taxPrice - tax;
            return noTaxAmount.toFixed(2);
          }
          return "0.00";
        } else if (standardHeader === "WD上填列Airfare数") {
          // WD上填列Airfare数 = AA3+N3+O3+Q3 (机票不含税金额 + 票价 + 燃油附加费 + 保险费)
          const noTaxAmountIndex = columnMapping["机票不含税金额"];
          const ticketPriceIndex = columnMapping["票价"];
          const fuelFeeIndex = columnMapping["燃油附加费"];
          const insuranceFeeIndex = columnMapping["保险费"];

          if (noTaxAmountIndex !== undefined && ticketPriceIndex !== undefined &&
              fuelFeeIndex !== undefined && insuranceFeeIndex !== undefined) {
            const noTaxAmount = parseFloat(String(originalRow[noTaxAmountIndex] || '').replace(/,/g, '')) || 0;
            const ticketPrice = parseFloat(String(originalRow[ticketPriceIndex] || '').replace(/,/g, '')) || 0;
            const fuelFee = parseFloat(String(originalRow[fuelFeeIndex] || '').replace(/,/g, '')) || 0;
            const insuranceFee = parseFloat(String(originalRow[insuranceFeeIndex] || '').replace(/,/g, '')) || 0;
            const airfareAmount = noTaxAmount + ticketPrice + fuelFee + insuranceFee;
            return airfareAmount.toFixed(2);
          }
          return "0.00";
        } else if (standardHeader === "代理商服务费增值税") {
          // 代理商服务费增值税 = ROUND(T3/1.06*0.06,2)
          const serviceFeeIndex = columnMapping["小计"];

          if (serviceFeeIndex !== undefined) {
            const serviceFee = parseFloat(String(originalRow[serviceFeeIndex] || '').replace(/,/g, '')) || 0;
            const serviceFeeTax = Math.round(serviceFee / 1.06 * 0.06 * 100) / 100;
            return serviceFeeTax.toFixed(2);
          }
          return "0.00";
        } else if (standardHeader === "代理商不含税服务金额") {
          // 代理商不含税服务金额 = T3-AC3 (小计 - 代理商服务费增值税)
          const serviceFeeIndex = columnMapping["小计"];
          const serviceFeeTaxIndex = columnMapping["代理商服务费增值税"];

          if (serviceFeeIndex !== undefined && serviceFeeTaxIndex !== undefined) {
            const serviceFee = parseFloat(String(originalRow[serviceFeeIndex] || '').replace(/,/g, '')) || 0;
            const serviceFeeTax = parseFloat(String(originalRow[serviceFeeTaxIndex] || '').replace(/,/g, '')) || 0;
            const noTaxServiceFee = serviceFee - serviceFeeTax;
            return noTaxServiceFee.toFixed(2);
          }
          return "0.00";
        } else if (standardHeader === "机票增值税+服务费税额") {
          // 机票增值税+服务费税额 = Z3+AC3 (机票增值税 + 代理商服务费增值税)
          const ticketTaxIndex = columnMapping["机票增值税"];
          const serviceFeeTaxIndex = columnMapping["代理商服务费增值税"];

          if (ticketTaxIndex !== undefined && serviceFeeTaxIndex !== undefined) {
            const ticketTax = parseFloat(String(originalRow[ticketTaxIndex] || '').replace(/,/g, '')) || 0;
            const serviceFeeTax = parseFloat(String(originalRow[serviceFeeTaxIndex] || '').replace(/,/g, '')) || 0;
            const totalTax = ticketTax + serviceFeeTax;
            return totalTax.toFixed(2);
          }
          return "0.00";
        } else if (standardHeader === "Airfare+服务费不含税") {
          // Airfare+服务费不含税 = AB3+AD3 (WD上填列Airfare数 + 代理商不含税服务金额)
          const airfareIndex = columnMapping["WD上填列Airfare数"];
          const noTaxServiceFeeIndex = columnMapping["代理商不含税服务金额"];

          if (airfareIndex !== undefined && noTaxServiceFeeIndex !== undefined) {
            const airfare = parseFloat(String(originalRow[airfareIndex] || '').replace(/,/g, '')) || 0;
            const noTaxServiceFee = parseFloat(String(originalRow[noTaxServiceFeeIndex] || '').replace(/,/g, '')) || 0;
            const totalNoTax = airfare + noTaxServiceFee;
            return totalNoTax.toFixed(2);
          }
          return "0.00";
        } else if (standardHeader === "Checking") {
          // Checking = W3-Z3-AB3-AC3-AD3 (总金额 - 机票增值税 - WD上填列Airfare数 - 代理商服务费增值税 - 代理商不含税服务金额)
          const totalAmountIndex = columnMapping["实收"];
          const ticketTaxIndex = columnMapping["机票增值税"];
          const airfareIndex = columnMapping["WD上填列Airfare数"];
          const serviceFeeTaxIndex = columnMapping["代理商服务费增值税"];
          const noTaxServiceFeeIndex = columnMapping["代理商不含税服务金额"];

          if (totalAmountIndex !== undefined && ticketTaxIndex !== undefined &&
              airfareIndex !== undefined && serviceFeeTaxIndex !== undefined && noTaxServiceFeeIndex !== undefined) {
            const totalAmount = parseFloat(String(originalRow[totalAmountIndex] || '').replace(/,/g, '')) || 0;
            const ticketTax = parseFloat(String(originalRow[ticketTaxIndex] || '').replace(/,/g, '')) || 0;
            const airfare = parseFloat(String(originalRow[airfareIndex] || '').replace(/,/g, '')) || 0;
            const serviceFeeTax = parseFloat(String(originalRow[serviceFeeTaxIndex] || '').replace(/,/g, '')) || 0;
            const noTaxServiceFee = parseFloat(String(originalRow[noTaxServiceFeeIndex] || '').replace(/,/g, '')) || 0;
            const checking = totalAmount - ticketTax - airfare - serviceFeeTax - noTaxServiceFee;
            return checking.toFixed(2);
          }
          return "0.00";
        }

        return '';
      };

      // 转换原始数据为标准格式
      const transformedData = items.map((item, itemIndex) => {
        const originalRow = item.完整行数据;
        const standardRow: any[] = [];

        // 根据标准表头生成新行数据
        standardHeaders.forEach((standardHeader, index) => {
          standardRow[index] = transformRowDataForDetails(originalRow, standardHeader, itemIndex);
        });

        return standardRow;
      });

      console.log(`${companyName} - 标准表头列数: ${standardHeaders.length}`);
      console.log(`${companyName} - 转换后数据样例列数: ${transformedData[0]?.length || 0}`);

      // 如果还没有数据，先添加标准表头
      if (companyDetails.value[companyName].length === 0) {
        companyDetails.value[companyName].push(standardHeaders);
        console.log(`${companyName} - 添加标准表头，列数: ${standardHeaders.length}`);
      }

      companyDetails.value[companyName].push(...transformedData);
      console.log(`${companyName} - 添加 ${transformedData.length} 条详细数据后，总长度: ${companyDetails.value[companyName].length}`);
    });
  });
};

// 获取分组信息
const getGroupInfo = () => {
  console.log('🔍 getGroupInfo 开始执行');
  console.log('📊 allSheetData.value:', Object.keys(allSheetData.value));
  const companyGroups = new Map<string, any>();

  Object.entries(allSheetData.value).forEach(([sheetKey, sheetData]) => {
    console.log(`📋 处理工作表: ${sheetKey}, 数据长度: ${sheetData?.length}`);
    if (!sheetData || sheetData.length === 0) {
      console.log(`  ❌ 工作表 ${sheetKey} 无数据`);
      return;
    }

    // 查找部门列
    const headers = sheetData[0] as any[];
    console.log(`  📝 表头数据:`, headers);
    const departmentColumnIndex = headers.findIndex(
      (cell: any) => cell && cell.toString().includes("乘机人部门")
    );

    console.log(`  🎯 部门列索引: ${departmentColumnIndex}`);
    if (departmentColumnIndex === -1) {
      console.log(`  ❌ 工作表 ${sheetKey} 未找到"乘机人部门"列`);
      return;
    }

    // 统计该公司在此工作表中的数据
    const companyCountMap = new Map<string, number>();

    sheetData.slice(1).forEach((row: any[]) => {
      const departmentValue = row[departmentColumnIndex];
      if (!departmentValue) return;

      const departmentText = departmentValue.toString();

      // 过滤掉非有效数据
      const summaryKeywords = [
        "合计", "总计", "小计", "汇总", "count", "Count", "COUNT", "总数", "张数", "金额"
      ];
      const isSummaryRow = summaryKeywords.some(keyword =>
        departmentText.includes(keyword)
      );

      const isPureNumber = /^\d+$/.test(departmentText);
      const isEmptyOrSpecial =
        departmentText.trim() === "" ||
        /^[\-_=+]+$/.test(departmentText.trim()) ||
        departmentText.length < 2;

      if (isSummaryRow || isPureNumber || isEmptyOrSpecial)
        return;

      // 提取公司名称
      let companyName: string;
      if (departmentText.includes("-")) {
        companyName = departmentText.split("-")[0].trim();
      } else {
        companyName = departmentText.trim();
      }

      if (companyName) {
        companyCountMap.set(companyName, (companyCountMap.get(companyName) || 0) + 1);
      }
    });

    // 更新公司分组信息
    companyCountMap.forEach((count, companyName) => {
      if (!companyGroups.has(companyName)) {
        companyGroups.set(companyName, {
          groupName: companyName,
          totalCount: 0,
          editableFileName: companyName
        });
      }

      const group = companyGroups.get(companyName)!;
      if (sheetKey.includes('机票') || sheetKey.includes('航班')) {
        group.flightInfo = {
          count: count,
          rowRange: `数据行${count}条`
        };
      }
      group.totalCount += count;
    });
  });

  const result = Array.from(companyGroups.values());
  console.log('🎯 getGroupInfo 最终结果:', result);
  console.log('📈 分组数量:', result.length);
  return result;
};

// 获取分组数量
const getGroupCount = () => {
  return getGroupInfo().length;
};

// 处理公司点击事件
const handleCompanyClick = (companyName: string) => {
  if (selectedCompany.value === companyName) {
    selectedCompany.value = "";
  } else {
    selectedCompany.value = companyName;
  }
};

// 获取选中公司的详细数据
const getSelectedCompanyDetails = () => {
  if (!selectedCompany.value || !companyDetails.value[selectedCompany.value]) {
    return [];
  }
  return [...companyDetails.value[selectedCompany.value]];
};

// 更新文件名
const updateFileName = (groupName: string, newFileName: string) => {
  const groupInfo = getGroupInfo();
  const group = groupInfo.find(g => g.groupName === groupName);
  if (group) {
    group.editableFileName = newFileName;
  }
};

// 生成文件名
const generateFileName = (groupName: string) => {
  return groupName;
};

// 列映射函数：将原表列映射到标准表头
const mapColumnsToStandard = (originalHeaders: string[]) => {
  console.log('=== 开始列映射调试 ===');
  console.log('原始表头:', originalHeaders);

  // 标准表头定义
  const standardHeaders = [
    "序号", "出票日期", "承运人", "印刷序号(发票号码)", "电子客票号",
    "乘机人", "部门", "乘机日期", "国际/国内", "航程", "航班",
    "票价", "燃油附加费", "民航发展基金", "保险费", "改签费",
    "退票费", "小计", "保险", "服务费", "改签费", "退票费", "实收", "备注", "机票计税价格（票价+燃油附加费）", "机票增值税", "机票不含税金额", "WD上填列Airfare数", "代理商服务费增值税", "代理商不含税服务金额", "机票增值税+服务费税额", "Airfare+服务费不含税", "Checking"
  ];

  // 列映射规则
  const columnMapping: Record<string, number> = {};

  originalHeaders.forEach((header, index) => {
    const headerText = header ? header.toString().toLowerCase().trim() : "";
    console.log(`处理列 ${index}: "${header}" -> "${headerText}"`);

    // 专门调试乘机人列
    if (index === 1 || (header && header.toString().includes("乘机人"))) {
      console.log(`🔍 乘机人列详细调试:`);
      console.log(`  - 原始值: "${header}"`);
      console.log(`  - 类型: ${typeof header}`);
      console.log(`  - 长度: ${header ? header.toString().length : 'null'}`);
      console.log(`  - 转换后: "${headerText}"`);
      console.log(`  - headerText.includes("乘机人"): ${headerText.includes("乘机人")}`);
      console.log(`  - "乘机人".includes(headerText): ${"乘机人".includes(headerText)}`);
      if (header) {
        const headerStr = header.toString();
        console.log(`  - 字符编码: ${Array.from(headerStr).map(c => `${c}(${c.charCodeAt(0)})`).join(', ')}`);
      }
    }

    // 根据您提供的映射关系进行匹配
    if (headerText.includes("序号") || headerText.includes("no") || headerText.includes("#")) {
      columnMapping["序号"] = index;
      console.log(`  -> 映射到"序号"`);
    } else if (headerText.includes("出票日期") || headerText.includes("记账日期")) {
      columnMapping["出票日期"] = index;
      console.log(`  -> 映射到"出票日期"`);
    } else if (headerText.includes("承运人") || headerText.includes("航空公司")) {
      columnMapping["承运人"] = index;
      console.log(`  -> 映射到"承运人"`);
    } else if (headerText.includes("印刷序号") ) {
      columnMapping["印刷序号(发票号码)"] = index;
      console.log(`  -> 映射到"印刷序号(发票号码)"`);
    } else if (headerText === "票号" ) {
      columnMapping["电子客票号"] = index;
      console.log(`  -> 映射到"电子客票号"`);
    } else if ((headerText === "乘机人")) {
      columnMapping["乘机人"] = index;
      console.log(`  -> 映射到"乘机人" (关键映射!)`);
      console.log(`✅ 成功! index=${index}, headerText="${headerText}"`);
    } else if (headerText === "乘机人部门") {
      columnMapping["部门"] = index;
      console.log(`  -> 映射到"部门"`);
    } else if (headerText === "出发日期") {
      columnMapping["乘机日期"] = index;
      console.log(`  -> 映射到"乘机日期"`);
    } else if (headerText === "国际") {
      columnMapping["国际/国内"] = index;
      console.log(`  -> 映射到"国际/国内"`);
    } else if (headerText === "出发城市") {
      // 出发城市列，用于航程拼接
      columnMapping["出发城市"] = index;
      console.log(`  -> 映射到"出发城市"，列索引: ${index}`);
    } else if (headerText === "到达城市") {
      // 到达城市列，用于航程拼接
      columnMapping["到达城市"] = index;
      console.log(`  -> 映射到"到达城市"，列索引: ${index}`);
    } else if (headerText.includes("航班") || headerText.includes("航班号")) {
      columnMapping["航班"] = index;
      console.log(`  -> 映射到"航班"`);
    } else if (headerText.includes("票价") || headerText.includes("票面价")) {
      columnMapping["票价"] = index;
      console.log(`  -> 映射到"票价"`);
    } else if (headerText.includes("燃油附加费") || headerText.includes("燃油")) {
      columnMapping["燃油附加费"] = index;
      console.log(`  -> 映射到"燃油附加费"`);
    } else if (headerText.includes("民航发展基金") || headerText.includes("发展基金") || headerText.includes("基建费") || headerText.includes("机建")) {
      columnMapping["民航发展基金"] = index;
      console.log(`  -> 映射到"民航发展基金"`);
    } else if (headerText.includes("保险费") || headerText.includes("保险")) {
      // 优先映射到"保险费"
      if (!columnMapping["保险费"]) {
        columnMapping["保险费"] = index;
        console.log(`  -> 映射到"保险费"`);
      }
    } else if (headerText.includes("改签费")) {
      columnMapping["改签费"] = index;
      console.log(`  -> 映射到"改签费"`);
    } else if (headerText.includes("退票费")) {
      columnMapping["退票费"] = index;
      console.log(`  -> 映射到"退票费"`);
    } else if (headerText.includes("小计")) {
      columnMapping["小计"] = index;
      console.log(`  -> 映射到"小计"`);
    } else if (headerText.includes("服务费") || headerText.includes("系统使用费")) {
      columnMapping["服务费"] = index;
      console.log(`  -> 映射到"服务费"`);
    } else if (headerText.includes("实收") || headerText.includes("总金额") || headerText.includes("实付") || headerText.includes("合计")) {
      columnMapping["实收"] = index;
      console.log(`  -> 映射到"实收"`);
    } else if (headerText.includes("备注") || headerText.includes("说明")) {
      columnMapping["备注"] = index;
      console.log(`  -> 映射到"备注"`);
    } else {
      console.log(`  -> 未匹配到任何标准列`);
    }
  });

  console.log('=== 乘机人列映射调试 ===');
  console.log('乘机人列映射索引:', columnMapping["乘机人"]);
  if (columnMapping["乘机人"] !== undefined) {
    console.log('乘机人原始列名:', originalHeaders[columnMapping["乘机人"]]);
  } else {
    console.log('❌ 乘机人列未映射! 这就是问题所在');
  }

  console.log('最终列映射结果:', columnMapping);
  console.log('🔍 部门列映射调试:');
  console.log('  - 部门映射索引:', columnMapping["部门"]);
  if (columnMapping["部门"] !== undefined) {
    console.log('  - 部门原始列名:', originalHeaders[columnMapping["部门"]]);
  } else {
    console.log('  - ❌ 部门列未映射!');
  }
  console.log('=== 列映射调试结束 ===');
  return { standardHeaders, columnMapping };
};

// 生成分组Excel文件
const generateGroupedExcelFiles = async () => {
  console.log('🚀 generateGroupedExcelFiles 函数开始执行');
  generating.value = true;
  const groupInfo = getGroupInfo();
  console.log(`📊 groupInfo 长度: ${groupInfo.length}`, groupInfo);

  try {
    console.log(`开始生成分组Excel文件，共 ${groupInfo.length} 个公司`);

    // 创建一个工作簿，包含所有公司的工作表
    const newWorkbook = new ExcelJS.Workbook();

    // 为每个公司创建一个工作表
    for (const companyGroup of groupInfo) {
      console.log(`为公司 ${companyGroup.groupName} 创建工作表`);

      // 获取工作表名称，如果是戴德梁行公司，使用配置的shortName
      let worksheetName = companyGroup.groupName;
      const companyInfo = cushmanWakefieldConfig.getCompanyInfo(companyGroup.groupName);
      if (companyInfo.shortName && companyInfo.shortName !== companyGroup.groupName) {
        worksheetName = companyInfo.shortName;
        console.log(`  使用配置的短名称: ${companyInfo.shortName}`);
      }

      const worksheet = newWorkbook.addWorksheet(worksheetName, {
        views: [{ showGridLines: true }]
      });
      worksheet.properties.defaultRowHeight = 40;

      let hasData = false;
      const departmentSumRows: Map<string, number> = new Map(); // 记录每个部门的求和行行号

      // 处理所有原始工作表数据，合并到这个公司的工作表中
      Object.entries(allSheetData.value).forEach(([originalSheetKey, sheetData]) => {
        if (!sheetData || sheetData.length === 0) return;

        // 查找部门列
        const headers = sheetData[0] as any[];
        const departmentColumnIndex = headers.findIndex(
          (cell: any) => cell && cell.toString().includes("乘机人部门")
        );

        if (departmentColumnIndex === -1) return;

        // 筛选该公司的数据，保留完整行以避免列错位
        const companyData = sheetData
          .slice(1)
          .filter((row: any[]) => {
            const departmentValue = row[departmentColumnIndex];
            if (!departmentValue) return false;

            const departmentText = departmentValue.toString();

            // 过滤掉非有效数据
            const summaryKeywords = [
              "合计", "总计", "小计", "汇总", "count", "Count", "COUNT", "总数", "张数", "金额"
            ];
            const isSummaryRow = summaryKeywords.some(keyword =>
              departmentText.includes(keyword)
            );

            const isPureNumber = /^\d+$/.test(departmentText);
            const isEmptyOrSpecial =
              departmentText.trim() === "" ||
              /^[\-_=+]+$/.test(departmentText.trim()) ||
              departmentText.length < 2;

            if (isSummaryRow || isPureNumber || isEmptyOrSpecial)
              return false;

            // 提取公司名称进行匹配
            let companyName: string;
            if (departmentText.includes("-")) {
              companyName = departmentText.split("-")[0].trim();
            } else {
              companyName = departmentText.trim();
            }

            return companyName === companyGroup.groupName;
          })
          .map(row => {
            // 确保保留完整的行数据，包括空单元格
            return [...row];
          });

        if (companyData.length > 0) {
          hasData = true;

          console.log(`  工作表 ${originalSheetKey}: 表头列数=${headers.length}, 数据样例列数=${companyData[0]?.length}`);

          // 获取列映射
          const { standardHeaders, columnMapping } = mapColumnsToStandard(headers);

          // 数据转换函数：处理特殊的列转换逻辑
          const transformRowData = (originalRow: any[], standardHeader: string) => {
            const originalColIndex = columnMapping[standardHeader];

            if (originalColIndex !== undefined) {
              let value = originalRow[originalColIndex] || '';

              // 特殊处理逻辑
              if (standardHeader === "承运人") {
                // 承运人 = 票号 "-" 分割【0】
                const ticketNumberColIndex = columnMapping["电子客票号"];
                if (ticketNumberColIndex !== undefined) {
                  const ticketNumber = originalRow[ticketNumberColIndex] || '';
                  if (ticketNumber && typeof ticketNumber === 'string') {
                    value = ticketNumber.split('-')[0] || value;
                  }
                }
              } else if (standardHeader === "航程") {
                // 航程 = 出发城市-到达城市 来拼接
                const departureCityIndex = columnMapping["出发城市"];
                const arrivalCityIndex = columnMapping["到达城市"];

                if (departureCityIndex !== undefined && arrivalCityIndex !== undefined) {
                  const departureCity = originalRow[departureCityIndex] || '';
                  const arrivalCity = originalRow[arrivalCityIndex] || '';
                  if (departureCity && arrivalCity) {
                    value = `${departureCity}-${arrivalCity}`;
                  } else {
                    value = departureCity || arrivalCity || '';
                  }
                } else {
                  value = '';
                }
              }


              // 专门调试乘机人列
              if (standardHeader === "乘机人") {
                console.log(`=== Excel生成乘机人转换调试 ===`);
                console.log(`乘机人映射列索引: ${originalColIndex}`);
                console.log(`原始行数据:`, originalRow);
                console.log(`乘机人原始值: "${originalRow[originalColIndex]}"`);
                console.log(`转换后值: "${value}"`);
                console.log(`=== Excel生成乘机人转换调试结束 ===`);
              }

              return value;
            }

            // 特殊处理未映射的列
            if (standardHeader === "序号") {
              return ''; // 序号会在后面统一生成
            }

            return '';
          };

          // 如果这是第一个有数据的工作表，添加标准标题行
          if (worksheet.rowCount === 0) {
            // 添加标准标题行
            standardHeaders.forEach((header, colIndex) => {
              const cell = worksheet.getCell(1, colIndex + 1);
              cell.value = header;
              cell.font = { bold: true };
              // 特殊处理表头颜色
              if (header === "序号") {
                cell.fill = {
                  type: 'pattern',
                  pattern: 'solid',
                  fgColor: { argb: 'FFB6CEA3' } // #B6CEA3 背景色
                } as any;
              } else if (header === "出票日期" || header === "承运人" || header === "乘机人" ||
                        header === "乘机日期" || header === "航程" || header === "航班" ||
                        header === "票价" || header === "民航发展基金" || header === "保险费" ||
                        header === "改签费" || header === "小计" || header === "服务费" ||
                        header === "保险" || header === "退票费" || header === "实收" || header === "备注") {
                cell.fill = {
                  type: 'pattern',
                  pattern: 'solid',
                  fgColor: { argb: 'FFC9E4B4' } // #C9E4B4 背景色
                } as any;
              } else if ([
                "印刷序号(发票号码)", "电子客票号", "部门", "国际/国内", "燃油附加费",
                "机票计税价格（票价+燃油附加费）", "机票不含税金额", "Checking"
              ].includes(header)) {
                cell.fill = {
                  type: 'pattern',
                  pattern: 'solid',
                  fgColor: { argb: 'FFFFFF00' } // #FFFF00 背景色
                } as any;
              } else if ([
                "WD上填列Airfare数", "代理商服务费增值税", "代理商不含税服务金额"
              ].includes(header)) {
                cell.fill = {
                  type: 'pattern',
                  pattern: 'solid',
                  fgColor: { argb: 'FFFDE38A' } // #FDE38A 背景色
                } as any;
              } else if (header === "机票增值税") {
                cell.fill = {
                  type: 'pattern',
                  pattern: 'solid',
                  fgColor: { argb: 'FF00B0F0' } // #00B0F0 背景色
                } as any;
              } else if ([
                "机票增值税+服务费税额", "Airfare+服务费不含税"
              ].includes(header)) {
                cell.fill = {
                  type: 'pattern',
                  pattern: 'solid',
                  fgColor: { argb: 'FFF6C9A1' } // #F6C9A1 背景色
                } as any;
              } else {
                cell.fill = {
                  type: 'pattern',
                  pattern: 'solid',
                  fgColor: { argb: 'FFE6F3FF' }
                };
              }
              cell.border = {
                top: { style: "thin" },
                bottom: { style: "thin" },
                left: { style: "thin" },
                right: { style: "thin" }
              };
              cell.alignment = {
                horizontal: "center",
                vertical: "middle"
              };
            });
            console.log(`  工作表 ${originalSheetKey}: 使用标准表头，共 ${standardHeaders.length} 列`);

            // 设置表头行高为38磅
            worksheet.getRow(1).height = 38;
          }

          // 按部门分组数据
          const departmentMappingIndex = columnMapping["部门"];
          const groupedData: Record<string, any[]> = {};

          // 清空部门求和行记录，为新的原始工作表做准备
          departmentSumRows.clear();

          console.log(`🔍 开始部门分组，部门映射索引: ${departmentMappingIndex}`);

          companyData.forEach((row, rowIndex) => {
            let department = '';
            if (departmentMappingIndex !== undefined && departmentMappingIndex !== -1) {
              department = row[departmentMappingIndex] || '未知部门';
            } else {
              department = '未知部门';
            }

            console.log(`  行${rowIndex} -> 部门: "${department}"`);

            if (!groupedData[department]) {
              groupedData[department] = [];
            }
            groupedData[department].push(row);
          });

          console.log(`分组结果:`, Object.keys(groupedData).map(key => `${key}: ${groupedData[key].length}条`));

          // 添加分组后的数据行
          let globalRowIndex = 0; // 全局行号，用于生成序号

          Object.entries(groupedData).forEach(([department, departmentRows], departmentIndex) => {
            console.log(`处理部门 ${departmentIndex + 1}: "${department}" (${departmentRows.length}条数据)`);

            // 添加该部门的数据行
            departmentRows.forEach((row, rowIndex) => {
              const actualRowIndex = worksheet.rowCount + 1;

              // 根据标准表头列数添加数据
              standardHeaders.forEach((standardHeader, colIndex) => {
                const cell = worksheet.getCell(actualRowIndex, colIndex + 1);

                // 特殊处理：序号列自动生成（全局递增）
                if (standardHeader === "序号") {
                  globalRowIndex++;
                  cell.value = globalRowIndex.toString();
                  // 移除数据行背景色，只保留表头背景色
                } else if (standardHeader === "部门") {
                  // 部门信息从"乘机人部门"列获取
                  console.log(`🔍 Excel生成部门调试 - 行${rowIndex}: 部门值="${department}"`);
                  cell.value = department;
                  console.log(`  ✅ 设置部门值: "${cell.value}"`);
                } else if (standardHeader === "国际/国内") {
                  cell.value = "国内";
                } else if (standardHeader === "航程") {
                  // 航程 = 出发城市-到达城市 来拼接
                  const departureCityIndex = columnMapping["出发城市"];
                  const arrivalCityIndex = columnMapping["到达城市"];

                  console.log(`🔍 Excel生成航程调试 - 部门"${department}"行${rowIndex}:`);
                  console.log(`  出发城市映射索引: ${departureCityIndex}`);
                  console.log(`  到达城市映射索引: ${arrivalCityIndex}`);

                  if (departureCityIndex !== undefined && arrivalCityIndex !== undefined) {
                    const departureCity = row[departureCityIndex] || '';
                    const arrivalCity = row[arrivalCityIndex] || '';
                    console.log(`  出发城市原值: "${departureCity}"`);
                    console.log(`  到达城市原值: "${arrivalCity}"`);

                    if (departureCity && arrivalCity) {
                      cell.value = `${departureCity}-${arrivalCity}`;
                      console.log(`  ✅ 生成航程: "${cell.value}"`);
                    } else {
                      cell.value = departureCity || arrivalCity || '';
                      console.log(`  ⚠️ 部分城市为空，生成航程: "${cell.value}"`);
                    }
                  } else {
                    cell.value = '';
                    console.log(`  ❌ 未找到出发城市或到达城市列映射`);
                  }
                } else if (colIndex === 14 || colIndex === 15 || colIndex === 16 || colIndex === 17 || colIndex === 18) {
                  // O(14), P(15), Q(16), R(17), S(18)列设置为0
                  cell.value = 0;
                  cell.numFmt = '#,##0.00';
                } else {
                  // 使用转换函数获取转换后的数据
                  cell.value = transformRowData(row, standardHeader);

                  // 设置金额列的单元格格式为货币格式
                  if (standardHeader === "票价" || standardHeader === "燃油附加费" || standardHeader === "民航发展基金" ||
                      standardHeader === "保险费" || standardHeader === "改签费" || standardHeader === "退票费" ||
                      standardHeader === "小计" || standardHeader === "保险" || standardHeader === "服务费" ||
                      standardHeader === "实收" || standardHeader === "机票计税价格（票价+燃油附加费）" || standardHeader === "机票增值税" ||
                      standardHeader === "机票不含税金额" || standardHeader === "WD上填列Airfare数" || standardHeader === "代理商服务费增值税" ||
                      standardHeader === "代理商不含税服务金额" || standardHeader === "机票增值税+服务费税额" || standardHeader === "Airfare+服务费不含税" ||
                      standardHeader === "Checking") {
                    // 机票计税价格使用公式：L列+M列
                    if (standardHeader === "机票计税价格（票价+燃油附加费）") {
                      cell.value = {
                        formula: `L${actualRowIndex + 1}+M${actualRowIndex + 1}`,
                        result: 0
                      };
                    } else if (standardHeader === "机票增值税") {
                      // 机票增值税公式：=IF(OR(E3="",I3<>"国内"),0,ROUND(L3/1.09*0.09,2)+ROUND(M3/1.09*0.09,2))
                      cell.value = {
                        formula: `IF(OR(E${actualRowIndex + 1}="",I${actualRowIndex + 1}<>"国内"),0,ROUND(L${actualRowIndex + 1}/1.09*0.09,2)+ROUND(M${actualRowIndex + 1}/1.09*0.09,2))`,
                        result: 0
                      };
                          // 设置蓝色背景
                      cell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FF00B0F0' } // #00B0F0 蓝色背景
                      } as any;
                    } else if (standardHeader === "机票不含税金额") {
                      // 机票不含税金额公式：=Y3-Z3
                      cell.value = {
                        formula: `Y${actualRowIndex + 1}-Z${actualRowIndex + 1}`,
                        result: 0
                      };
                    } else if (standardHeader === "WD上填列Airfare数") {
                      // WD上填列Airfare数公式：=AA3+N3+O3+Q3
                      cell.value = {
                        formula: `AA${actualRowIndex + 1}+N${actualRowIndex + 1}+O${actualRowIndex + 1}+Q${actualRowIndex + 1}`,
                        result: 0
                      };
                    } else if (standardHeader === "代理商服务费增值税") {
                      // 代理商服务费增值税公式：=ROUND(T3/1.06*0.06,2)
                      cell.value = {
                        formula: `ROUND(T${actualRowIndex + 1}/1.06*0.06,2)`,
                        result: 0
                      };
                    } else if (standardHeader === "代理商不含税服务金额") {
                      // 代理商不含税服务金额公式：=T3-AC3
                      cell.value = {
                        formula: `T${actualRowIndex + 1}-AC${actualRowIndex + 1}`,
                        result: 0
                      };
                    } else if (standardHeader === "机票增值税+服务费税额") {
                      // 机票增值税+服务费税额公式：=Z3+AC3
                      cell.value = {
                        formula: `Z${actualRowIndex + 1}+AC${actualRowIndex + 1}`,
                        result: 0
                      };
                    } else if (standardHeader === "Airfare+服务费不含税") {
                      // Airfare+服务费不含税公式：=AB3+AD3
                      cell.value = {
                        formula: `AB${actualRowIndex + 1}+AD${actualRowIndex + 1}`,
                        result: 0
                      };
                    } else if (standardHeader === "Checking") {
                      // Checking公式：=W3-Z3-AB3-AC3-AD3
                      cell.value = {
                        formula: `W${actualRowIndex + 1}-Z${actualRowIndex + 1}-AB${actualRowIndex + 1}-AC${actualRowIndex + 1}-AD${actualRowIndex + 1}`,
                        result: 0
                      };
                    } else {
                      // 将值转换为数字并设置货币格式，空值赋值为0
                      const numValue = parseFloat(String(cell.value || '').replace(/,/g, ''));
                      if (!isNaN(numValue)) {
                        cell.value = numValue;
                      } else {
                        cell.value = 0; // 空值或无效值赋值为0
                      }
                    }
                    cell.numFmt = '#,##0.00'; // 设置Excel货币格式，带千分位和两位小数
                  }
                }

                // PDF数据集成：使用PDF提取的数据匹配Excel中的电子客票号
                console.log(`  🔍 列处理: "${standardHeader}"`);
                if (standardHeader === "印刷序号(发票号码)") {
                  console.log(`  🎯 找到印刷序列! 开始PDF匹配调试`);
                  console.log(`    PDF数据总数: ${pdfData.value.length}`);
                  console.log(`    PDF数据内容:`, pdfData.value);


                  // 获取当前行的电子客票号（E列）
                  const ticketNumberIndex = columnMapping["电子客票号"];
                  console.log(`    电子客票号列索引: ${ticketNumberIndex}`);
                  console.log(`    列映射:`, columnMapping);

                  if (ticketNumberIndex !== undefined) {
                    const currentTicketNumber = String(row[ticketNumberIndex] || '').trim();
                    console.log(`    Excel电子客票号: "${currentTicketNumber}"`);
                    console.log(`    当前行数据:`, row);

                    if (currentTicketNumber && pdfData.value.length > 0) {
                      console.log(`    ✅ 条件满足，开始匹配PDF数据...`);

                      // 遍历所有PDF数据，查找匹配的记录
                      for (let i = 0; i < pdfData.value.length; i++) {
                        const pdfRecord = pdfData.value[i];
                        // 预处理：去掉电子客票号中的"-"符号后再进行比较
                        const normalizedCurrentTicketNumber = currentTicketNumber.replace(/-/g, '');
                        const normalizedPdfTicketNumber = pdfRecord.ticketNumber ? pdfRecord.ticketNumber.replace(/-/g, '') : '';
                        const normalizedOriginalValue = pdfRecord.originalValue ? pdfRecord.originalValue.replace(/-/g, '') : '';

                        console.log(`    检查PDF记录 ${i + 1}:`, {
                          ticketNumber: pdfRecord.ticketNumber,
                          invoiceNumber: pdfRecord.invoiceNumber,
                          originalValue: pdfRecord.originalValue,
                          currentTicketNumber: currentTicketNumber,
                          normalizedCurrentTicketNumber: normalizedCurrentTicketNumber,
                          normalizedPdfTicketNumber: normalizedPdfTicketNumber,
                          normalizedOriginalValue: normalizedOriginalValue
                        });

                        // 使用多种匹配方式，都基于去除"-"符号后的值
                        const isMatch =
                          (normalizedPdfTicketNumber && normalizedPdfTicketNumber === normalizedCurrentTicketNumber) ||
                          (normalizedOriginalValue && normalizedOriginalValue === normalizedCurrentTicketNumber)

                        console.log(`    匹配结果 ${i + 1}: ${isMatch}`);

                        if (isMatch) {
                          // 优先使用invoiceNumber，如果没有则使用originalValue
                          cell.value = pdfRecord.invoiceNumber || pdfRecord.originalValue;
                          console.log(`  🎉 PDF匹配成功! D列"印刷序号(发票号码)" = "${cell.value}"`);
                          console.log(`  📄 Excel电子客票号: "${currentTicketNumber}"`);
                          console.log(`  📄 打印匹配的PDF记录 ${i + 1}:`);
                          console.log(`     ticketNumber: ${pdfRecord.ticketNumber}`);
                          console.log(`     invoiceNumber: ${pdfRecord.invoiceNumber}`);
                          console.log(`     originalValue: ${pdfRecord.originalValue}`);
                          console.log(`     remark: ${pdfRecord.remark}`);
                          console.log(`     pageNum: ${pdfRecord.pageNum}`);
                          console.log(`     confidence: ${pdfRecord.confidence}`);
                          break; // 找到第一个匹配就停止
                        }
                      }

                      if (!cell.value || (typeof cell.value === 'string' && cell.value.startsWith("TEST_D_COLUMN_"))) {
                        console.log(`  ❌ PDF匹配失败: 未找到匹配的记录`);
                        console.log(`  📄 所有PDF记录详情:`);
                        pdfData.value.forEach((record, index) => {
                          console.log(`    记录 ${index + 1}:`, record);
                        });
                      }
                    } else {
                      console.log(`  ⚠️ PDF匹配条件不满足: currentTicketNumber="${currentTicketNumber}", pdfData.length=${pdfData.value.length}`);
                    }
                  } else {
                    console.log(`  ❌ 未找到电子客票号列映射`);
                  }
                } else if (standardHeader === "备注" && pdfData.value.length > 0) {
                  // 获取当前行的电子客票号（E列）
                  const ticketNumberIndex = columnMapping["电子客票号"];
                  if (ticketNumberIndex !== undefined) {
                    const currentTicketNumber = String(row[ticketNumberIndex] || '').trim();

                    if (currentTicketNumber) {
                      // 预处理：去掉电子客票号中的"-"符号后再进行比较
                      const normalizedCurrentTicketNumber = currentTicketNumber.replace(/-/g, '');

                      // 查找匹配的PDF记录
                      for (const pdfRecord of pdfData.value) {
                        const normalizedPdfTicketNumber = pdfRecord.ticketNumber ? pdfRecord.ticketNumber.replace(/-/g, '') : '';
                        const normalizedOriginalValue = pdfRecord.originalValue ? pdfRecord.originalValue.replace(/-/g, '') : '';

                        // 使用多种匹配方式，都基于去除"-"符号后的值
                        const isMatch =
                          (normalizedPdfTicketNumber && normalizedPdfTicketNumber === normalizedCurrentTicketNumber) ||
                          (normalizedOriginalValue && normalizedOriginalValue === normalizedCurrentTicketNumber) ||
                          (normalizedPdfTicketNumber && normalizedCurrentTicketNumber.includes(normalizedPdfTicketNumber)) ||
                          (normalizedPdfTicketNumber && normalizedPdfTicketNumber.includes(currentTicketNumber.split('-')[1] ? currentTicketNumber.split('-')[1] : ''))

                        if (isMatch) {
                          // 如果PDF数据有匹配，填写"电子行程单"
                          cell.value = "电子行程单";
                          console.log(`  📄 PDF备注匹配成功: 电子客票号"${currentTicketNumber}" -> 备注"${cell.value}"`);
                          console.log(`  📄 匹配的PDF记录: ticketNumber=${pdfRecord.ticketNumber}, invoiceNumber=${pdfRecord.invoiceNumber}`);
                          break;
                        }
                      }
                    }
                  }
                }

                cell.border = {
                  top: { style: "thin" },
                  bottom: { style: "thin" },
                  left: { style: "thin" },
                  right: { style: "thin" }
                };
                cell.alignment = {
                  horizontal: "center",
                  vertical: "middle"
                };
              });

              // 设置数据行高为24磅
              worksheet.getRow(actualRowIndex).height = 24;
            });

            // 为每个部门添加求和行
            {
              console.log(`在部门"${department}"后添加求和行`);
              const sumRowIndex = worksheet.rowCount + 1;

              // 计算该部门数据在Excel中的起始行和结束行
              // 注意：由于之后会插入标题行，实际数据会下移1位，所以这里+1
              const departmentStartRow = sumRowIndex - departmentRows.length + 1;
              const departmentEndRow = sumRowIndex - 1 + 1;

              console.log(`  部门"${department}"求和行调试: sumRowIndex=${sumRowIndex}, departmentRows.length=${departmentRows.length}, departmentStartRow=${departmentStartRow}, departmentEndRow=${departmentEndRow}`);

              standardHeaders.forEach((standardHeader, colIndex) => {
                const cell = worksheet.getCell(sumRowIndex, colIndex + 1);

                // 找到对应的Excel列字母（支持A-Z和AA-AZ等）
                let columnLetter: string;
                if (colIndex < 26) {
                  columnLetter = String.fromCharCode(65 + colIndex); // A, B, C, ..., Z
                } else {
                  // AA, AB, AC, ...
                  const firstLetter = String.fromCharCode(65 + Math.floor(colIndex / 26) - 1);
                  const secondLetter = String.fromCharCode(65 + (colIndex % 26));
                  columnLetter = firstLetter + secondLetter;
                }

                // 处理特定位置的列：O(14), P(15), Q(16), R(17)
                const isSpecialColumn = colIndex === 14 || colIndex === 15 || colIndex === 16 || colIndex === 17 || colIndex === 18;

                if (standardHeader === "序号") {
                  cell.value = ''; // 序号列留空，不显示"合计"
                } else if (standardHeader === "票价" || standardHeader === "燃油附加费" || standardHeader === "民航发展基金" ||
                          standardHeader === "保险" || standardHeader === "服务费" || standardHeader === "实收" ||
                          standardHeader === "改签费" || standardHeader === "退票费" || standardHeader === "机票计税价格（票价+燃油附加费）" ||
                          standardHeader === "机票增值税" || standardHeader === "机票不含税金额" || standardHeader === "WD上填列Airfare数" ||
                          standardHeader === "代理商服务费增值税" || standardHeader === "代理商不含税服务金额" ||
                          standardHeader === "机票增值税+服务费税额" || standardHeader === "Airfare+服务费不含税" ||
                          standardHeader === "Checking") {
                  // 设置求和公式，包括机票计税价格列
                  // 例如 =SUM(L2:L4)
                  cell.value = {
                    formula: `SUM(${columnLetter}${departmentStartRow}:${columnLetter}${departmentEndRow})`,
                    result: 0
                  };
                  cell.numFmt = '#,##0.00'; // 设置货币格式
                  cell.font = { bold: true };
                  console.log(`  设置求和公式: ${columnLetter}${departmentStartRow}:${columnLetter}${departmentEndRow}`);
                } else if (isSpecialColumn) {
                  // O(14), P(15), Q(16), R(17), S(18)列设置为0
                  cell.value = 0;
                  cell.numFmt = '#,##0.00';
                  cell.font = { bold: true };
                  console.log(`  设置固定值0: 列${colIndex + 1}(${String.fromCharCode(65 + colIndex)})`);
                } else {
                  cell.value = ''; // 其他列为空
                }

                // 设置求和行的样式
                cell.border = {
                  top: { style: "thin" },
                  bottom: { style: "thin" }, // 单实线底部边框
                  left: { style: "thin" },
                  right: { style: "thin" }
                };
                cell.alignment = {
                  horizontal: "center",
                  vertical: "middle"
                };
                cell.fill = {
                  type: 'pattern',
                  pattern: 'solid',
                  fgColor: { argb: 'FFFFFF00' } // 黄色背景
                } as any;
              });

              // 设置求和行高为24磅
              worksheet.getRow(sumRowIndex).height = 24;

              // 记录求和行行号
              departmentSumRows.set(department, sumRowIndex);
            }
          });

          // 添加总计行（对所有部门求和行的求和）
          if (departmentSumRows.size > 0) {
            console.log(`在工作表 ${originalSheetKey} 添加总计行，汇总 ${departmentSumRows.size} 个部门`);
            const grandTotalRowIndex = worksheet.rowCount + 1;

            standardHeaders.forEach((standardHeader, colIndex) => {
              const cell = worksheet.getCell(grandTotalRowIndex, colIndex + 1);

              // 找到对应的Excel列字母（支持A-Z和AA-AZ等）
              let columnLetter: string;
              if (colIndex < 26) {
                columnLetter = String.fromCharCode(65 + colIndex); // A, B, C, ..., Z
              } else {
                // AA, AB, AC, ...
                const firstLetter = String.fromCharCode(65 + Math.floor(colIndex / 26) - 1);
                const secondLetter = String.fromCharCode(65 + (colIndex % 26));
                columnLetter = firstLetter + secondLetter;
              }

              // 处理特定位置的列：O(14), P(15), Q(16), R(17)
              const isSpecialColumn = colIndex === 14 || colIndex === 15 || colIndex === 16 || colIndex === 17 || colIndex === 18;

              if (colIndex === 1) {
                // 出票日期列显示"总计"
                cell.value = "";
                cell.alignment = { horizontal: "center", vertical: "middle" };
              } else if (standardHeader === "票价" || standardHeader === "燃油附加费" || standardHeader === "民航发展基金" ||
                        standardHeader === "保险" || standardHeader === "服务费" || standardHeader === "实收" ||
                        standardHeader === "改签费" || standardHeader === "退票费" || standardHeader === "机票计税价格（票价+燃油附加费）" ||
                        standardHeader === "机票增值税" || standardHeader === "机票不含税金额" || standardHeader === "WD上填列Airfare数" ||
                        standardHeader === "代理商服务费增值税" || standardHeader === "代理商不含税服务金额" ||
                        standardHeader === "机票增值税+服务费税额" || standardHeader === "Airfare+服务费不含税" ||
                        standardHeader === "Checking") {
                // 创建对所有部门求和行的求和公式，格式类似：=SUM(L24+L20+L31)
                const sumRowIndices = Array.from(departmentSumRows.values());
                const cellReferences = sumRowIndices.map(rowIndex => `${columnLetter}${rowIndex}`);
                const sumFormula = cellReferences.join('+');

                console.log(`  总计行公式调试: 部门求和行=${sumRowIndices.join(', ')}, 公式=${sumFormula}`);

                cell.value = {
                  formula: `SUM(${sumFormula})`,
                  result: 0
                };
                cell.numFmt = '#,##0.00';
                cell.font = { bold: true };
                console.log(`  总计行设置公式: SUM(${sumFormula}) for ${standardHeader}`);
              } else if (isSpecialColumn) {
                // O(14), P(15), Q(16), R(17), S(18)列设置为0
                cell.value = 0;
                cell.numFmt = '#,##0.00';
                cell.font = { bold: true };
                console.log(`  总计行设置固定值0: 列${colIndex + 1}(${columnLetter})`);
              } else {
                cell.value = null;
              }

              // 设置总计行的样式
              cell.border = {
                top: { style: "thin" }, // 单线顶部边框
                bottom: { style: "thin" },
                left: { style: "thin" },
                right: { style: "thin" }
              };
              cell.alignment = {
                horizontal: "center",
                vertical: "middle"
              };
              cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FF84BC49' } // 浅绿色背景
              } as any;
            });

            // 设置总计行高为24磅
            worksheet.getRow(grandTotalRowIndex).height = 24;
          }

          console.log(`  工作表 ${originalSheetKey}: 添加 ${companyData.length} 行数据，使用标准表头 ${standardHeaders.length} 列`);
        }
      });

      // 如果没有数据，删除这个工作表
      if (!hasData) {
        console.log(`公司 ${companyGroup.groupName} 没有数据，删除工作表`);
        const sheetIndex = newWorkbook.worksheets.findIndex(ws => ws.name === companyGroup.groupName);
        if (sheetIndex !== -1) {
          newWorkbook.removeWorksheet(sheetIndex + 1);
        }
      } else {
        // 隐藏指定位置的列：O(14), P(15), Q(16), R(17)
        const columnsToHide = [14, 15, 16, 17, 18]; // 对应O, P, Q, R, S列
        columnsToHide.forEach((colIndex) => {
          const column = worksheet.getColumn(colIndex + 1);
          column.hidden = true;
          const columnName = String.fromCharCode(65 + colIndex); // A=0, B=1, ..., O=14, S=18
          console.log(`  隐藏列: ${columnName} (第${colIndex + 1}列)`);
        });

        // 自动调整列宽（更紧凑）
        worksheet.columns.forEach((column) => {
          let maxLength = 0;

          column.eachCell((cell, rowNumber) => {
            if (cell.value) {
              const text = cell.value.toString();

              // 特殊处理需要换行的列表头
              const wrapTextHeaders = [
                "机票计税价格（票价+燃油附加费）",
                "机票不含税金额",
                "WD上填列Airfare数",
                "代理商服务费增值税",
                "代理商不含税服务金额",
                "机票增值税+服务费税额",
                "Airfare+服务费不含税"
              ];

              if (rowNumber === 1 && wrapTextHeaders.includes(text)) {
                // 特殊处理机票计税价格列，宽度增加2
                if (text === "机票计税价格（票价+燃油附加费）") {
                  column.width = 18; // 从16增加到18
                } else {
                  column.width = 16; // 其他需要换行的列保持16
                }
                // 设置表头文字自动换行
                cell.alignment = {
                  horizontal: "center",
                  vertical: "middle",
                  wrapText: true // 启用文字自动换行
                };
                console.log(`  列 ${column.letter} ("${text}") 宽度设置为: ${column.width}，启用文字换行`);
                return; // 跳过该列的自动宽度计算
              }

              const charWidth = text.split("").reduce((width, char) => {
                return width + (char.charCodeAt(0) > 127 ? 2 : 1);
              }, 0);
              if (charWidth > maxLength) {
                maxLength = charWidth;
              }
            }
          });

          // 只有当列宽没有被特殊设置时才进行自动调整，使用更紧凑的宽度
          if (column.width !== 16 && column.width !== 12 && column.width !== 14 && column.width !== 10 && column.width !== 8 && column.width !== 6 && column.width !== 18 && column.width !== 3.7) {
            column.width = Math.max(maxLength * 0.8, 10); // 从1.1改为0.8，从15改为10，更紧凑
          }

          // 特殊处理出票日期、电子客票号、乘机日期、印刷序号列，设置更大的宽度
          const columnIndex = column.number - 1; // 列索引（从0开始）
          if (columnIndex === 1 || columnIndex === 3 || columnIndex === 4 || columnIndex === 7) { // 出票日期(1)、印刷序号(3)、电子客票号(4)、乘机日期(7)
            let minWidth = 18;
            let columnName = '';

            if (columnIndex === 1) {
              columnName = '出票日期';
              minWidth = 14; // 出票日期设置为14
            } else if (columnIndex === 3) {
              columnName = '印刷序号(发票号码)';
              minWidth = 22; // 印刷序号设置为22
            } else if (columnIndex === 4) {
              columnName = '电子客票号';
              minWidth = 18; // 电子客票号保持18
            } else if (columnIndex === 7) {
              columnName = '乘机日期';
              minWidth = 14; // 乘机日期设置为14
            }

            if (column.width < minWidth) {
              column.width = minWidth;
              console.log(`  列 ${column.letter} (${columnName}) 宽度设置为: ${minWidth}`);
            }
          }

          // 特殊处理承运人列，设置更小的宽度
          if (columnIndex === 2) { // 承运人列（第2列，C列）
            column.width = 8;
            console.log(`  列 ${column.letter} (承运人) 宽度设置为: 9 (紧凑宽度)`);
          }

          // 特殊处理国际/国内列，设置较小的宽度
          if (columnIndex === 8) { // 国际/国内列（第8列，I列）
            column.width = 10;
            console.log(`  列 ${column.letter} (国际/国内) 宽度设置为: 10 (紧凑宽度)`);
          }

          // 特殊处理序号列，设置更小的宽度
          if (columnIndex === 0) { // 序号列（第0列，A列）
            column.width = 5; // 更精确的设置，尝试接近Excel中的4.25字符
            console.log(`  列 ${column.letter} (序号) 宽度设置为: 3.7 (ExcelJS单位，精确调整)`);
          }

          // 特殊处理计算类列，设置更小的宽度
          if (columnIndex === 26 || columnIndex === 27 || columnIndex === 28 || columnIndex === 29) { // 代理商服务费增值税(26)、代理商不含税服务金额(27)、机票增值税+服务费税额(28)、Airfare+服务费不含税(29)
            column.width = 14;
            const columnNames = ['代理商服务费增值税', '代理商不含税服务金额', '机票增值税+服务费税额', 'Airfare+服务费不含税'];
            console.log(`  列 ${column.letter} (${columnNames[columnIndex - 26]}) 宽度设置为: 14 (紧凑宽度)`);
          }

          // 特殊处理备注列（X列），设置合适的宽度
          if (columnIndex === 23) { // 备注列（第23列，X列）
            column.width = 16; // 设置为16，适合显示"电子行程单"等内容
            console.log(`  列 ${column.letter} (备注) 宽度设置为: 16 (适合显示电子行程单)`);
          }
        });
      }

      // 在工作表处理完成后添加标题行（这样不会影响列宽计算）
      if (hasData && worksheet.rowCount > 0) {
        // 生成标题
        const currentDate = new Date();
        const currentYear = currentDate.getFullYear();
        const currentMonth = currentDate.getMonth();
        const lastMonth = currentMonth === 0 ? 12 : currentMonth;
        const lastMonthStr = lastMonth.toString().padStart(2, '0');

        const titleText = `${currentYear}年${lastMonthStr}月份深圳市特航航空服务有限公司与${companyGroup.groupName}机票结算表(830039)`;

        // 在现有数据前插入一行作为标题行（第1行），将所有现有数据下移一行
        worksheet.insertRow(1, []);

        // 更新所有记录的行号，因为插入了一行标题行
        console.log(`  📍 插入标题行前的部门求和行记录:`, Array.from(departmentSumRows.entries()).map(([dept, row]) => `${dept}=${row}`));

        const updatedDepartmentSumRows = new Map<string, number>();
        departmentSumRows.forEach((rowIndex, department) => {
          updatedDepartmentSumRows.set(department, rowIndex + 1);
          console.log(`    🔄 更新 ${department}: ${rowIndex} → ${rowIndex + 1}`);
        });

        // 更新原始Map
        departmentSumRows.clear();
        updatedDepartmentSumRows.forEach((rowIndex, department) => {
          departmentSumRows.set(department, rowIndex);
        });

        console.log(`  📍 插入标题行后的部门求和行记录:`, Array.from(departmentSumRows.entries()).map(([dept, row]) => `${dept}=${row}`));
        console.log(`  工作表 ${companyGroup.groupName}: 标题行插入后，更新了 ${departmentSumRows.size} 个部门求和行的行号`);

        // 更新总计行中的公式引用
        if (departmentSumRows.size > 0) {
          // 找到总计行的位置（应该是最后一个有数据的行，在标题行插入后）
          // 总计行是所有部门求和行之后的那一行
          const maxDepartmentSumRow = Math.max(...Array.from(departmentSumRows.values()));
          const grandTotalRowIndex = maxDepartmentSumRow + 1; // 总计行在最后一个部门求和行的下一行

          console.log(`  🔍 总计行位置计算:`);
          console.log(`    - 最后一个部门求和行位置: ${maxDepartmentSumRow}`);
          console.log(`    - 总计行位置: ${grandTotalRowIndex}`);
          console.log(`    - 工作表总行数: ${worksheet.rowCount}`);

          // 定义需要更新公式的列索引（对应standardHeaders中的索引）
          const formulaColumnIndices = [11, 12, 13, 14, 15, 16, 17, 18, 20, 21, 22, 24, 25, 26, 27, 28, 29, 30];
          const columnNames = ["票价", "燃油附加费", "民航发展基金", "保险费", "改签费", "退票费", "小计", "保险", "服务费", "改签费", "退票费", "实收", "机票计税价格（票价+燃油附加费）", "机票增值税", "机票不含税金额", "WD上填列Airfare数", "代理商服务费增值税", "代理商不含税服务金额"];

          console.log(`  部门求和行记录:`, Array.from(departmentSumRows.entries()).map(([dept, row]) => `${dept}=${row}`));

          formulaColumnIndices.forEach((colIndex, arrayIndex) => {
            const cell = worksheet.getCell(grandTotalRowIndex, colIndex + 1);

            // 找到对应的Excel列字母
            let columnLetter: string;
            if (colIndex < 26) {
              columnLetter = String.fromCharCode(65 + colIndex);
            } else {
              const firstLetter = String.fromCharCode(65 + Math.floor(colIndex / 26) - 1);
              const secondLetter = String.fromCharCode(65 + (colIndex % 26));
              columnLetter = firstLetter + secondLetter;
            }

            // 创建新的求和公式，使用更新后的部门求和行号
            const sumRowIndices = Array.from(departmentSumRows.values());
            const cellReferences = sumRowIndices.map(rowIndex => `${columnLetter}${rowIndex}`);
            const newFormula = cellReferences.join('+');

            cell.value = {
              formula: `SUM(${newFormula})`,
              result: 0
            };

            const columnName = columnNames[arrayIndex] || `未知列${colIndex}`;
            console.log(`    更新列 ${columnLetter} (${columnName}) 索引${colIndex} 公式: SUM(${newFormula})`);

            // 特别打印票价列的详细信息
            if (colIndex === 11) {
              console.log(`    🎫 票价列详细信息:`);
              console.log(`      - 总计行位置: ${grandTotalRowIndex}`);
              console.log(`      - 部门求和行位置: [${sumRowIndices.join(', ')}]`);
              console.log(`      - 生成公式: SUM(${newFormula})`);
              console.log(`      - 单元格地址: ${columnLetter}${grandTotalRowIndex}`);

              // 检查更新前后的公式
              const beforeValue = cell.value;
              console.log(`      - 更新前单元格值:`, beforeValue);
              console.log(`      - 更新后单元格值:`, cell.value);
            }
          });
        }

        // 合并标题行从A列到X列（第1-24列）
        worksheet.mergeCells(1, 1, 1, 24);
        const titleCell = worksheet.getCell(1, 1);
        titleCell.value = titleText;
        titleCell.font = {
          bold: true,
          size: 16
        };
        titleCell.alignment = {
          horizontal: "center",
          vertical: "middle"
        };

        worksheet.getRow(1).height = 40;
        console.log(`  工作表 ${companyGroup.groupName}: 已添加标题行，总行数: ${worksheet.rowCount}`);

        // 在标题行设置完成后，添加付款提示行
        if (departmentSumRows.size > 0) {
          const paymentReminderRowIndex = worksheet.rowCount + 1;
          const currentDate = new Date();
          const currentYear = currentDate.getFullYear();
          const currentMonth = currentDate.getMonth() + 1;
          const paymentDate = `${currentYear}年${currentMonth.toString().padStart(2, '0')}月02日`;

          const totalAmountColumnLetter = 'W';
          const maxDepartmentSumRow = Math.max(...Array.from(departmentSumRows.values()));

          // 按照标题行的方式：先合并，再设置内容和格式
          console.log(`  📝 在标题行后添加付款提示行: 第${paymentReminderRowIndex}行，第1-24列`);
          worksheet.mergeCells(paymentReminderRowIndex, 1, paymentReminderRowIndex, 24);

          const reminderCell = worksheet.getCell(paymentReminderRowIndex, 1);
          reminderCell.value = {
            formula: `CONCATENATE("总计：", TEXT(${totalAmountColumnLetter}${maxDepartmentSumRow + 1}, "0"), "元。请贵公司在${paymentDate}前结款，付款后请提供银行水单或致电联系查询款项是否到账，谢谢合作！")`,
            result: ''
          };

          reminderCell.font = { size: 12, bold: false };
          reminderCell.alignment = { horizontal: "left", vertical: "middle", wrapText: true };
          reminderCell.border = {
            top: { style: "thin" }, bottom: { style: "thin" },
            left: { style: "thin" }, right: { style: "thin" }
          };

          worksheet.getRow(paymentReminderRowIndex).height = 24;

          console.log(`  ✅ 付款提示行合并完成 (第${paymentReminderRowIndex}行)`);

          // 添加银行账户信息行
          const bankInfoRowIndex = worksheet.rowCount + 1;
          const bankInfoText = "开户行：光大银行(光大银行深圳八卦岭支行),账号：38980188000607612,名称：深圳市特航航空服务有限公司";

          // 按照标题行的方式：先合并，再设置内容和格式
          console.log(`  📝 添加银行信息行: 第${bankInfoRowIndex}行，第1-24列`);
          worksheet.mergeCells(bankInfoRowIndex, 1, bankInfoRowIndex, 24);

          const bankInfoCell = worksheet.getCell(bankInfoRowIndex, 1);
          bankInfoCell.value = bankInfoText;

          // 设置银行信息行格式
          bankInfoCell.font = {
            size: 12,
            bold: false,
            color: { argb: 'FFFF0000' } // 红色
          };
          bankInfoCell.alignment = {
            horizontal: "left",
            vertical: "middle",
            wrapText: true
          };
          bankInfoCell.border = {
            top: { style: "thin" },
            bottom: { style: "thin" },
            left: { style: "thin" },
            right: { style: "thin" }
          };

          // 设置银行信息行高为24磅
          worksheet.getRow(bankInfoRowIndex).height = 24;

          console.log(`  ✅ 银行信息行合并完成 (第${bankInfoRowIndex}行)`);

          // 添加制表人行
          const creatorRowIndex = worksheet.rowCount + 1;
          const creatorText = "制表人：王欣欣";

          // 按照标题行的方式：先合并，再设置内容和格式
          console.log(`  📝 添加制表人行: 第${creatorRowIndex}行，第1-24列`);
          worksheet.mergeCells(creatorRowIndex, 1, creatorRowIndex, 24);

          const creatorCell = worksheet.getCell(creatorRowIndex, 1);
          creatorCell.value = creatorText;

          // 设置制表人行格式
          creatorCell.font = {
            size: 12,
            bold: false
          };
          creatorCell.alignment = {
            horizontal: "right", // 文字靠右对齐
            vertical: "middle",
            wrapText: true
          };
          creatorCell.border = {
            top: { style: "thin" },
            bottom: { style: "thin" },
            left: { style: "thin" },
            right: { style: "thin" }
          };

          // 设置制表人行高为24磅
          worksheet.getRow(creatorRowIndex).height = 24;

          console.log(`  ✅ 制表人行合并完成 (第${creatorRowIndex}行)`);

          // 添加当前月份日期行
          const dateRowIndex = worksheet.rowCount + 1;

          // 获取当前月份的1号
          const today = new Date();
          const thisYear = today.getFullYear();
          const thisMonth = today.getMonth() + 1;
          const dateText = `${thisYear}/${thisMonth}/1`;

          // 按照标题行的方式：先合并，再设置内容和格式
          console.log(`  📝 添加日期行: 第${dateRowIndex}行，第1-24列，日期: ${dateText}`);
          worksheet.mergeCells(dateRowIndex, 1, dateRowIndex, 24);

          const dateCell = worksheet.getCell(dateRowIndex, 1);
          dateCell.value = dateText;

          // 设置日期行格式
          dateCell.font = {
            size: 12,
            bold: false
          };
          dateCell.alignment = {
            horizontal: "right", // 文字靠右对齐
            vertical: "middle",
            wrapText: true
          };
          dateCell.border = {
            top: { style: "thin" },
            bottom: { style: "thin" },
            left: { style: "thin" },
            right: { style: "thin" }
          };

          // 设置日期行高为24磅
          worksheet.getRow(dateRowIndex).height = 24;

          console.log(`  ✅ 日期行合并完成 (第${dateRowIndex}行)`);
        }
      }
    }

    // 生成一个包含所有公司工作表的Excel文件
    if (newWorkbook.worksheets.length > 0) {
      const excelBuffer = await newWorkbook.xlsx.writeBuffer();
      const blob = new Blob([excelBuffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      });

      const fileName = "戴德梁行账单拆分结果.xlsx";
      saveAs(blob, fileName);

      console.log(`成功生成Excel文件: ${fileName}，包含 ${newWorkbook.worksheets.length} 个工作表`);
      ElMessage.success(`成功生成Excel文件：${fileName}，包含 ${newWorkbook.worksheets.length} 个公司工作表！`);
    } else {
      ElMessage.warning("没有找到任何数据，无法生成Excel文件");
    }

  } catch (error) {
    console.error("生成Excel文件失败:", error);
    ElMessage.error("生成Excel文件失败");
  } finally {
    generating.value = false;
  }
};

const beforeUpload = (file: File) => {
  const isExcel = file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
                  file.type === 'application/vnd.ms-excel' ||
                  file.name.endsWith('.xlsx') ||
                  file.name.endsWith('.xls');

  if (!isExcel) {
    ElMessage.error('只能上传Excel文件！');
    return false;
  }

  const isLt10M = file.size / 1024 / 1024 < 10;
  if (!isLt10M) {
    ElMessage.error("文件大小不能超过10MB！");
    return false;
  }

  return true;
};

// 空请求函数，用于禁用默认上传行为
const noopRequest = () => Promise.resolve()

// ZIP文件处理函数 - 递归解压ZIP包中的PDF文件
const processZipFile = async (zipFile: File): Promise<File[]> => {
  console.log('开始处理ZIP文件:', zipFile.name)

  try {
    const zip = new JSZip()
    const zipData = await zip.loadAsync(zipFile)
    const pdfFiles: File[] = []

    // 递归函数，用于遍历ZIP包中的所有文件和文件夹
    const traverseZip = async (zipObj: any) => {
      for (const [relativePath, file] of Object.entries(zipObj.files)) {
        const zipEntry = file as any

        // 跳过目录
        if (zipEntry.dir) {
          console.log(`跳过目录: ${relativePath}`)
          continue
        }

        // 检查是否为PDF文件
        if (relativePath.toLowerCase().endsWith('.pdf')) {
          try {
            console.log(`找到PDF文件: ${relativePath}`)
            const pdfBlob = await zipEntry.async('blob')

            // 创建File对象，保持原始文件名
            const fileName = relativePath.split('/').pop() || `pdf_${Date.now()}.pdf`
            const pdfFile = new File([pdfBlob], fileName, {
              type: 'application/pdf'
            })

            pdfFiles.push(pdfFile)
            console.log(`成功提取PDF文件: ${fileName}`)
          } catch (error) {
            console.error(`提取PDF文件失败 ${relativePath}:`, error)
          }
        }
      }
    }

    await traverseZip(zipData)

    console.log(`ZIP文件处理完成，共提取 ${pdfFiles.length} 个PDF文件`)
    return pdfFiles

  } catch (error) {
    console.error('ZIP文件处理失败:', error)
    ElMessage.error(`ZIP文件 "${zipFile.name}" 处理失败，请检查文件是否损坏`)
    return []
  }
}

// PDF文件变化处理函数 - 支持PDF和ZIP文件
const handlePdfFileChange = async (file: any, fileList: any[]) => {
  console.log('文件变化:', file.name, fileList.length)

  // 验证文件
  if (!file.raw) {
    ElMessage.error('文件无效！')
    return
  }

  const fileName = file.raw.name.toLowerCase()
  const fileSize = file.raw.size / 1024 / 1024 // MB

  // 检查文件大小
  if (fileSize > 100) {
    ElMessage.error("文件大小不能超过100MB！")
    return
  }

  let filesToProcess: File[] = []

  if (fileName.endsWith('.zip')) {
    // 处理ZIP文件
    console.log('检测到ZIP文件，开始解压...')

    // 检查ZIP文件是否已经存在
    const zipExists = uploadedPdfFiles.value.some(existingFile =>
      existingFile.name === file.raw.name && existingFile.size === file.raw.size
    )

    if (zipExists) {
      ElMessage.warning(`ZIP文件 "${file.raw.name}" 已经存在，跳过重复上传`)
      return
    }

    try {
      const extractedFiles = await processZipFile(file.raw)

      if (extractedFiles.length === 0) {
        ElMessage.warning(`ZIP文件 "${file.raw.name}" 中未找到PDF文件`)
        return
      }

      filesToProcess = extractedFiles

      // 添加ZIP文件到记录
      uploadedPdfFiles.value.push(file.raw)

      ElMessage.success(`ZIP文件解压成功，共找到 ${extractedFiles.length} 个PDF文件`)

    } catch (error) {
      console.error('ZIP文件处理失败:', error)
      ElMessage.error(`处理ZIP文件 "${file.raw.name}" 失败`)
      return
    }

  } else if (fileName.endsWith('.pdf')) {
    // 处理单个PDF文件
    console.log('检测到PDF文件')

    // 检查PDF文件是否已经存在
    const fileExists = uploadedPdfFiles.value.some(existingFile =>
      existingFile.name === file.raw.name && existingFile.size === file.raw.size
    )

    if (fileExists) {
      ElMessage.warning(`PDF文件 "${file.raw.name}" 已经存在，跳过重复上传`)
      return
    }

    filesToProcess = [file.raw]

    // 添加PDF文件到记录
    uploadedPdfFiles.value.push(file.raw)

  } else {
    ElMessage.error('只支持上传PDF文件或ZIP压缩包！')
    return
  }

  // 批量处理所有PDF文件
  console.log(`开始批量处理 ${filesToProcess.length} 个PDF文件`)

  try {
    // 设置loading状态
    pdfProcessingCount.value += filesToProcess.length
    pdfLoading.value = true

    // 并发处理PDF文件以提高效率
    const processPromises = filesToProcess.map(async (pdfFile, index) => {
      try {
        console.log(`处理第 ${index + 1}/${filesToProcess.length} 个PDF文件: ${pdfFile.name}`)
        await processPdfFile(pdfFile)
      } catch (error) {
        console.error(`处理PDF文件 "${pdfFile.name}" 失败:`, error)
        // 不抛出错误，继续处理其他文件
      }
    })

    await Promise.all(processPromises)

    ElMessage.success(`批量处理完成，成功处理 ${filesToProcess.length} 个PDF文件`)

  } catch (error) {
    console.error('批量处理失败:', error)
    ElMessage.error('批量处理PDF文件失败')
  } finally {
    // 重置loading状态
    pdfProcessingCount.value -= filesToProcess.length
    if (pdfProcessingCount.value <= 0) {
      pdfLoading.value = false
      pdfProcessingCount.value = 0
    }
  }
}

// PDF处理函数（保持向后兼容）
const handlePdfUpload = async (file: File) => {
  if (!file.name.toLowerCase().endsWith('.pdf')) {
    ElMessage.error('只能上传PDF文件！');
    return false;
  }

  const isLt50M = file.size / 1024 / 1024 < 50;
  if (!isLt50M) {
    ElMessage.error("PDF文件大小不能超过50MB！");
    return false;
  }

  // 检查文件是否已经存在
  const fileExists = uploadedPdfFiles.value.some(existingFile =>
    existingFile.name === file.name && existingFile.size === file.size
  );

  if (fileExists) {
    ElMessage.warning(`文件 "${file.name}" 已经存在，跳过重复上传`);
    return false;
  }

  // 添加到文件列表
  uploadedPdfFiles.value.push(file);

  // 使用await确保文件按顺序处理，避免并发问题
  try {
    await processPdfFile(file);
  } catch (error) {
    console.error(`处理文件 "${file.name}" 失败:`, error);
    ElMessage.error(`处理文件 "${file.name}" 失败`);
  }

  return false; // 阻止自动上传
};

const handlePdfRemove = async (file: any, fileList: any[]) => {
  // 从文件列表中移除
  uploadedPdfFiles.value = fileList;

  // 重新处理剩余的PDF文件 - 不直接清空，而是重新处理所有剩余文件
  const remainingFiles = fileList.map(f => f.raw);

  if (remainingFiles.length > 0) {
    // 清空现有数据，然后重新处理所有剩余文件以确保数据一致性
    pdfData.value = [];

    // 重置处理计数器并设置loading状态
    pdfProcessingCount.value = 0;
    pdfLoading.value = true;

    try {
      // 逐个处理剩余文件
      await Promise.all(remainingFiles.map(f => processPdfFile(f)));
      ElMessage.success(`PDF文件已更新，移除"${file.name}"，当前总计${pdfData.value.length}条记录`);
    } catch (error) {
      console.error('重新处理PDF文件失败:', error);
      ElMessage.error('重新处理PDF文件失败');
    }
  } else {
    // 如果没有剩余文件，才清空数据
    pdfData.value = [];
    ElMessage.success('所有PDF文件已移除');
  }
};

// 配置PDF.js worker - 使用本地worker文件路径（与pdf.vue保持一致）
pdfjsLib.GlobalWorkerOptions.workerSrc = "/pdf.worker.min.mjs";

const processPdfFile = async (file: File) => {
  // 使用计数器来避免多个文件同时处理时loading状态混乱
  pdfProcessingCount.value++;
  pdfLoading.value = true;

  try {
    console.log('开始处理PDF文件:', file.name);

    // 将File转换为ArrayBuffer
    const arrayBuffer = await file.arrayBuffer();

    // 加载PDF文档，添加更多配置选项
    const loadingTask = pdfjsLib.getDocument({
      data: arrayBuffer,
      // 尝试使用标准配置，让pdfjs自己处理worker
    });

    const pdf = await loadingTask.promise;
    console.log(`PDF加载成功，共${pdf.numPages}页`);

    const extractedData: any[] = [];

    // 逐页处理PDF
    for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
      const page = await pdf.getPage(pageNum);
      const textContent = await page.getTextContent();

      // 提取并组合文本内容
      const pageText = textContent.items
        .map((item: any) => item.str)
        .join(' ');

      console.log(`第${pageNum}页文本长度:`, pageText.length);
      console.log(`=== 第${pageNum}页PDF完整文本内容 ===`);
      console.log('原始文本:', pageText);

      // 预处理文本：移除数字和字母之间的空格
      const cleanedText = pageText
        .replace(/(\d)\s+(?=\d)/g, '$1')  // 移除数字间的空格
        .replace(/([A-Z])\s+(?=[A-Z])/g, '$1')  // 移除字母间的空格
        .replace(/([A-Z])\s+(?=\d)/g, '$1')  // 移除字母数字间的空格
        .replace(/(\d)\s+(?=[A-Z])/g, '$1'); // 移除数字字母间的空格

      console.log('=== 清理后的文本 ===');
      console.log('清理后文本:', cleanedText);
      console.log('=== 文本内容结束 ===');

      // 使用简化的提取函数
      const pageData = extractInvoiceInfo(cleanedText, pageNum);
      extractedData.push(...pageData);
    }

    // 去重并排序
    console.log('🔍 PDF处理结果检查:');
    console.log('  extractedData:', extractedData);
    console.log('  extractedData.length:', extractedData.length);

    const uniqueData = removeDuplicates(extractedData);
    console.log('  uniqueData (去重后):', uniqueData);
    console.log('  uniqueData.length:', uniqueData.length);

    // 线程安全地合并新数据到现有数据
    // 使用响应式API确保数据更新的原子性
    const currentData = [...pdfData.value];
    const mergedData = removeDuplicates([...currentData, ...uniqueData]);

    // 原子性更新pdfData，避免并发问题
    pdfData.value = mergedData;
    console.log('✅ pdfData.value 已更新:', pdfData.value);
    console.log('✅ pdfData.value.length:', pdfData.value.length);

    console.log(`PDF处理完成，新增${uniqueData.length}条记录，总计${mergedData.length}条发票信息`);
    ElMessage.success(`PDF处理完成，文件"${file.name}"新增${uniqueData.length}条记录，总计${mergedData.length}条发票信息`);

  } catch (error: any) {
    console.error('PDF处理失败:', error);

    // 提供更具体的错误信息
    let errorMessage = 'PDF文件处理失败';
    if (error.message && error.message.includes('worker')) {
      errorMessage = 'PDF.js worker配置失败，请刷新页面重试';
    } else if (error.message && error.message.includes('Invalid PDF')) {
      errorMessage = '无效的PDF文件，请检查文件是否损坏';
    } else if (error.message && error.message.includes('password')) {
      errorMessage = 'PDF文件受密码保护，无法处理';
    } else if (error.message && error.message.includes('size')) {
      errorMessage = 'PDF文件过大，请选择较小的文件';
    }

    ElMessage.error(errorMessage);
  } finally {
    // 减少处理计数器
    pdfProcessingCount.value--;

    // 只有当所有文件都处理完成时才关闭loading
    if (pdfProcessingCount.value <= 0) {
      pdfLoading.value = false;
      pdfProcessingCount.value = 0; // 重置为0，避免负数
    }
  }
};

const calculateConfidence = (invoiceNumber: string, text: string): number => {
  if (!invoiceNumber) return 0;

  let confidence = 0.5; // 基础置信度

  // 长度合理性 (8-12位最佳)
  if (invoiceNumber.length >= 8 && invoiceNumber.length <= 12) {
    confidence += 0.2;
  }

  // 包含数字和字母的组合
  if (/\d/.test(invoiceNumber) && /[A-Za-z]/.test(invoiceNumber)) {
    confidence += 0.1;
  }

  // 纯数字且长度合理
  if (/^\d+$/.test(invoiceNumber) && invoiceNumber.length >= 8) {
    confidence += 0.15;
  }

  // 在文本中的位置和上下文
  const textLower = text.toLowerCase();
  const invoiceIndex = textLower.indexOf(invoiceNumber.toLowerCase());

  // 检查是否在关键词附近
  const keywords = ['印刷序号', '发票号码', '票据号', '票号', 'invoice'];
  const contextWindow = 50; // 上下文字符窗口

  for (const keyword of keywords) {
    const keywordIndex = textLower.indexOf(keyword);
    if (keywordIndex !== -1 && Math.abs(keywordIndex - invoiceIndex) <= contextWindow) {
      confidence += 0.2;
      break;
    }
  }

  return Math.min(confidence, 1.0); // 最大置信度为1.0
};

const removeDuplicates = (data: any[]) => {
  console.log('🔍 removeDuplicates 输入数据:', data);
  console.log('🔍 removeDuplicates 输入数据长度:', data.length);

  // 简化去重逻辑：基于ticketNumber+invoiceNumber组合去重
  const seen = new Set<string>();
  const uniqueData = data.filter(item => {
    const key = `${item.ticketNumber || ''}-${item.invoiceNumber || ''}`;
    console.log(`  检查项目: ticketNumber="${item.ticketNumber}", invoiceNumber="${item.invoiceNumber}"`);
    if (seen.has(key)) {
      console.log(`    ❌ 重复，跳过`);
      return false;
    }
    seen.add(key);
    console.log(`    ✅ 保留`);
    return true;
  });

  console.log('🔍 removeDuplicates 过滤后数据:', uniqueData);
  console.log('🔍 removeDuplicates 过滤后长度:', uniqueData.length);

  // 按页码排序
  const sortedData = uniqueData.sort((a, b) => a.pageNum - b.pageNum);
  console.log('🔍 removeDuplicates 最终结果:', sortedData);
  return sortedData;
};
</script>

<style scoped>
.bill-split-container {
  padding: 20px;
}

.upload-section {
  margin-bottom: 30px;
}

.upload-dragger {
  width: 100%;
}

.data-section {
  background: white;
  border-radius: 8px;
  box-shadow: 0 2px 12px rgba(0, 0, 0, 0.1);
  padding: 20px;
}

.data-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 20px;
}

.data-header h3 {
  margin: 0;
  color: #303133;
}

.data-summary {
  margin-bottom: 20px;
}

.company-name {
  cursor: pointer;
  padding: 8px 12px;
  border-radius: 4px;
  transition: all 0.3s ease;
  font-weight: 500;
}

.company-name:hover {
  background-color: #f0f9ff;
  color: #1890ff;
}

.company-name.selected {
  background-color: #1890ff;
  color: white;
}

.detail-table {
  margin-top: 20px;
  padding: 20px;
  background: #f8f9fa;
  border-radius: 8px;
  border: 1px solid #e9ecef;
}

/* PDF上传区域样式 */
.pdf-upload-section {
  margin: 20px 0;
}

.pdf-upload-card {
  border-radius: 8px;
  box-shadow: 0 2px 12px rgba(0, 0, 0, 0.1);
}

.pdf-uploader {
  width: 100%;
}

.pdf-uploader .el-upload-dragger {
  width: 100%;
  height: 120px;
  border: 2px dashed #d9d9d9;
  border-radius: 8px;
  background: #fafafa;
  transition: all 0.3s ease;
}

.pdf-uploader .el-upload-dragger:hover {
  border-color: #409eff;
  background: #f0f9ff;
}

.pdf-data-preview {
  margin-top: 20px;
  max-height: 400px;
  overflow: auto;
}

.more-data-hint {
  margin-top: 10px;
  padding: 8px 12px;
  background: #f0f9ff;
  border-left: 4px solid #409eff;
  color: #666;
  font-size: 14px;
}

.pdf-loading {
  text-align: center;
  padding: 40px 0;
}

.pdf-loading .el-icon {
  font-size: 24px;
  color: #409eff;
}

.pdf-loading p {
  margin-top: 10px;
  color: #666;
  font-size: 14px;
}

.card-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  font-weight: 600;
  color: #303133;
}
</style>
