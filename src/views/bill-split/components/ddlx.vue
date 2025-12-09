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
import { UploadFilled } from "@element-plus/icons-vue";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";

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
          if (standardHeader === "票面价" || standardHeader === "燃油" || standardHeader === "机建" ||
              standardHeader === "保险费" || standardHeader === "改签费" || standardHeader === "退票费" ||
              standardHeader === "小计" || standardHeader === "保险" || standardHeader === "系统使用费" ||
              standardHeader === "总金额" || standardHeader === "机票计税价格（票价+燃油附加费）" || standardHeader === "机票增值税" ||
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
          // 机票计税价格 = 票面价 + 燃油
          const ticketPriceIndex = columnMapping["票面价"];
          const fuelFeeIndex = columnMapping["燃油"];

          if (ticketPriceIndex !== undefined && fuelFeeIndex !== undefined) {
            const ticketPrice = parseFloat(String(originalRow[ticketPriceIndex] || '').replace(/,/g, '')) || 0;
            const fuelFee = parseFloat(String(originalRow[fuelFeeIndex] || '').replace(/,/g, '')) || 0;
            const taxPrice = ticketPrice + fuelFee;
            return taxPrice.toFixed(2);
          }
          return "0.00";
        } else if (standardHeader === "机票增值税") {
          // 机票增值税 = IF(OR(E3="",I3<>"国内"),0,ROUND(L3/1.09*0.09,2)+ROUND(M3/1.09*0.09,2))
          // E列是记账日期, I列是国际/国内, L列是票面价, M列是燃油
          const recordDateIndex = columnMapping["记账日期"];
          const domesticIndex = columnMapping["国际/国内"];
          const ticketPriceIndex = columnMapping["票面价"];
          const fuelFeeIndex = columnMapping["燃油"];

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
          // WD上填列Airfare数 = AA3+N3+O3+Q3 (机票不含税金额 + 票面价 + 燃油 + 保险费)
          const noTaxAmountIndex = columnMapping["机票不含税金额"];
          const ticketPriceIndex = columnMapping["票面价"];
          const fuelFeeIndex = columnMapping["燃油"];
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
          const totalAmountIndex = columnMapping["总金额"];
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
  const companyGroups = new Map<string, any>();

  Object.entries(allSheetData.value).forEach(([sheetKey, sheetData]) => {
    if (!sheetData || sheetData.length === 0) return;

    // 查找部门列
    const headers = sheetData[0] as any[];
    const departmentColumnIndex = headers.findIndex(
      (cell: any) => cell && cell.toString().includes("乘机人部门")
    );

    if (departmentColumnIndex === -1) return;

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

  return Array.from(companyGroups.values());
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
    "序号", "记账日期", "承运人", "印刷序号(发票号码)", "电子客票号",
    "乘机人", "部门", "乘机日期", "国际/国内", "航程", "航班号",
    "票面价", "燃油", "机建", "保险费", "改签费",
    "退票费", "小计", "保险", "系统使用费", "改签费", "退票费", "总金额", "备注", "机票计税价格（票价+燃油附加费）", "机票增值税", "机票不含税金额", "WD上填列Airfare数", "代理商服务费增值税", "代理商不含税服务金额", "机票增值税+服务费税额", "Airfare+服务费不含税", "Checking"
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
    } else if (headerText.includes("记账日期") || headerText.includes("出票日期")) {
      columnMapping["记账日期"] = index;
      console.log(`  -> 映射到"记账日期"`);
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
      columnMapping["航班号"] = index;
      console.log(`  -> 映射到"航班号"`);
    } else if (headerText.includes("票面价") || headerText.includes("票价")) {
      columnMapping["票面价"] = index;
      console.log(`  -> 映射到"票面价"`);
    } else if (headerText.includes("燃油附加费") || headerText.includes("燃油")) {
      columnMapping["燃油"] = index;
      console.log(`  -> 映射到"燃油"`);
    } else if (headerText.includes("民航发展基金") || headerText.includes("发展基金") || headerText.includes("基建费") || headerText.includes("机建")) {
      columnMapping["机建"] = index;
      console.log(`  -> 映射到"机建"`);
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
    } else if (headerText.includes("系统使用费") || headerText.includes("服务费")) {
      columnMapping["系统使用费"] = index;
      console.log(`  -> 映射到"系统使用费"`);
    } else if (headerText.includes("总金额") || headerText.includes("实收") || headerText.includes("实付") || headerText.includes("合计")) {
      columnMapping["总金额"] = index;
      console.log(`  -> 映射到"总金额"`);
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
  generating.value = true;
  const groupInfo = getGroupInfo();

  try {
    console.log(`开始生成分组Excel文件，共 ${groupInfo.length} 个公司`);

    // 创建一个工作簿，包含所有公司的工作表
    const newWorkbook = new ExcelJS.Workbook();

    // 为每个公司创建一个工作表
    for (const companyGroup of groupInfo) {
      console.log(`为公司 ${companyGroup.groupName} 创建工作表`);

      const worksheet = newWorkbook.addWorksheet(companyGroup.groupName, {
        views: [{ showGridLines: true }]
      });
      worksheet.properties.defaultRowHeight = 40;

      let hasData = false;

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
              cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFE6F3FF' }
              };
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
          const departmentSumRows: Map<string, number> = new Map(); // 记录每个部门的求和行行号

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
                } else if (colIndex === 14 || colIndex === 15 || colIndex === 16 || colIndex === 17) {
                  // O(14), P(15), Q(16), R(17)列设置为0
                  cell.value = 0;
                  cell.numFmt = '#,##0.00';
                } else {
                  // 使用转换函数获取转换后的数据
                  cell.value = transformRowData(row, standardHeader);

                  // 设置金额列的单元格格式为货币格式
                  if (standardHeader === "票面价" || standardHeader === "燃油" || standardHeader === "机建" ||
                      standardHeader === "保险费" || standardHeader === "改签费" || standardHeader === "退票费" ||
                      standardHeader === "小计" || standardHeader === "保险" || standardHeader === "系统使用费" ||
                      standardHeader === "总金额" || standardHeader === "机票计税价格（票价+燃油附加费）" || standardHeader === "机票增值税" ||
                      standardHeader === "机票不含税金额" || standardHeader === "WD上填列Airfare数" || standardHeader === "代理商服务费增值税" ||
                      standardHeader === "代理商不含税服务金额" || standardHeader === "机票增值税+服务费税额" || standardHeader === "Airfare+服务费不含税" ||
                      standardHeader === "Checking") {
                    // 机票计税价格使用公式：L列+M列
                    if (standardHeader === "机票计税价格（票价+燃油附加费）") {
                      cell.value = {
                        formula: `L${actualRowIndex}+M${actualRowIndex}`,
                        result: 0
                      };
                    } else if (standardHeader === "机票增值税") {
                      // 机票增值税公式：=IF(OR(E3="",I3<>"国内"),0,ROUND(L3/1.09*0.09,2)+ROUND(M3/1.09*0.09,2))
                      cell.value = {
                        formula: `IF(OR(E${actualRowIndex}="",I${actualRowIndex}<>"国内"),0,ROUND(L${actualRowIndex}/1.09*0.09,2)+ROUND(M${actualRowIndex}/1.09*0.09,2))`,
                        result: 0
                      };
                      // 设置浅蓝色背景
                      cell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FF019FD9' } // 浅蓝色背景
                      } as any;
                    } else if (standardHeader === "机票不含税金额") {
                      // 机票不含税金额公式：=Y3-Z3
                      cell.value = {
                        formula: `Y${actualRowIndex}-Z${actualRowIndex}`,
                        result: 0
                      };
                    } else if (standardHeader === "WD上填列Airfare数") {
                      // WD上填列Airfare数公式：=AA3+N3+O3+Q3
                      cell.value = {
                        formula: `AA${actualRowIndex}+N${actualRowIndex}+O${actualRowIndex}+Q${actualRowIndex}`,
                        result: 0
                      };
                    } else if (standardHeader === "代理商服务费增值税") {
                      // 代理商服务费增值税公式：=ROUND(T3/1.06*0.06,2)
                      cell.value = {
                        formula: `ROUND(T${actualRowIndex}/1.06*0.06,2)`,
                        result: 0
                      };
                    } else if (standardHeader === "代理商不含税服务金额") {
                      // 代理商不含税服务金额公式：=T3-AC3
                      cell.value = {
                        formula: `T${actualRowIndex}-AC${actualRowIndex}`,
                        result: 0
                      };
                    } else if (standardHeader === "机票增值税+服务费税额") {
                      // 机票增值税+服务费税额公式：=Z3+AC3
                      cell.value = {
                        formula: `Z${actualRowIndex}+AC${actualRowIndex}`,
                        result: 0
                      };
                    } else if (standardHeader === "Airfare+服务费不含税") {
                      // Airfare+服务费不含税公式：=AB3+AD3
                      cell.value = {
                        formula: `AB${actualRowIndex}+AD${actualRowIndex}`,
                        result: 0
                      };
                    } else if (standardHeader === "Checking") {
                      // Checking公式：=W3-Z3-AB3-AC3-AD3
                      cell.value = {
                        formula: `W${actualRowIndex}-Z${actualRowIndex}-AB${actualRowIndex}-AC${actualRowIndex}-AD${actualRowIndex}`,
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
              const departmentStartRow = sumRowIndex - departmentRows.length;
              const departmentEndRow = sumRowIndex - 1;

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
                const isSpecialColumn = colIndex === 14 || colIndex === 15 || colIndex === 16 || colIndex === 17;

                if (standardHeader === "序号") {
                  cell.value = ''; // 序号列留空，不显示"合计"
                } else if (standardHeader === "票面价" || standardHeader === "燃油" || standardHeader === "机建" ||
                          standardHeader === "保险" || standardHeader === "系统使用费" || standardHeader === "总金额" ||
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
                  // O(14), P(15), Q(16), R(17)列设置为0
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
              const isSpecialColumn = colIndex === 14 || colIndex === 15 || colIndex === 16 || colIndex === 17;

              if (colIndex === 1) {
                // 记账日期列显示"总计"
                cell.value = "";
                cell.alignment = { horizontal: "center", vertical: "middle" };
              } else if (standardHeader === "票面价" || standardHeader === "燃油" || standardHeader === "机建" ||
                        standardHeader === "保险" || standardHeader === "系统使用费" || standardHeader === "总金额" ||
                        standardHeader === "改签费" || standardHeader === "退票费" || standardHeader === "机票计税价格（票价+燃油附加费）" ||
                        standardHeader === "机票增值税" || standardHeader === "机票不含税金额" || standardHeader === "WD上填列Airfare数" ||
                        standardHeader === "代理商服务费增值税" || standardHeader === "代理商不含税服务金额" ||
                        standardHeader === "机票增值税+服务费税额" || standardHeader === "Airfare+服务费不含税" ||
                        standardHeader === "Checking") {
                // 创建对所有部门求和行的求和公式，格式类似：=SUM(L24+L20+L31)
                const sumRowIndices = Array.from(departmentSumRows.values());
                const cellReferences = sumRowIndices.map(rowIndex => `${columnLetter}${rowIndex}`);
                const sumFormula = cellReferences.join('+');

                cell.value = {
                  formula: `SUM(${sumFormula})`,
                  result: 0
                };
                cell.numFmt = '#,##0.00';
                cell.font = { bold: true };
                console.log(`  总计行设置公式: SUM(${sumFormula}) for ${standardHeader}`);
              } else if (isSpecialColumn) {
                // O(14), P(15), Q(16), R(17)列设置为0
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
        const columnsToHide = [14, 15, 16, 17]; // 对应O, P, Q, R列
        columnsToHide.forEach((colIndex) => {
          const column = worksheet.getColumn(colIndex + 1);
          column.hidden = true;
          const columnName = String.fromCharCode(65 + colIndex); // A=0, B=1, ..., O=14
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
          if (column.width !== 16 && column.width !== 12 && column.width !== 20 && column.width !== 14 && column.width !== 10 && column.width !== 8 && column.width !== 6 && column.width !== 18) {
            column.width = Math.max(maxLength * 0.8, 10); // 从1.1改为0.8，从15改为10，更紧凑
          }

          // 特殊处理记账日期、电子客票号、乘机日期、印刷序号列，设置更大的宽度
          const columnIndex = column.number - 1; // 列索引（从0开始）
          if (columnIndex === 1 || columnIndex === 3 || columnIndex === 4 || columnIndex === 7) { // 记账日期(1)、印刷序号(3)、电子客票号(4)、乘机日期(7)
            let minWidth = 18;
            let columnName = '';

            if (columnIndex === 1) {
              columnName = '记账日期';
              minWidth = 14; // 记账日期设置为14
            } else if (columnIndex === 3) {
              columnName = '印刷序号(发票号码)';
              minWidth = 20; // 印刷序号设置为20
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
            column.width = 6;
            console.log(`  列 ${column.letter} (序号) 宽度设置为: 6 (最紧凑宽度)`);
          }

          // 特殊处理计算类列，设置更小的宽度
          if (columnIndex === 26 || columnIndex === 27 || columnIndex === 28 || columnIndex === 29) { // 代理商服务费增值税(26)、代理商不含税服务金额(27)、机票增值税+服务费税额(28)、Airfare+服务费不含税(29)
            column.width = 14;
            const columnNames = ['代理商服务费增值税', '代理商不含税服务金额', '机票增值税+服务费税额', 'Airfare+服务费不含税'];
            console.log(`  列 ${column.letter} (${columnNames[columnIndex - 26]}) 宽度设置为: 14 (紧凑宽度)`);
          }
        });
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
</style>
