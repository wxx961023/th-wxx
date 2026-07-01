<script setup lang="ts">
import { computed, h, onMounted, ref } from "vue";
import dayjs from "dayjs";
import { Search, Refresh, Download } from "@element-plus/icons-vue";
import {
  ElButton,
  ElForm,
  ElFormItem,
  ElInput,
  ElMessage,
  ElSelect,
  ElOption
} from "element-plus";
import * as XLSX from "xlsx";
import { addDialog, closeDialog } from "@/components/ReDialog";
import { useUserStoreHook } from "@/store/modules/user";
import {
  getCreditBills,
  type CreditBillItem,
  type CreditBillsRequest
} from "@/api/enterprise-bills";

defineOptions({
  name: "EnterpriseBillsIndex"
});

const loading = ref(false);
const billList = ref<CreditBillItem[]>([]);
const selectedRows = ref<CreditBillItem[]>([]);
const currentPage = ref(1);
const pageSize = ref(20);
const total = ref(0);

const searchForm = ref({
  month: dayjs().startOf("month").format("YYYY-MM-DD"),
  corpNameLike: "",
  status: "",
  overdue: ""
});

const columns = [
  { prop: "shortCorpName", label: "客户简称", minWidth: 180 },
  { prop: "corpName", label: "企业名称", minWidth: 220 },
  { prop: "billNo", label: "账单编号", minWidth: 180 },
  { prop: "billDate", label: "出账日", minWidth: 120 },
  { prop: "billStartDate", label: "账单起始日", minWidth: 160 },
  { prop: "billEndDate", label: "账单截止日", minWidth: 160 },
  { prop: "repaymentDate", label: "最迟还款日", minWidth: 160 },
  { prop: "repaidTime", label: "实际还款时间", minWidth: 180 },
  { prop: "billAmount", label: "账单金额", minWidth: 120 },
  { prop: "paidAmount", label: "已还金额", minWidth: 120 },
  { prop: "debtAmount", label: "欠款金额", minWidth: 120 },
  { prop: "status", label: "状态", minWidth: 120 },
  { prop: "overdueText", label: "是否逾期", minWidth: 100 }
];

const availableColumns = [
  { key: "billDate", label: "出账日" },
  { key: "shortCorpName", label: "客户简称" },
  { key: "billNo", label: "账单编号" },
  { key: "billStartDate", label: "账单起始日" },
  { key: "billEndDate", label: "账单截止日" },
  { key: "repaymentDate", label: "最迟还款日" },
  { key: "repaidTime", label: "实际还款时间" },
  { key: "billAmount", label: "账单金额" },
  { key: "paidAmount", label: "已还金额" },
  { key: "debtAmount", label: "欠款金额" },
  { key: "status", label: "状态" },
  { key: "overdueText", label: "是否逾期" }
];

const exportColumns = ref<string[]>([
  "出账日",
  "客户简称",
  "账单起始日",
  "账单截止日",
  "最迟还款日",
  "账单金额",
]);

const statusOptions = [
  { label: "全部状态", value: "" },
  { label: "待还款", value: "UNPAID" },
  { label: "部分还款", value: "PARTIAL_REPAID" },
  { label: "已还清", value: "REPAID" },
  { label: "已关闭", value: "CLOSED" }
];

const overdueOptions = [
  { label: "全部", value: "" },
  { label: "是", value: "true" },
  { label: "否", value: "false" }
];

const monthValue = computed({
  get: () => (searchForm.value.month ? dayjs(searchForm.value.month).toDate() : null),
  set: value => {
    searchForm.value.month = value ? dayjs(value).startOf("month").format("YYYY-MM-DD") : "";
  }
});

const normalizeBillItem = (item: CreditBillItem) => {
  const debtAmount = Number(item.debtAmount || 0);
  const paidAmount = Number(item.paidAmount || 0);
  const billAmount = Number(item.billAmount || 0);
  const overdue =
    typeof item.overdue === "boolean"
      ? item.overdue
      : Boolean(item.repaymentDate && !item.repaidTime && dayjs(item.repaymentDate).isBefore(dayjs(), "day"));

  return {
    ...item,
    shortCorpName:
      item.shortCorpName || item.shortName || item.corpShortName || item.corpName || "",
    billAmount,
    paidAmount,
    debtAmount,
    repaymentDate: item.repaymentDate || item.latestRepayDate || "",
    overdue,
    overdueText: overdue ? "是" : "否"
  };
};

const buildRequestParams = (): Partial<CreditBillsRequest> => ({
  billDateStart: searchForm.value.month
    ? dayjs(searchForm.value.month).startOf("month").toISOString()
    : null,
  corpNameLike: searchForm.value.corpNameLike || null,
  status: searchForm.value.status || null,
  overdue:
    searchForm.value.overdue === ""
      ? null
      : searchForm.value.overdue === "true",
  pageNumber: currentPage.value,
  pageSize: pageSize.value
});

const loadBillData = async () => {
  try {
    loading.value = true;
    const response = await getCreditBills(buildRequestParams());

    if (response.code === 0 && response.data) {
      billList.value = (response.data.content || []).map(normalizeBillItem);
      total.value = response.data.totalElements || 0;
    } else {
      if (response.code === 401) {
        showLoginDialog();
        return;
      }
      billList.value = [];
      total.value = 0;
      ElMessage.error(response.message || "获取企业账单失败");
    }
  } catch (error: any) {
    if (error.response?.status === 401) {
      showLoginDialog();
      return;
    }
    billList.value = [];
    total.value = 0;
    ElMessage.error(error.message || "获取企业账单失败");
  } finally {
    loading.value = false;
  }
};

const handleSearch = async () => {
  currentPage.value = 1;
  await loadBillData();
};

const handleReset = async () => {
  searchForm.value = {
    month: dayjs().startOf("month").format("YYYY-MM-DD"),
    corpNameLike: "",
    status: "",
    overdue: ""
  };
  currentPage.value = 1;
  await loadBillData();
};

const handleSizeChange = async (size: number) => {
  pageSize.value = size;
  currentPage.value = 1;
  await loadBillData();
};

const handleCurrentChange = async (page: number) => {
  currentPage.value = page;
  await loadBillData();
};

const handleSelectionChange = (selection: CreditBillItem[]) => {
  selectedRows.value = selection;
};

const formatDateTime = (value?: string) => {
  if (!value) return "-";
  const date = dayjs(value);
  return date.isValid() ? date.format("YYYY-MM-DD") : value;
};

const formatAmount = (value?: number | string) => {
  if (value === null || value === undefined || value === "") return "-";
  const num = Number(value);
  return Number.isNaN(num) ? String(value) : num.toFixed(2);
};

const exportExcelData = (data: CreditBillItem[]) => {
  const selectedExportColumns = availableColumns.filter(col =>
    exportColumns.value.includes(col.label)
  );

  if (!selectedExportColumns.length) {
    ElMessage.warning("请至少选择一列进行导出");
    return;
  }

  const exportData = data.map(item => {
    const rowData: Record<string, any> = {};
    selectedExportColumns.forEach(col => {
      const rawValue = item[col.key];
      const value =
        ["billDate", "billStartDate", "billEndDate", "repaymentDate", "repaidTime"].includes(col.key)
          ? formatExportDateTime(rawValue)
          : col.key === "billAmount"
            ? formatExportAmount(rawValue)
            : rawValue;
      rowData[col.label] = value ?? "";
    });
    return rowData;
  });

  const ws = XLSX.utils.json_to_sheet(exportData);
  ws["!cols"] = selectedExportColumns.map(() => ({ wch: 18 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "企业账单");

  const monthText = searchForm.value.month
    ? dayjs(searchForm.value.month).format("YYYY-MM")
    : dayjs().format("YYYY-MM");
  XLSX.writeFile(wb, `企业账单_${monthText}.xlsx`);
};

const formatExportDateTime = (value?: string | number | Date) => {
  if (!value) return "";
  const date = dayjs(value);
  if (!date.isValid()) return String(value);
  return date.format("YYYY-MM-DD");
};

const formatExportAmount = (value?: string | number) => {
  if (value === null || value === undefined || value === "") return "";
  const num = Number(value);
  return Number.isNaN(num) ? value : num.toFixed(2);
};

const getAllDataAndExport = async () => {
  try {
    if (selectedRows.value.length > 0) {
      exportExcelData(selectedRows.value);
      ElMessage.success(`已导出选中的 ${selectedRows.value.length} 条账单`);
      return;
    }

    ElMessage.info("正在获取全部账单数据，请稍候...");
    const totalCount = total.value || 0;
    const pageSize = 200;

    let allData: CreditBillItem[] = [];

    if (totalCount > 200) {
      // 分段请求再合并
      const totalPages = Math.ceil(totalCount / pageSize);
      const requests = [];
      for (let page = 1; page <= totalPages; page++) {
        requests.push(
          getCreditBills({
            ...buildRequestParams(),
            pageNumber: page,
            pageSize
          })
        );
      }
      const responses = await Promise.all(requests);
      for (const response of responses) {
        if (response.code === 0 && response.data) {
          allData.push(...(response.data.content || []).map(normalizeBillItem));
        }
      }
    } else {
      // 数据量不大，一次请求即可
      const response = await getCreditBills({
        ...buildRequestParams(),
        pageNumber: 1,
        pageSize: totalCount || 200
      });
      if (response.code === 0 && response.data) {
        allData = (response.data.content || []).map(normalizeBillItem);
      } else {
        ElMessage.error(response.message || "获取全部账单失败");
        return;
      }
    }

    if (allData.length > 0) {
      exportExcelData(allData);
      ElMessage.success(`已导出全部 ${allData.length} 条账单`);
    } else {
      ElMessage.warning("未获取到账单数据");
    }
  } catch (error: any) {
    ElMessage.error(error.message || "导出失败");
  }
};

const selectAllColumns = () => {
  exportColumns.value = availableColumns.map(col => col.label);
};

const clearAllColumns = () => {
  exportColumns.value = [];
};

const headerCellClassName = ({ column }: { column: any }) => {
  return exportColumns.value.includes(column.label) ? "highlight-header" : "";
};

const loginFormData = ref({
  identity: "17688731379",
  password: "xin90879"
});
const loginLoading = ref(false);

const showLoginDialog = () => {
  loginFormData.value = { identity: "17688731379", password: "xin90879" };
  loginLoading.value = false;

  addDialog({
    title: "登录已过期，请重新登录",
    width: "400px",
    draggable: true,
    closeOnClickModal: false,
    closeOnPressEscape: false,
    showClose: false,
    hideFooter: true,
    contentRenderer: ({ options, index }) =>
      h("div", { style: { padding: "20px 20px 0" } }, [
        h(
          ElForm,
          {
            labelWidth: "70px",
            style: { maxWidth: "100%" }
          },
          () => [
            h(ElFormItem, { label: "账号" }, () =>
              h(ElInput, {
                modelValue: loginFormData.value.identity,
                "onUpdate:modelValue": (val: string) => {
                  loginFormData.value.identity = val;
                },
                placeholder: "请输入账号",
                clearable: true
              })
            ),
            h(ElFormItem, { label: "密码" }, () =>
              h(ElInput, {
                modelValue: loginFormData.value.password,
                "onUpdate:modelValue": (val: string) => {
                  loginFormData.value.password = val;
                },
                type: "password",
                placeholder: "请输入密码",
                showPassword: true,
                clearable: true
              })
            ),
            h(ElFormItem, { style: { marginBottom: "0" } }, () =>
              h(
                ElButton,
                {
                  type: "primary",
                  loading: loginLoading.value,
                  style: { width: "100%" },
                  onClick: async () => {
                    if (!loginFormData.value.identity || !loginFormData.value.password) {
                      ElMessage.warning("请输入账号和密码");
                      return;
                    }
                    loginLoading.value = true;
                    try {
                      const res = await useUserStoreHook().loginByReal({
                        identity: loginFormData.value.identity,
                        password: loginFormData.value.password
                      });
                      if (res.success) {
                        ElMessage.success("登录成功");
                        closeDialog(options, index);
                        await loadBillData();
                      } else {
                        ElMessage.error("登录失败，请检查账号密码");
                      }
                    } catch (error: any) {
                      ElMessage.error(`登录失败: ${error.message || "网络错误"}`);
                    } finally {
                      loginLoading.value = false;
                    }
                  }
                },
                () => "登录"
              )
            )
          ]
        )
      ])
  });
};

onMounted(() => {
  loadBillData();
});
</script>

<template>
  <div class="enterprise-bills-container">
    <div class="bills-content">
      <div class="search-section">
        <el-card>
          <el-form :inline="true" class="search-form">
            <el-form-item label="账单月份">
              <el-date-picker
                v-model="monthValue"
                type="month"
                placeholder="请选择账单月份"
                format="YYYY-MM"
                style="width: 160px"
              />
            </el-form-item>
            <el-form-item label="企业名称">
              <el-input
                v-model="searchForm.corpNameLike"
                placeholder="请输入企业名称"
                clearable
                @keyup.enter="handleSearch"
              />
            </el-form-item>
            <el-form-item label="状态">
              <el-select
                v-model="searchForm.status"
                placeholder="请选择状态"
                clearable
                style="width: 160px"
              >
                <el-option
                  v-for="item in statusOptions"
                  :key="item.value"
                  :label="item.label"
                  :value="item.value"
                />
              </el-select>
            </el-form-item>
            <el-form-item label="是否逾期">
              <el-select
                v-model="searchForm.overdue"
                placeholder="请选择"
                clearable
                style="width: 120px"
              >
                <el-option
                  v-for="item in overdueOptions"
                  :key="item.value"
                  :label="item.label"
                  :value="item.value"
                />
              </el-select>
            </el-form-item>
            <el-form-item>
              <el-button type="primary" @click="handleSearch">
                <el-icon><Search /></el-icon>
                搜索
              </el-button>
              <el-button @click="handleReset">
                <el-icon><Refresh /></el-icon>
                重置
              </el-button>
              <el-button
                type="success"
                @click="getAllDataAndExport"
                :disabled="billList.length === 0"
              >
                <el-icon><Download /></el-icon>
                导出Excel
              </el-button>
              <el-select
                v-model="exportColumns"
                multiple
                collapse-tags
                collapse-tags-tooltip
                :max-collapse-tags="8"
                placeholder="选择导出列"
                style="width: 620px; margin-left: 8px"
                :disabled="billList.length === 0"
              >
                <template #header>
                  <div class="export-header">
                    <span class="export-title">选择要导出的列</span>
                    <div class="export-actions">
                      <el-button size="small" text @click="selectAllColumns">
                        全选
                      </el-button>
                      <el-button size="small" text @click="clearAllColumns">
                        清空
                      </el-button>
                    </div>
                  </div>
                </template>
                <el-option
                  v-for="column in availableColumns"
                  :key="column.key"
                  :label="column.label"
                  :value="column.label"
                />
              </el-select>
            </el-form-item>
          </el-form>
        </el-card>
      </div>

      <div class="table-section">
        <el-card>
          <el-table
            :data="billList"
            v-loading="loading"
            stripe
            border
            style="width: 100%"
            :header-cell-class-name="headerCellClassName"
            @selection-change="handleSelectionChange"
          >
            <el-table-column type="selection" width="55" />
            <el-table-column
              label="序号"
              width="60"
              type="index"
              :index="index => (currentPage - 1) * pageSize + index + 1"
              align="center"
            />
            <el-table-column
              v-for="col in columns"
              :key="col.prop"
              :prop="col.prop"
              :label="col.label"
              :min-width="col.minWidth"
              show-overflow-tooltip
            >
              <template
                v-if="
                  ['billDate', 'billStartDate', 'billEndDate', 'repaymentDate', 'repaidTime'].includes(
                    col.prop
                  )
                "
                #default="{ row }"
              >
                {{ formatDateTime(row[col.prop]) }}
              </template>
              <template
                v-else-if="['billAmount', 'paidAmount', 'debtAmount'].includes(col.prop)"
                #default="{ row }"
              >
                {{ formatAmount(row[col.prop]) }}
              </template>
              <template #default="{ row }" v-else>
                {{ row[col.prop] ?? "-" }}
              </template>
            </el-table-column>
          </el-table>

          <div class="pagination-wrapper">
            <el-pagination
              v-model:current-page="currentPage"
              v-model:page-size="pageSize"
              :page-sizes="[20, 50, 100, 300]"
              :total="total"
              layout="total, sizes, prev, pager, next, jumper"
              @size-change="handleSizeChange"
              @current-change="handleCurrentChange"
            />
          </div>
        </el-card>
      </div>
    </div>
  </div>
</template>

<style scoped>
.enterprise-bills-container {
  position: relative;
  overflow: hidden;
}

.bills-content {
  background: rgba(255, 255, 255, 0.9);
  border-radius: 8px;
  box-shadow: 0 2px 12px rgba(0, 0, 0, 0.1);
  padding: 20px;
  min-height: 400px;
}

.search-section {
  margin-bottom: 20px;
}

.search-form {
  display: flex;
  align-items: center;
  flex-wrap: wrap;
  gap: 16px;
}

.table-section {
  margin-top: 20px;
}

.pagination-wrapper {
  display: flex;
  justify-content: center;
  margin-top: 20px;
}

.export-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 8px 12px;
  border-bottom: 1px solid #ebeef5;
}

.export-title {
  font-size: 14px;
  font-weight: 600;
  color: #303133;
}

.export-actions {
  display: flex;
  gap: 8px;
}

.el-card {
  border-radius: 8px;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
}

@media (max-width: 768px) {
  .bills-content {
    padding: 15px;
  }

  .search-form {
    flex-direction: column;
    align-items: stretch;
  }

  .search-form .el-form-item {
    margin-bottom: 10px;
  }
}

:deep(.highlight-header) {
  background-color: #ecf5ff !important;
  color: #409eff !important;
  font-weight: bold;
}
</style>
