<script setup lang="ts">
import { ref, onMounted, h } from "vue";
import {
  Search,
  Refresh,
  Plus,
  Edit,
  Delete,
  Download,
  CaretBottom
} from "@element-plus/icons-vue";
import {
  ElMessage,
  ElMessageBox,
  ElForm,
  ElFormItem,
  ElInput,
  ElButton
} from "element-plus";
import {
  searchCorps,
  type CorpSearchRequest,
  type CorpItem
} from "@/api/institutional-clients";
import * as XLSX from "xlsx";
import { addDialog, closeDialog } from "@/components/ReDialog";
import { useUserStoreHook } from "@/store/modules/user";

defineOptions({
  name: "InstitutionalClientsIndex"
});

// 数据状态
const loading = ref(false);
const clientList = ref<CorpItem[]>([]);
const searchKeyword = ref("");
const settlementStaffNameKeyword = ref("王欣欣");
const currentPage = ref(1);
const pageSize = ref(20);
const total = ref(0);

// 表格列定义
const columns = [
  { prop: "shortName", label: "客户简称", minWidth: 250 },
  { prop: "businessUnit", label: "客户类型", minWidth: 100 },
  { prop: "location", label: "客户地址", minWidth: 180 },
  { prop: "corpType", label: "所属行业", minWidth: 100 },
  { prop: "createTime", label: "新增日期", minWidth: 140 },
  { prop: "contactName", label: "联系人", minWidth: 100 },
  { prop: "salesStaffName", label: "销售经理", minWidth: 100 },
  { prop: "customerStaffName", label: "客服经理", minWidth: 100 },
  { prop: "settlementStaffName", label: "结算经理", minWidth: 100 },
  { prop: "hasContractDesc", label: "合同状态", minWidth: 90 },
  { prop: "contractValidityStatusDesc", label: "合同时效", minWidth: 90 },
  { prop: "creditAmountText", label: "初始授信额度", minWidth: 120 },
  { prop: "billDayText", label: "账单日", minWidth: 100 },
  { prop: "billDurationText", label: "账期", minWidth: 100 },
  { prop: "billingPeriodText", label: "结算方式", minWidth: 100 },
  { prop: "billAmount", label: "上期账单", minWidth: 100 },
  { prop: "spName", label: "归属服务商", minWidth: 150 }
];

/**
 * 加载机构客户数据
 */
const loadClientData = async () => {
  try {
    loading.value = true;

    const params: Partial<CorpSearchRequest> = {
      pageNumber: currentPage.value,
      pageSize: pageSize.value,
      nameLike: searchKeyword.value || null,
      settlementStaffNameLike: settlementStaffNameKeyword.value || null
    };

    const response = await searchCorps(params);
    console.log("📥 API响应:", response);

    if (response.code === 0) {
      // 映射API响应数据到表格数据
      clientList.value = response.data.content.map(item => ({
        ...item,
        // 组合地址字段
        location: [item.province, item.city, item.area]
          .filter(Boolean)
          .join(" "),
        // 获取嵌套的员工姓名
        salesStaffName:
          item.salesStaffs && item.salesStaffs.length > 0
            ? item.salesStaffs[0].staffName
            : "-",
        customerStaffName:
          item.customerStaffs && item.customerStaffs.length > 0
            ? item.customerStaffs[0].staffName
            : "-",
        settlementStaffName:
          item.settlementStaffs && item.settlementStaffs.length > 0
            ? item.settlementStaffs[0].staffName
            : "-"
      }));

      total.value = response.data.totalElements;
      console.log("✅ 数据加载成功，共", total.value, "条记录");
      console.log("📊 处理后的数据样例:", clientList.value[0]);
    } else {
      console.error("❌ API返回失败:", response);
      if (response.code == 401) {
        showLoginDialog();
        return;
      }
      ElMessage.error(response.message || "获取机构客户数据失败");
    }
  } catch (error: any) {
    console.error("💥 加载机构客户数据失败:", error);

    // 处理不同类型的错误
    if (error.response?.status === 401) {
      showLoginDialog();
    } else if (error.response?.status === 403) {
      ElMessage.error("权限不足，无法访问该数据");
    } else if (error.code === "NETWORK_ERROR") {
      ElMessage.error("网络连接失败，请检查网络设置");
    } else {
      ElMessage.error(error.message || "加载机构客户数据失败");
    }

    // 清空数据
    clientList.value = [];
    total.value = 0;
  } finally {
    loading.value = false;
    console.log("🏁 数据加载完成，loading状态已关闭");
  }
};

// 搜索处理
const handleSearch = async () => {
  currentPage.value = 1; // 搜索时重置到第一页
  await loadClientData();
};

// 重置搜索
const handleReset = async () => {
  searchKeyword.value = "";
  settlementStaffNameKeyword.value = "";
  currentPage.value = 1;
  await loadClientData();
};

// 新增机构
const handleAdd = () => {
  ElMessage.info("新增机构功能待实现");
  // TODO: 实现新增逻辑，可能需要跳转到新增页面或打开对话框
};

// 编辑机构
const handleEdit = (row: CorpItem) => {
  ElMessage.info(`编辑机构功能待实现: ${row.name}`);
  // TODO: 实现编辑逻辑，可能需要跳转到编辑页面或打开对话框
};

// 删除机构
const handleDelete = async (row: CorpItem) => {
  try {
    await ElMessageBox.confirm(
      `确定要删除机构 "${row.name}" 吗？此操作不可恢复。`,
      "确认删除",
      {
        confirmButtonText: "确定",
        cancelButtonText: "取消",
        type: "warning"
      }
    );

    ElMessage.info("删除机构功能待实现");
    // TODO: 实现删除逻辑，调用删除API
    // const response = await deleteCorp(row.id);
    // if (response.success) {
    //   ElMessage.success('删除成功');
    //   await loadClientData();
    // } else {
    //   ElMessage.error('删除失败');
    // }
  } catch (error) {
    // 用户取消删除
    console.log("用户取消删除");
  }
};

// 分页处理
const handleSizeChange = async (size: number) => {
  pageSize.value = size;
  currentPage.value = 1; // 改变页面大小时重置到第一页
  await loadClientData();
};

const handleCurrentChange = async (page: number) => {
  currentPage.value = page;
  await loadClientData();
};

// 格式化状态 - 根据API返回的实际状态值调整
const formatStatus = (status: string) => {
  const statusMap: Record<string, { text: string; type: string }> = {
    ACTIVE: { text: "活跃", type: "success" },
    INACTIVE: { text: "停用", type: "danger" },
    PENDING: { text: "待审核", type: "warning" },
    SUSPENDED: { text: "暂停", type: "info" },
    // 兼容可能的英文状态
    active: { text: "活跃", type: "success" },
    inactive: { text: "停用", type: "danger" },
    pending: { text: "待审核", type: "warning" }
  };
  return statusMap[status] || { text: status || "未知", type: "info" };
};

// 格式化时间
const formatDateTime = (dateTime: string) => {
  if (!dateTime) return "-";
  try {
    return new Date(dateTime).toLocaleString("zh-CN");
  } catch {
    return dateTime;
  }
};

// 登录弹窗相关
const loginFormData = ref({
  identity: "17688731379",
  password: "xin90879"
});
const loginLoading = ref(false);

// 显示登录弹窗
const showLoginDialog = () => {
  // 重置登录表单（使用默认账号密码）
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
                    if (
                      !loginFormData.value.identity ||
                      !loginFormData.value.password
                    ) {
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
                        // 重新加载数据
                        await loadClientData();
                      } else {
                        ElMessage.error("登录失败，请检查账号密码");
                      }
                    } catch (err: any) {
                      console.error("登录错误:", err);
                      ElMessage.error(`登录失败: ${err.message || "网络错误"}`);
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

// 组件挂载时加载初始数据
onMounted(() => {
  console.log("🎯 onMounted 钩子被调用了！");
  console.log("🎯 当前路由:", window.location.pathname);
  console.log("🎯 即将调用 loadClientData...");
  loadClientData();
});

// 手动测试API调用
const testApiCall = () => {
  console.log("🧪 手动测试API调用");
  loadClientData();
};

// 表格选中状态
const selectedRows = ref<CorpItem[]>([]);
const tableRef = ref();

const exportColumns = ref<string[]>([
  "客户简称",
  "销售经理",
  "结算经理",
  "合同状态",
  "合同时效",
  "账单日",
  "账期",
  "结算方式"
]);

// 所有可选的列定义
const availableColumns = [
  { key: "shortName", label: "客户简称" },
  { key: "salesStaffName", label: "销售经理" },
  { key: "settlementStaffName", label: "结算经理" },
  { key: "hasContractDesc", label: "合同状态" },
  { key: "contractValidityStatusDesc", label: "合同时效" },
  { key: "creditAmountText", label: "初始授信额度" },
  { key: "billDayText", label: "账单日" },
  { key: "billDurationText", label: "账期" },
  { key: "billingPeriodText", label: "结算方式" },
  { key: "businessUnit", label: "客户类型" },
  { key: "location", label: "客户地址" },
  { key: "corpType", label: "所属行业" },
  { key: "createTime", label: "新增日期" },
  { key: "contactName", label: "联系人" },
  { key: "customerStaffName", label: "客服经理" },
  { key: "billAmount", label: "上期账单" },
  { key: "spName", label: "归属服务商" }
];

// 获取全部数据并导出Excel
const getAllDataAndExport = async () => {
  try {
    ElMessage.info("正在获取全部数据，请稍候...");

    // 如果没有选中数据，则获取全部数据
    if (selectedRows.value.length === 0) {
      const totalCount = total.value || 100;
      const pageSize = 120;

      // 对原始数据做统一转换的辅助函数
      const transformItem = (item: any) => ({
        ...item,
        location: [item.province, item.city, item.area]
          .filter(Boolean)
          .join(" "),
        salesStaffName:
          item.salesStaffs && item.salesStaffs.length > 0
            ? item.salesStaffs[0].staffName
            : "-",
        customerStaffName:
          item.customerStaffs && item.customerStaffs.length > 0
            ? item.customerStaffs[0].staffName
            : "-",
        settlementStaffName:
          item.settlementStaffs && item.settlementStaffs.length > 0
            ? item.settlementStaffs[0].staffName
            : "-"
      });

      let allData: any[] = [];

      if (totalCount > 120) {
        // 分段请求再合并
        const totalPages = Math.ceil(totalCount / pageSize);
        const requests = [];
        for (let page = 1; page <= totalPages; page++) {
          const params: Partial<CorpSearchRequest> = {
            pageNumber: page,
            pageSize,
            nameLike: searchKeyword.value || null,
            settlementStaffNameLike: settlementStaffNameKeyword.value || null
          };
          requests.push(searchCorps(params));
        }
        console.log(`📡 分段获取全部数据，共 ${totalCount} 条，分 ${totalPages} 页请求`);
        const responses = await Promise.all(requests);
        for (const response of responses) {
          if (response.code === 0 && response.data) {
            allData.push(...response.data.content.map(transformItem));
          }
        }
      } else {
        // 数据量不大，一次请求即可
        const allDataParams: Partial<CorpSearchRequest> = {
          pageNumber: 1,
          pageSize: totalCount,
          nameLike: searchKeyword.value || null,
          settlementStaffNameLike: settlementStaffNameKeyword.value || null
        };
        console.log("📡 获取全部数据，参数:", allDataParams);
        const response = await searchCorps(allDataParams);
        if (response.code === 0 && response.data) {
          allData = response.data.content.map(transformItem);
        } else {
          ElMessage.error("获取全部数据失败");
          return;
        }
      }

      if (allData.length > 0) {
        await exportExcelData(allData);
        ElMessage.success(`成功导出全部 ${allData.length} 条数据！`);
      } else {
        ElMessage.warning("未获取到数据");
      }
    } else {
      // 导出选中的数据
      await exportExcelData(selectedRows.value);
      ElMessage.success(`成功导出选中的 ${selectedRows.value.length} 条数据！`);
    }
  } catch (error) {
    console.error("导出Excel失败:", error);
    ElMessage.error("导出Excel失败，请重试");
  }
};

// 导出Excel数据
const exportExcelData = async (data: CorpItem[]) => {
  // 使用用户选择的列
  const selectedExportColumns = availableColumns.filter(col =>
    exportColumns.value.includes(col.label)
  );

  if (selectedExportColumns.length === 0) {
    ElMessage.warning("请至少选择一列进行导出");
    return;
  }

  // 准备导出数据
  const exportData = data.map((item, index) => {
    const rowData: any = {};
    // const rowData: any = { '序号': index + 1 }

    selectedExportColumns.forEach(col => {
      let value = item[col.key];
      if (value === null || value === undefined || value === "") {
        value = "";
      }
      rowData[col.label] = value;
    });

    return rowData;
  });

  // 创建工作簿
  const ws = XLSX.utils.json_to_sheet(exportData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "机构客户数据");

  // 设置列宽
  const colWidths = selectedExportColumns.map(() => ({ wch: 15 }));
  colWidths.unshift({ wch: 8 }); // 序号列
  ws["!cols"] = colWidths;

  // 生成Excel文件
  const fileName = `机构客户数据_${new Date().toLocaleDateString("zh-CN").replace(/\//g, "-")}.xlsx`;
  XLSX.writeFile(wb, fileName);
};

// 表格选择变化处理
const handleSelectionChange = (selection: CorpItem[]) => {
  selectedRows.value = selection;
};

// 全选所有列
const selectAllColumns = () => {
  exportColumns.value = availableColumns.map(col => col.label);
};

// 清空所有选择
const clearAllColumns = () => {
  exportColumns.value = [];
};

// 表头高亮样式逻辑
const headerCellClassName = ({ column }: { column: any }) => {
  if (exportColumns.value.includes(column.label)) {
    return "highlight-header";
  }
  return "";
};
</script>

<template>
  <div class="institutional-clients-container">
    <div class="clients-content">
      <div class="search-section">
        <el-card>
          <el-form :inline="true" class="search-form">
            <el-form-item label="机构名称">
              <el-input
                v-model="searchKeyword"
                placeholder="请输入机构名称"
                clearable
                @keyup.enter="handleSearch"
              />
            </el-form-item>
            <el-form-item label="结算经理">
              <el-input
                v-model="settlementStaffNameKeyword"
                placeholder="请输入结算经理"
                clearable
                @keyup.enter="handleSearch"
              />
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
              <!-- 导出Excel按钮 -->
              <el-button
                type="success"
                @click="getAllDataAndExport"
                :disabled="clientList.length === 0"
              >
                <el-icon><Download /></el-icon>
                导出Excel
              </el-button>
              <!-- 选择导出列 -->
              <el-select
                v-model="exportColumns"
                multiple
                collapse-tags
                collapse-tags-tooltip
                :max-collapse-tags="10"
                placeholder="选择导出列"
                style="width: 840px; margin-left: 8px"
                :disabled="clientList.length === 0"
                clearable
                @clear="() => (exportColumns = [])"
              >
                <template #header>
                  <div
                    style="
                      display: flex;
                      justify-content: space-between;
                      align-items: center;
                      padding: 8px 12px;
                      border-bottom: 1px solid #ebeef5;
                    "
                  >
                    <span
                      style="font-size: 14px; font-weight: 600; color: #303133"
                      >选择要导出的列</span
                    >
                    <div style="display: flex; gap: 8px">
                      <el-button size="small" text @click="selectAllColumns"
                        >全选</el-button
                      >
                      <el-button size="small" text @click="clearAllColumns"
                        >清空</el-button
                      >
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

      <!-- 数据表格区域 -->
      <div class="table-section">
        <el-card>
          <el-table
            ref="tableRef"
            :data="clientList"
            v-loading="loading"
            stripe
            border
            style="width: 100%"
            :header-cell-class-name="headerCellClassName"
            @selection-change="handleSelectionChange"
          >
            <!-- 选择列 -->
            <el-table-column type="selection" width="55" />
            <!-- 序号列 -->
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
              <template #default="{ row }" v-if="col.prop === 'status'">
                <el-tag :type="formatStatus(row.status).type as any">
                  {{ formatStatus(row.status).text }}
                </el-tag>
              </template>
              <template
                #default="{ row }"
                v-else-if="col.prop === 'createTime'"
              >
                {{ formatDateTime(row[col.prop]) }}
              </template>
              <template #default="{ row }" v-else>
                {{ row[col.prop] || "-" }}
              </template>
            </el-table-column>
            <!--
            <el-table-column label="操作" width="200" fixed="right">
              <template #default="{ row }">
                <el-button
                  type="primary"
                  size="small"
                  @click="handleEdit(row)"
                >
                  <el-icon><Edit /></el-icon>
                  编辑
                </el-button>
                <el-button
                  type="danger"
                  size="small"
                  @click="handleDelete(row)"
                >
                  <el-icon><Delete /></el-icon>
                  删除
                </el-button>
              </template>
            </el-table-column> -->
          </el-table>

          <!-- 分页 -->
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
.institutional-clients-container {
  position: relative;
  overflow: hidden;
}

.clients-content {
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

.el-card {
  border-radius: 8px;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
}

.el-table {
  border-radius: 4px;
}

/* 响应式样式 */
@media (max-width: 768px) {
  .clients-content {
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

:deep(.highlight-header .cell) {
  color: #409eff !important;
}
</style>
