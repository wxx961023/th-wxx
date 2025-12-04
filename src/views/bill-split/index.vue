<script setup lang="ts">
import { ref } from "vue";
import CxjgBillSplit from "./components/cxjg.vue";
import DmltBillSplit from "./components/dmlt.vue";
import SdmBillSplit from "./components/sdm.vue";
import FzjtBillSplit from "./components/fzjt.vue";
import GdhyqcBillSplit from "./components/gdhyqc.vue";

defineOptions({
  name: "BillSplitIndex"
});

// 当前选中的组件
const selectedComponent = ref<string>("cxjg");
</script>

<template>
  <div class="bill-split-container">
    <div class="bill-split-content">
      <!-- 组件选择区域 -->
      <div class="component-selection">
        <el-card>
          <h3>选择分账方式</h3>
          <el-radio-group v-model="selectedComponent" size="large">
            <el-radio label="cxjg">
              <div class="component-option">
                <h4>创鑫激光</h4>
                <p>适用于创鑫激光的账单分账处理</p>
              </div>
            </el-radio>
            <el-radio label="dameng-longtu">
              <div class="component-option">
                <h4>大梦龙途</h4>
                <p>适用于大梦龙途的账单分账处理</p>
              </div>
            </el-radio>
            <el-radio label="sdm">
              <div class="component-option">
                <h4>森达美</h4>
                <p>适用于森达美的账单分账处理（按人名拆分）</p>
              </div>
            </el-radio>
            <el-radio label="fzjt">
              <div class="component-option">
                <h4>纺织集团</h4>
                <p>适用于纺织集团的账单分账处理（按人名合并多工作表）</p>
              </div>
            </el-radio>
            <el-radio label="gdhyqc">
              <div class="component-option">
                <h4>广东鸿粤汽车</h4>
                <p>适用于广东鸿粤汽车的账单分账处理（按开票单位拆分）</p>
              </div>
            </el-radio>
          </el-radio-group>
        </el-card>
      </div>

      <!-- 动态组件渲染 -->
      <div class="component-content">
        <CxjgBillSplit v-if="selectedComponent === 'cxjg'" />
        <DmltBillSplit v-else-if="selectedComponent === 'dameng-longtu'" />
        <SdmBillSplit v-else-if="selectedComponent === 'sdm'" />
        <FzjtBillSplit v-else-if="selectedComponent === 'fzjt'" />
        <GdhyqcBillSplit v-else-if="selectedComponent === 'gdhyqc'" />
      </div>
    </div>
  </div>
</template>

<style scoped>
.bill-split-container {
  position: relative;
  overflow: hidden;
}

.bill-split-header p {
  color: #606266;
  margin: 0;
  font-size: 16px;
}

.bill-split-content {
  background: rgba(255, 255, 255, 0.9);
  border-radius: 8px;
  box-shadow: 0 2px 12px rgba(0, 0, 0, 0.1);
  padding: 40px;
  min-height: 400px;
}

.component-selection {
  margin-bottom: 30px;
}

.component-selection h3 {
  margin: 0 0 20px 0;
  color: #303133;
  font-size: 18px;
}

.component-option {
  padding: 10px;
}

.component-option h4 {
  margin: 0 0 5px 0;
  color: #303133;
  font-size: 16px;
}

.component-option p {
  margin: 0;
  color: #909399;
  font-size: 14px;
}

.component-content {
  margin-top: 30px;
}

.placeholder-content {
  display: flex;
  align-items: center;
  justify-content: center;
  height: 200px;
}
</style>
