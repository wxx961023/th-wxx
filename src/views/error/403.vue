<script setup lang="ts">
defineOptions({
  name: "403"
});
</script>

<template>
  <div class="error-403-container"></div>
</template>

<style scoped>
.error-403-container {
  min-height: 100vh;
  background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
  position: relative;
  overflow: hidden;
}

/* 错误信息区域 */
.error-info-section {
  transition: all 0.3s ease-in-out;
  position: relative;
  z-index: 1;
}

.error-info-section.collapsed {
  transform: scale(0.8);
  opacity: 0.3;
  pointer-events: none;
}

.button-group {
  display: flex;
  gap: 12px;
  flex-wrap: wrap;
}

/* iframe容器样式 */
.iframe-container {
  position: fixed;
  top: 0;
  left: 0;
  right: 0;
  bottom: 0;
  z-index: 1000;
  background: rgba(0, 0, 0, 0.8);
  backdrop-filter: blur(5px);
  display: flex;
  flex-direction: column;
  animation: fadeIn 0.3s ease-in-out;
}

@keyframes fadeIn {
  from {
    opacity: 0;
    transform: scale(0.95);
  }
  to {
    opacity: 1;
    transform: scale(1);
  }
}

/* iframe头部 */
.iframe-header {
  background: #fff;
  padding: 16px 24px;
  border-bottom: 1px solid #e5e7eb;
  display: flex;
  justify-content: space-between;
  align-items: center;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
}

.iframe-title {
  display: flex;
  align-items: center;
  gap: 8px;
  font-size: 16px;
  font-weight: 600;
  color: #1f2937;
}

.iframe-controls {
  display: flex;
  gap: 8px;
}

/* iframe内容区域 */
.batch-download-iframe {
  flex: 1;
  width: 100%;
  border: none;
  background: #fff;
  min-height: calc(100vh - 80px);
}

/* 加载状态 */
.iframe-loading {
  flex: 1;
  display: flex;
  align-items: center;
  justify-content: center;
  background: #fff;
}

.loading-placeholder {
  width: 100%;
  height: 400px;
  display: flex;
  align-items: center;
  justify-content: center;
}

/* 错误状态 */
.iframe-error {
  flex: 1;
  display: flex;
  align-items: center;
  justify-content: center;
  background: #fff;
  padding: 40px;
}

/* 响应式设计 */
@media (max-width: 768px) {
  .error-info-section {
    padding: 20px;
  }

  .error-info-section .flex {
    flex-direction: column;
    text-align: center;
    height: auto !important;
    min-height: 500px;
  }

  .error-info-section .ml-12 {
    margin-left: 0;
    margin-top: 20px;
  }

  .button-group {
    justify-content: center;
  }

  .iframe-header {
    padding: 12px 16px;
    flex-direction: column;
    gap: 12px;
    text-align: center;
  }

  .iframe-title {
    justify-content: center;
  }

  .batch-download-iframe {
    min-height: calc(100vh - 120px);
  }
}

@media (max-width: 480px) {
  .iframe-container {
    position: fixed;
    top: 0;
    left: 0;
    right: 0;
    bottom: 0;
  }

  .iframe-header {
    padding: 8px 12px;
  }

  .iframe-title {
    font-size: 14px;
  }

  .batch-download-iframe {
    min-height: calc(100vh - 100px);
  }
}

/* 暗色主题适配 */
@media (prefers-color-scheme: dark) {
  .error-403-container {
    background: linear-gradient(135deg, #1f2937 0%, #111827 100%);
  }

  .iframe-header {
    background: #374151;
    border-bottom-color: #4b5563;
    color: #f9fafb;
  }

  .iframe-title {
    color: #f9fafb;
  }

  .iframe-loading,
  .iframe-error {
    background: #374151;
  }
}

/* 安全提示样式 */
.iframe-container::before {
  content: "";
  position: absolute;
  top: 0;
  left: 0;
  right: 0;
  bottom: 0;
  pointer-events: none;
  z-index: -1;
}

/* 滚动条样式 */
.batch-download-iframe::-webkit-scrollbar {
  width: 8px;
}

.batch-download-iframe::-webkit-scrollbar-track {
  background: #f1f1f1;
}

.batch-download-iframe::-webkit-scrollbar-thumb {
  background: #c1c1c1;
  border-radius: 4px;
}

.batch-download-iframe::-webkit-scrollbar-thumb:hover {
  background: #a8a8a8;
}
</style>
