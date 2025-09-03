<script setup lang="ts">
defineOptions({
  name: "403"
});
</script>

<template>
  <div class="error-403-container">
    <iframe
      src="/th/batch_download.html"
      sandbox="allow-scripts allow-same-origin allow-forms allow-popups allow-downloads"
      title="批量账单下载工具"
    />
  </div>
</template>

<style scoped>
.error-403-container {
  min-height: 100vh;
  background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
  position: relative;
  overflow: hidden;
}
</style>
