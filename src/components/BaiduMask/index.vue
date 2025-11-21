<template>
  <teleport to="body">
    <!-- 全屏遮罩层 -->
    <transition name="mask-fade">
      <div v-show="visible" class="baidu-mask">
        <div class="mask-content" @click.stop>
          <!-- 百度风格的内容区域 -->
          <div class="search-container">
            <div class="logo">
              <img src="@/assets/bd.png" alt="百度" />
            </div>
            <div class="search-box">
              <div class="search-input-wrapper">
                <input
                  type="text"
                  placeholder="搜索一下"
                  class="search-input"
                  @click.stop
                  autofocus
                />
                <button class="search-btn" @click.stop>百度一下</button>
              </div>
            </div>

            <!-- 热搜榜 -->
            <div class="hot-search">
              <div class="hot-title">百度热搜</div>
              <div class="hot-links">
                <a href="#" class="hot-link" @click.stop>春节档电影</a>
                <a href="#" class="hot-link" @click.stop>油价调整</a>
                <a href="#" class="hot-link" @click.stop>天气查询</a>
                <a href="#" class="hot-link" @click.stop>春节放假</a>
                <a href="#" class="hot-link" @click.stop>热搜榜</a>
                <a href="#" class="hot-link" @click.stop>科技新闻</a>
              </div>
            </div>
          </div>
        </div>
      </div>
    </transition>

    <!-- 右下角浮动控制按钮 -->
    <transition name="float-btn-fade">
      <div
        v-show="showFloatBtn"
        class="float-control-btn"
        :class="{ 'mask-active': visible }"
        @click="toggleMask"
        :title="visible ? '关闭遮罩' : '显示遮罩'"
      >
        <svg
          v-if="visible"
          class="icon"
          viewBox="0 0 24 24"
          fill="currentColor"
        >
          <path
            d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"
          />
        </svg>
        <svg v-else class="icon" viewBox="0 0 24 24" fill="currentColor">
          <path
            d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-2 15l-5-5 1.41-1.41L10 14.17l7.59-7.59L19 8l-9 9z"
          />
        </svg>
      </div>
    </transition>
  </teleport>
</template>

<script setup lang="ts">
import { ref, onMounted, onBeforeUnmount } from "vue";

const visible = ref(false);
const showFloatBtn = ref(true);
let inactivityTimer: number | null = null;
const INACTIVITY_TIMEOUT = 1000 * 10; // 10秒

const toggleMask = () => {
  visible.value = !visible.value;
};

const showMask = () => {
  visible.value = true;
};

const hideMask = () => {
  visible.value = false;
};

const handleKeydown = (e: KeyboardEvent) => {
  if (e.key === "Escape" && visible.value) {
    visible.value = false;
  }
  resetInactivityTimer();
};

const handleUserActivity = () => {
  resetInactivityTimer();
};

const resetInactivityTimer = () => {
  if (inactivityTimer) {
    clearTimeout(inactivityTimer);
  }

  inactivityTimer = window.setTimeout(() => {
    showMask();
  }, INACTIVITY_TIMEOUT);
};

onMounted(() => {
  document.addEventListener("keydown", handleKeydown);
  document.addEventListener("mousemove", handleUserActivity);
  document.addEventListener("mousedown", handleUserActivity);
  document.addEventListener("click", handleUserActivity);
  document.addEventListener("scroll", handleUserActivity);
  document.addEventListener("keypress", handleUserActivity);

  // 启动计时器
  resetInactivityTimer();
});

onBeforeUnmount(() => {
  document.removeEventListener("keydown", handleKeydown);
  document.removeEventListener("mousemove", handleUserActivity);
  document.removeEventListener("mousedown", handleUserActivity);
  document.removeEventListener("click", handleUserActivity);
  document.removeEventListener("scroll", handleUserActivity);
  document.removeEventListener("keypress", handleUserActivity);

  if (inactivityTimer) {
    clearTimeout(inactivityTimer);
  }
});
</script>

<style lang="scss" scoped>
.baidu-mask {
  position: fixed;
  top: 0;
  left: 0;
  width: 100vw;
  height: 100vh;
  background: #fff;
  z-index: 9999;
  display: flex;
  flex-direction: column;
  align-items: center;
  cursor: pointer;
  font-family: Arial, "PingFang SC", "Microsoft YaHei", sans-serif;

  .mask-content {
    width: 100%;
    max-width: 640px;
    padding: 0 20px;
    margin-top: 120px;
    text-align: center;
  }

  .search-container {
    .logo {
      margin-bottom: 24px;

      img {
        height: 120px;
        width: auto;
      }

      .logo-text {
        display: inline-block;
        font-size: 28px;
        font-weight: 400;
        color: #4e6ef2;
        letter-spacing: 2px;
        margin-top: 10px;
      }
    }

    .search-box {
      position: relative;
      width: 100%;
      max-width: 512px;
      height: 44px;
      margin: 0 auto 24px;

      .search-input-wrapper {
        width: 100%;
        height: 100%;
        border: 2px solid #4e6ef2;
        border-radius: 0 10px 10px 0;
        background: #fff;
        display: flex;
        align-items: center;
        overflow: hidden;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);

        .search-input {
          flex: 1;
          border: none;
          background: transparent;
          padding: 10px 16px;
          font-size: 16px;
          outline: none;
          color: #222;
          height: 100%;
        }

        .search-input::placeholder {
          color: #9195a3;
        }

        .search-btn {
          width: 108px;
          height: 44px;
          background: #4e6ef2;
          color: white;
          border: none;
          font-size: 17px;
          cursor: pointer;
          transition: background 0.3s ease;
          padding: 0;
          border-radius: 0 10px 10px 0;
          margin-left: -2px;

          &:hover {
            background: #4662d9;
          }

          &:active {
            background: #3b5ce5;
          }
        }
      }
    }

    .hot-search {
      margin-top: 30px;

      .hot-title {
        font-size: 14px;
        color: #9195a3;
        margin-bottom: 12px;
      }

      .hot-links {
        display: flex;
        flex-wrap: wrap;
        gap: 12px;
        justify-content: center;

        .hot-link {
          font-size: 14px;
          color: #626675;
          text-decoration: none;
          padding: 6px 12px;
          border-radius: 4px;
          transition: all 0.2s ease;

          &:hover {
            background: #f5f5f5;
            color: #4e6ef2;
          }
        }
      }
    }

    .bottom-nav {
      position: absolute;
      bottom: 40px;
      left: 0;
      right: 0;
      display: flex;
      justify-content: center;
      gap: 32px;
      font-size: 14px;
      color: #626675;

      .nav-item {
        cursor: pointer;
        transition: color 0.2s ease;

        &:hover {
          color: #4e6ef2;
        }
      }
    }
  }
}

.float-control-btn {
  position: fixed;
  bottom: 30px;
  right: 30px;
  width: 56px;
  height: 56px;
  background: #4285f4;
  color: white;
  border-radius: 50%;
  display: flex;
  align-items: center;
  justify-content: center;
  cursor: pointer;
  box-shadow: 0 4px 16px rgba(66, 133, 244, 0.4);
  transition: all 0.3s ease;
  z-index: 10000;

  &:hover {
    transform: scale(1.1);
    box-shadow: 0 6px 20px rgba(66, 133, 244, 0.6);
  }

  &.mask-active {
    background: #f44336;
    box-shadow: 0 4px 16px rgba(244, 67, 54, 0.4);

    &:hover {
      box-shadow: 0 6px 20px rgba(244, 67, 54, 0.6);
    }
  }

  .icon {
    width: 24px;
    height: 24px;
    transition: transform 0.3s ease;
  }
}

// 动画效果
.mask-fade-enter-active,
.mask-fade-leave-active {
  transition: all 0.1s ease;
}

.mask-fade-enter-from,
.mask-fade-leave-to {
  opacity: 0;
  transform: scale(0.9);
}

.float-btn-fade-enter-active,
.float-btn-fade-leave-active {
  transition: all 0.1s ease;
}

.float-btn-fade-enter-from,
.float-btn-fade-leave-to {
  opacity: 0;
  transform: scale(0.5) translateY(20px);
}

@keyframes fadeInDown {
  from {
    opacity: 0;
    transform: translateY(-30px);
  }
  to {
    opacity: 1;
    transform: translateY(0);
  }
}

@keyframes fadeInUp {
  from {
    opacity: 0;
    transform: translateY(30px);
  }
  to {
    opacity: 1;
    transform: translateY(0);
  }
}

// 响应式设计
@media (max-width: 768px) {
  .baidu-mask {
    .search-container {
      .logo {
        font-size: 36px;
        margin-bottom: 30px;
      }

      .search-box {
        margin: 0 20px;

        .search-input {
          font-size: 16px;
          padding: 10px 16px;
        }

        .search-btn {
          padding: 10px 20px;
          font-size: 14px;
        }
      }
    }
  }

  .float-control-btn {
    bottom: 20px;
    right: 20px;
    width: 48px;
    height: 48px;

    .icon {
      width: 20px;
      height: 20px;
    }
  }
}
</style>
