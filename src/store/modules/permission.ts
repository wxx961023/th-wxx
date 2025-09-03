import { defineStore } from "pinia"
import {
  type cacheType,
  store,
  debounce,
  getKeyList,
  constantMenus
} from "../utils"
import { useMultiTagsStoreHook } from "./multiTags"

export const usePermissionStore = defineStore("pure-permission", {
  state: () => ({
    // 静态路由生成的菜单
    constantMenus,
    // 整体路由生成的菜单（静态、动态）
    wholeMenus: constantMenus,
    // 整体路由（一维数组格式）
    flatteningRoutes: constantMenus,
    // 缓存页面keepAlive
    cachePageList: []
  }),
  actions: {
    /** 简化的菜单处理 */
    handleWholeMenus(routes: any[]) {
      this.wholeMenus = this.constantMenus.concat(routes)
      this.flatteningRoutes = this.constantMenus.concat(routes)
    },
    cacheOperate({ mode, name }: cacheType) {
      const delIndex = this.cachePageList.findIndex(v => v === name)
      switch (mode) {
        case "refresh":
          this.cachePageList = this.cachePageList.filter(v => v !== name)
          break
        case "add":
          this.cachePageList.push(name)
          break
        case "delete":
          delIndex !== -1 && this.cachePageList.splice(delIndex, 1)
          break
      }
      /** 监听缓存页面是否存在于标签页，不存在则删除 */
      debounce(() => {
        let cacheLength = this.cachePageList.length
        const nameList = getKeyList(useMultiTagsStoreHook().multiTags, "name")
        while (cacheLength > 0) {
          nameList.findIndex(v => v === this.cachePageList[cacheLength - 1]) ===
            -1 &&
            this.cachePageList.splice(
              this.cachePageList.indexOf(this.cachePageList[cacheLength - 1]),
              1
            )
          cacheLength--
        }
      })()
    },
    /** 清空缓存页面 */
    clearAllCachePage() {
      this.wholeMenus = []
      this.cachePageList = []
    }
  }
})

export function usePermissionStoreHook() {
  return usePermissionStore(store)
}
