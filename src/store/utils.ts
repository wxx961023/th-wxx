export { store } from "@/store"
export { routerArrays } from "@/layout/types"
export { router, resetRouter, constantMenus } from "@/router"
export { getConfig, responsiveStorageNameSpace } from "@/config"
// 路由工具函数已简化，移除复杂的权限管理函数
export {
  isUrl,
  isEqual,
  isNumber,
  debounce,
  isBoolean,
  getKeyList,
  storageLocal,
  deviceDetection
} from "@pureadmin/utils"
export type {
  setType,
  appType,
  userType,
  multiType,
  cacheType,
  positionType
} from "./types"
