import {
  type RouterHistory,
  type RouteRecordRaw,
  createWebHistory,
  createWebHashHistory
} from "vue-router"
import { router } from "./index"
import { isProxy, toRaw } from "vue"
import { usePermissionStoreHook } from "@/store/modules/permission"

/** 通过指定 `key` 获取父级路径集合，默认 `key` 为 `path` */
function getParentPaths(value: string, routes: RouteRecordRaw[], key = "path") {
  // 深度遍历查找
  function dfs(routes: RouteRecordRaw[], value: string, parents: string[]) {
    for (let i = 0; i < routes.length; i++) {
      const item = routes[i]
      // 返回父级path
      if (item[key] === value) return parents
      // children不存在或为空则不递归
      if (!item.children || !item.children.length) continue
      // 往下查找时将当前path入栈
      parents.push(item.path)

      if (dfs(item.children, value, parents).length) return parents
      // 深度遍历查找未找到时当前path 出栈
      parents.pop()
    }
    // 未找到时返回空数组
    return []
  }

  return dfs(routes, value, [])
}

/** 查找对应 `path` 的路由信息 */
function findRouteByPath(path: string, routes: RouteRecordRaw[]) {
  let res = routes.find((item: { path: string }) => item.path == path)
  if (res) {
    return isProxy(res) ? toRaw(res) : res
  } else {
    for (let i = 0; i < routes.length; i++) {
      if (
        routes[i].children instanceof Array &&
        routes[i].children.length > 0
      ) {
        res = findRouteByPath(path, routes[i].children)
        if (res) {
          return isProxy(res) ? toRaw(res) : res
        }
      }
    }
    return null
  }
}

function addPathMatch() {
  if (!router.hasRoute("pathMatch")) {
    router.addRoute({
      path: "/:pathMatch(.*)",
      name: "pathMatch",
      redirect: "/error/404"
    })
  }
}

/** 简化的路由初始化函数 */
function initRouter() {
  return new Promise(resolve => {
    // 添加404路由匹配
    addPathMatch()

    // 初始化菜单数据
    usePermissionStoreHook().handleWholeMenus([])

    resolve(router)
  })
}

/** 获取路由历史模式 */
function getHistoryMode(routerHistory): RouterHistory {
  // len为1 代表只有历史模式 为2 代表历史模式中存在base参数
  const historyMode = routerHistory.split(",")
  const leftMode = historyMode[0]
  const rightMode = historyMode[1]
  // no param
  if (historyMode.length === 1) {
    if (leftMode === "hash") {
      return createWebHashHistory("")
    } else if (leftMode === "h5") {
      return createWebHistory("")
    }
  } //has param
  else if (historyMode.length === 2) {
    if (leftMode === "hash") {
      return createWebHashHistory(rightMode)
    } else if (leftMode === "h5") {
      return createWebHistory(rightMode)
    }
  }
}

/** 简化的权限检查函数 - 总是返回true */
function hasAuth(_value: string | Array<string>): boolean {
  return true
}

/** 获取当前页面按钮级别的权限 */
function getAuths(): Array<string> {
  return []
}

/** 简化的获取顶级菜单函数 */
function getTopMenu(_tag = false): any {
  return { path: "/welcome", name: "Welcome", meta: { title: "首页" } }
}

/** 简化的处理缓存路由函数 */
function handleAliveRoute(_route: any, _mode?: string) {
  // 简化实现，不做任何处理
}

export {
  hasAuth,
  getAuths,
  getTopMenu,
  handleAliveRoute,
  initRouter,
  addPathMatch,
  getHistoryMode,
  getParentPaths,
  findRouteByPath
}
