import Cookies from "js-cookie"
import NProgress from "@/utils/progress"
import remainingRouter from "./modules/remaining"
import { cloneDeep } from "@pureadmin/utils"
import { getHistoryMode } from "./utils"
import {
  type Router,
  type RouteRecordRaw,
  createRouter
} from "vue-router"
import { multipleTabsKey } from "@/utils/auth"

/** 手动定义的静态路由 */
/** 账单下载路由 */
const homeRoute: RouteRecordRaw = {
  path: "/",
  name: "Home",
  component: () => import("@/layout/index.vue"),
  redirect: "/welcome",
  meta: {
    icon: "ep:home-filled",
    title: "账单下载",
    rank: 0
  },
  children: [
    {
      path: "/welcome",
      name: "Welcome",
      component: () => import("@/views/welcome/index.vue"),
      meta: {
        title: "账单下载"
      }
    }
  ]
}

/** PDF工具路由 */
const pdfRoute: RouteRecordRaw = {
  path: "/pdf",
  name: "PdfModule",
  component: () => import("@/layout/index.vue"),
  redirect: "/pdf/batch-rename",
  meta: {
    icon: "ep:document",
    title: "PDF工具",
    rank: 1
  },
  children: [
    {
      path: "/pdf/batch-rename",
      name: "PdfBatchRename",
      component: () => import("@/views/welcome/pdf.vue"),
      meta: {
        title: "发票"
      }
    },
    {
      path: "/pdf/parser",
      name: "PdfParser",
      component: () => import("@/views/pdf/index.vue"),
      meta: {
        title: "PDF解析"
      }
    }
  ]
}

/** 账单分账路由 */
const billSplitRoute: RouteRecordRaw = {
  path: "/bill-split",
  name: "BillSplitModule",
  component: () => import("@/layout/index.vue"),
  redirect: "/bill-split/index",
  meta: {
    icon: "ep:money",
    title: "账单分账",
    rank: 2
  },
  children: [
    {
      path: "/bill-split/index",
      name: "BillSplitIndex",
      component: () => import("@/views/bill-split/index.vue"),
      meta: {
        title: "账单分账"
      }
    }
  ]
}

/** 机构客户路由 */
const institutionalClientsRoute: RouteRecordRaw = {
  path: "/institutional-clients",
  name: "InstitutionalClientsModule",
  component: () => import("@/layout/index.vue"),
  redirect: "/institutional-clients/index",
  meta: {
    icon: "ep:office-building",
    title: "机构客户",
    rank: 3
  },
  children: [
    {
      path: "/institutional-clients/index",
      name: "InstitutionalClientsIndex",
      component: () => import("@/views/institutional-clients/index.vue"),
      meta: {
        title: "机构客户"
      }
    }
  ]
}

/** 导出静态路由 */
export const constantRoutes: Array<RouteRecordRaw> = [homeRoute, pdfRoute, billSplitRoute, institutionalClientsRoute]

/** 初始的静态路由，用于退出登录时重置路由 */
const initConstantRoutes: Array<RouteRecordRaw> = cloneDeep(constantRoutes)

/** 用于渲染菜单，保持原始层级 */
export const constantMenus: Array<any> = [homeRoute, pdfRoute, billSplitRoute, institutionalClientsRoute]

/** 不参与菜单的路由 */
export const remainingPaths = Object.keys(remainingRouter).map(v => {
  return remainingRouter[v].path
})

/** 创建路由实例 */
export const router: Router = createRouter({
  history: getHistoryMode(import.meta.env.VITE_ROUTER_HISTORY),
  routes: constantRoutes.concat(...(remainingRouter as any)),
  strict: true,
  scrollBehavior(to, from, savedPosition) {
    return new Promise(resolve => {
      if (savedPosition) {
        return savedPosition
      } else {
        if (from.meta.saveScrollTop) {
          const top: number =
            document.documentElement.scrollTop || document.body.scrollTop
          resolve({ left: 0, top })
        } else {
          resolve({ left: 0, top: 0 })
        }
      }
    })
  }
})

/** 重置路由 */
export function resetRouter() {
  router.clearRoutes()
  for (const route of initConstantRoutes.concat(...(remainingRouter as any))) {
    router.addRoute(route)
  }
  router.options.routes = constantRoutes.concat(...(remainingRouter as any))
}

/** 路由白名单 */
const whiteList = ["/login"]

/** 简化的路由守卫 */
router.beforeEach((to, _from, next) => {
  NProgress.start()

  // 设置页面标题
  if (to.meta?.title) {
    document.title = to.meta.title as string
  }

  // 简化的路由守卫逻辑
  if (Cookies.get(multipleTabsKey)) {
    // 已登录状态，允许访问所有路由
    next()
  } else {
    // 未登录状态，只允许访问白名单路由
    if (to.path !== "/login" && whiteList.indexOf(to.path) === -1) {
      next({ path: "/login" })
    } else {
      next()
    }
  }
})

router.afterEach(() => {
  NProgress.done()
})

export default router
