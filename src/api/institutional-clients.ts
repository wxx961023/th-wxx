import { http } from "@/utils/http"
import { formatToken, getToken } from "@/utils/auth"
import Cookies from "js-cookie"

// 机构客户搜索请求参数类型
export interface CorpSearchRequest {
  nameLike?: string | null
  corpType?: string | null
  status?: string | null
  salesManagerNameLike?: string | null
  customerManagerNameLike?: string | null
  settlementStaffNameLike?: string | null
  codeLike?: string | null
  canPayByUatp?: boolean | null
  accountNoOfUatp?: string
  corpLevelId?: number | null
  miniCorp?: boolean | null
  recommendNameLike?: string | null
  registerSource?: string | null
  spCode?: string | null
  platformCode?: string | null
  hasContract?: boolean | null
  contractValidityStatus?: string | null
  createDateStart?: string
  createDateEnd?: string
  businessUnits?: string[]
  pageNumber: number
  pageSize: number
}

// 机构客户响应数据类型
export interface CorpItem {
  id: number
  name: string
  shortName?: string
  code: string
  businessUnit?: string
  province?: string
  city?: string
  area?: string
  corpType?: string
  contactName?: string
  contactPerson?: string
  contactPhone?: string
  address?: string
  status?: string
  createTime?: string
  hasContractDesc?: string
  contractValidityStatusDesc?: string
  billAmount?: number | string
  // 嵌套的员工数组
  salesStaffs?: Array<{ staffName?: string }>
  customerStaffs?: Array<{ staffName?: string }>
  settlementStaffs?: Array<{ staffName?: string }>
  // 其他可能的字段
  [key: string]: any
}

// API响应类型
export interface CorpSearchResponse {
  success: boolean
  data: {
    content: CorpItem[]
    totalElements: number
    totalPages: number
    size: number
    number: number
    first: boolean
    last: boolean
    empty: boolean
  }
  message?: string
  code?: number
}

// 默认搜索参数
const DEFAULT_SEARCH_PARAMS: CorpSearchRequest = {
  nameLike: null,
  corpType: null,
  status: null,
  salesManagerNameLike: null,
  customerManagerNameLike: null,
  settlementStaffNameLike: "王欣欣",
  codeLike: null,
  canPayByUatp: null,
  accountNoOfUatp: "",
  corpLevelId: null,
  miniCorp: null,
  recommendNameLike: null,
  registerSource: null,
  spCode: null,
  platformCode: null,
  hasContract: null,
  contractValidityStatus: null,
  createDateStart: "",
  createDateEnd: "",
  businessUnits: ["TMC", "GJ_WD", "GJ_TY", "GN_TY"],
  pageNumber: 1,
  pageSize: 200
}

/**
 * 搜索机构客户列表
 * @param params 搜索参数
 * @returns Promise<CorpSearchResponse>
 */
export const searchCorps = (params: Partial<CorpSearchRequest>): Promise<CorpSearchResponse> => {
  // 合并默认参数和传入参数
  const searchParams = {
    ...DEFAULT_SEARCH_PARAMS,
    ...params
  }

  console.log('🚀 发送API请求，参数:', searchParams)

  // 获取token并手动添加到请求头
  const getAccessToken = (): string | null => {
    try {
      // 首先尝试从Cookie获取
      const cookieToken = Cookies.get('authorized-token')
      if (cookieToken && cookieToken.trim().startsWith('{')) {
        try {
          const parsedCookie = JSON.parse(cookieToken)
          if (parsedCookie && parsedCookie.accessToken) {
            console.log('✅ 从Cookie获取到token:', parsedCookie.accessToken.substring(0, 20) + '...')
            return parsedCookie.accessToken
          }
        } catch (cookieError) {
          console.error('❌ Cookie token JSON解析失败:', cookieError.message)
        }
      }

      // 然后尝试从localStorage获取
      const userInfo = localStorage.getItem('user-info')
      if (userInfo) {
        if (userInfo.trim().startsWith('{')) {
          try {
            const parsedUserInfo = JSON.parse(userInfo)
            if (parsedUserInfo && parsedUserInfo.accessToken) {
              console.log('✅ 从localStorage获取到token:', parsedUserInfo.accessToken.substring(0, 20) + '...')
              return parsedUserInfo.accessToken
            }
          } catch (storageError) {
            console.error('❌ localStorage JSON解析失败:', storageError.message)
          }
        } else if (userInfo.length > 10) {
          console.log('✅ 使用localStorage原始字符串作为token:', userInfo.substring(0, 20) + '...')
          return userInfo
        }
      }

      return null
    } catch (error) {
      console.error('❌ 获取token时发生意外错误:', error.message)
      return null
    }
  }

  const token = getAccessToken()

  return http.request<CorpSearchResponse>("post", "/admin/v1/corp/searchCorps", {
    data: searchParams,
    // 覆盖默认的基础URL，使用指定的API地址
    baseURL: 'https://staff-api-gateway.teyixing.com',
    // 手动添加Authorization头，确保token带上
    headers: {
      'Authorization': token ? formatToken(token) : undefined,
      'Content-Type': 'application/json'
    }
  })
}

/**
 * 获取机构客户详情（如果需要的话）
 * @param id 客户ID
 * @returns Promise
 */
export const getCorpDetail = (id: number) => {
  // 健壮地获取token
  const getAccessToken = (): string | null => {
    try {
      // 首先尝试从Cookie获取
      const cookieToken = Cookies.get('authorized-token')
      if (cookieToken) {
        try {
          const parsedCookie = JSON.parse(cookieToken)
          if (parsedCookie && parsedCookie.accessToken) {
            return parsedCookie.accessToken
          }
        } catch (cookieError) {
          console.error('Failed to parse cookie token:', cookieError)
        }
      }

      // 然后尝试从localStorage获取
      const userInfo = localStorage.getItem('user-info')
      if (userInfo) {
        try {
          const parsedUserInfo = JSON.parse(userInfo)
          if (parsedUserInfo && parsedUserInfo.accessToken) {
            return parsedUserInfo.accessToken
          }
        } catch (storageError) {
          console.error('Failed to parse localStorage user-info:', storageError)

          // 如果JSON解析失败，尝试直接检查是否是token字符串
          if (userInfo && typeof userInfo === 'string' && userInfo.length > 10) {
            return userInfo
          }
        }
      }

      return null
    } catch (error) {
      console.error('Unexpected error getting token:', error)
      return null
    }
  }

  const token = getAccessToken()

  return http.request("get", `/admin/v1/corp/${id}`, {
    baseURL: 'https://staff-api-gateway.teyixing.com',
    // 手动添加Authorization头，确保token带上
    headers: {
      'Authorization': token ? formatToken(token) : undefined,
      'Content-Type': 'application/json'
    }
  })
}
