import { http } from "@/utils/http"

export type UserResult = {
  success: boolean
  data: {
    /** 头像 */
    avatar: string
    /** 用户名 */
    username: string
    /** 昵称 */
    nickname: string
    /** 当前登录用户的角色 */
    roles: Array<string>
    /** 按钮级别权限 */
    permissions: Array<string>
    /** `token` */
    accessToken: string
    /** 用于调用刷新`accessToken`的接口时所需的`token` */
    refreshToken: string
    /** `accessToken`的过期时间（格式'xxxx/xx/xx xx:xx:xx'） */
    expires: Date
  }
}

export type RefreshTokenResult = {
  success: boolean
  data: {
    /** `token` */
    accessToken: string
    /** 用于调用刷新`accessToken`的接口时所需的`token` */
    refreshToken: string
    /** `accessToken`的过期时间（格式'xxxx/xx/xx xx:xx:xx'） */
    expires: Date
  }
}

/** 真实登录API响应类型 */
export type RealLoginResponse = {
  code: number
  message?: string
  data: {
    token: string
    staff: {
      id: number
      name: string
      workNo: string
      mobile?: string
      email?: string
    }
  }
}

/** 登录（Mock接口 - 保留用于开发环境） */
export const getLogin = (data?: object) => {
  return http.request<UserResult>("post", "/login", { data })
}

/** 真实登录接口 - 调用特易行员工登录API */
export const realLogin = async (data: { identity: string; password: string }): Promise<UserResult> => {
  try {
    const response = await fetch("https://staff-api-gateway.teyixing.com/v1/staff/login", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Accept": "application/json"
      },
      body: JSON.stringify(data)
    })

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}: ${response.statusText}`)
    }

    const result: RealLoginResponse = await response.json()

    if (result.code === 0 && result.data && result.data.token) {
      // 保存原始登录信息到localStorage（供其他模块使用）
      localStorage.setItem("userToken", result.data.token)
      localStorage.setItem("userInfo", JSON.stringify(result.data))

      // 计算过期时间（默认30天后过期）
      const expiresDate = new Date()
      expiresDate.setDate(expiresDate.getDate() + 30)

      // 转换为框架需要的格式
      return {
        success: true,
        data: {
          avatar: "",
          username: result.data.staff.workNo,
          nickname: result.data.staff.name,
          roles: ["admin"],
          permissions: ["*:*:*"],
          accessToken: result.data.token,
          refreshToken: result.data.token, // 使用同一个token
          expires: expiresDate
        }
      }
    } else {
      return {
        success: false,
        data: null
      } as unknown as UserResult
    }
  } catch (error) {
    console.error("登录失败:", error)
    throw error
  }
}

/** 刷新`token` */
export const refreshTokenApi = (data?: object) => {
  return http.request<RefreshTokenResult>("post", "/refresh-token", { data })
}
