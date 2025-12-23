/**
 * 艾比森部门对账人联系配置
 * 用于记录艾比森各部门对应的对账人及邮箱信息
 */

// 部门到对账人信息的映射表
export const ABSEN_DEPARTMENT_TO_CONTACT_MAP: Record<string, {
  accountant: string // 对账人姓名
  email: string // 邮箱地址
}> = {
  "北美Live大客户部": { accountant: "袁嘉惠", email: "claire.yuan5266@absen.com" },
  "北美Live业务一区": { accountant: "陈洁", email: "joyce.chen@absen.com" },
  "北美Live业务二区": { accountant: "陈洁", email: "joyce.chen@absen.com" },
  "北美Live业务三区": { accountant: "陈洁", email: "joyce.chen@absen.com" },
  "北美ProAV大客户部": { accountant: "石映梅", email: "Yoyo.shi@absen.com" },
  "北美ProAV业务一区": { accountant: "石映梅", email: "Yoyo.shi@absen.com" },
  "北美ProAV业务二区": { accountant: "石映梅", email: "Yoyo.shi@absen.com" },
  "北美ProAV业务三区": { accountant: "石映梅", email: "Yoyo.shi@absen.com" },
  "广告与体育市场(美国)": { accountant: "石映梅", email: "Yoyo.shi@absen.com" },
  "加拿大市场": { accountant: "石映梅", email: "Yoyo.shi@absen.com" },
  "美国平台": { accountant: "石映梅", email: "Yoyo.shi@absen.com" },
  "北美地区部管理": { accountant: "石映梅", email: "Yoyo.shi@absen.com" },
  "拉美业务一区": { accountant: "丁燕", email: "vicky.ding@absen.com" },
  "拉美业务二区": { accountant: "丁燕", email: "vicky.ding@absen.com" },
  "墨西哥市场": { accountant: "丁燕", email: "vicky.ding@absen.com" },
  "巴西市场": { accountant: "丁燕", email: "vicky.ding@absen.com" },
  "拉美地区部管理": { accountant: "丁燕", email: "vicky.ding@absen.com" },
  "欧洲Live业务一区": { accountant: "郭文静", email: "ava.guo@absen.com" },
  "欧洲Live业务二区": { accountant: "郭文静", email: "ava.guo@absen.com" },
  "欧洲ProAV大客户部": { accountant: "郭文静", email: "ava.guo@absen.com" },
  "大洋洲市场": { accountant: "李倩", email: "chloe.li@absen.com" },
  "德语市场": { accountant: "李倩", email: "chloe.li@absen.com" },
  "英国市场": { accountant: "李倩", email: "chloe.li@absen.com" },
  "欧洲ProAV业务一区": { accountant: "蓝家裕", email: "jane.lan@absen.com" },
  "欧洲ProAV业务二区": { accountant: "蓝家裕", email: "jane.lan@absen.com" },
  "欧洲ProAV业务三区": { accountant: "蓝家裕", email: "jane.lan@absen.com" },
  "欧洲平台": { accountant: "赖帆", email: "jessica.lai@absen.com" },
  "欧洲地区部管理": { accountant: "赖帆", email: "jessica.lai@absen.com" },
  "AbsenLive日本市场": { accountant: "陈毅玲", email: "elaine.chen@absen.com" },
  "日本市场": { accountant: "陈毅玲", email: "elaine.chen@absen.com" },
  "日本地区部管理": { accountant: "陈毅玲", email: "elaine.chen@absen.com" },
  "港澳台市场": { accountant: "袁文静", email: "joyce.yuan@absen.com" },
  "新马市场": { accountant: "袁文静", email: "joyce.yuan@absen.com" },
  "亚太业务二区": { accountant: "袁文静", email: "joyce.yuan@absen.com" },
  "韩国市场": { accountant: "袁文静", email: "joyce.yuan@absen.com" },
  "印度市场": { accountant: "袁文静", email: "joyce.yuan@absen.com" },
  "印尼市场": { accountant: "袁文静", email: "joyce.yuan@absen.com" },
  "亚太二地区部管理": { accountant: "袁文静", email: "joyce.yuan@absen.com" },
  "马来西亚市场": { accountant: "袁文静", email: "joyce.yuan@absen.com" },
  "俄罗斯市场": { accountant: "胡晶", email: "jing.hu@absen.com" },
  "俄语市场": { accountant: "胡晶", email: "jing.hu@absen.com" },
  "亚太三地区部管理": { accountant: "胡晶", email: "jing.hu@absen.com" },
  "亚太业务四区": { accountant: "吴洁", email: "abby.wu@absen.com" },
  "亚太四地区部管理": { accountant: "吴洁", email: "abby.wu@absen.com" },
  "沙特市场": { accountant: "余婉霞", email: "nora.yu@absen.com" },
  "亚太业务五区": { accountant: "余婉霞", email: "nora.yu@absen.com" },
  "中东平台": { accountant: "余婉霞", email: "nora.yu@absen.com" },
  "亚太五地区部管理": { accountant: "余婉霞", email: "nora.yu@absen.com" },
  "瑞乐事业部管理": { accountant: "王文纯", email: "sunny.wang@absen.com" },
  "瑞乐国际销售部": { accountant: "王文纯", email: "sunny.wang@absen.com" },
  "北美服务运营部": { accountant: "田文胜", email: "steven.tian@usabsen.com" },
  "拉美服务运营部": { accountant: "牛颖浩", email: "walker.niu@absen.com" },
  "欧洲服务运营部": { accountant: "方学群", email: "fisher.fang@absen.com" },
  "日本服务运营部": { accountant: "徐甫", email: "china@absen.com" },
  "亚太服务运营二部": { accountant: "钟长培", email: "peter.zhong@absen.com" },
  "亚太服务运营三部": { accountant: "", email: "" },
  "亚太服务运营四部": { accountant: "刘文涛", email: "enda.liu@absen.com" },
  "亚太服务运营五部": { accountant: "叶润泽", email: "ramzee.ye3457@absen.com" },
  "国际市场营销部": { accountant: "吴怡凤", email: "yvon.woo@absen.com" },
  "市场策划部": { accountant: "刘芳", email: "fang@absen.com" },
  "方案研究部": { accountant: "刘点", email: "dian.liu@absen.com" }
}

/**
 * 根据部门名称获取对账人信息
 * @param department 部门名称
 * @returns 对账人信息对象，如果未找到则返回 undefined
 */
export function getAbsenContactByDepartment(
  department: string
): { accountant: string; email: string } | undefined {
  return ABSEN_DEPARTMENT_TO_CONTACT_MAP[department?.trim()]
}

/**
 * 根据部门名称获取对账人姓名
 * @param department 部门名称
 * @returns 对账人姓名，如果未找到则返回 undefined
 */
export function getAbsenAccountantByDepartment(
  department: string
): string | undefined {
  return ABSEN_DEPARTMENT_TO_CONTACT_MAP[department?.trim()]?.accountant
}

/**
 * 根据部门名称获取对账人邮箱
 * @param department 部门名称
 * @returns 对账人邮箱，如果未找到则返回 undefined
 */
export function getAbsenEmailByDepartment(
  department: string
): string | undefined {
  return ABSEN_DEPARTMENT_TO_CONTACT_MAP[department?.trim()]?.email
}

/**
 * 获取所有艾比森部门名称列表
 * @returns 部门名称数组
 */
export function getAllAbsenDepartments(): string[] {
  return Object.keys(ABSEN_DEPARTMENT_TO_CONTACT_MAP)
}

/**
 * 获取所有艾比森对账人姓名列表
 * @returns 对账人姓名数组
 */
export function getAllAbsenAccountants(): string[] {
  return Object.values(ABSEN_DEPARTMENT_TO_CONTACT_MAP).map(contact => contact.accountant)
}

/**
 * 获取所有艾比森邮箱地址列表
 * @returns 邮箱地址数组
 */
export function getAllAbsenEmails(): string[] {
  return Object.values(ABSEN_DEPARTMENT_TO_CONTACT_MAP).map(contact => contact.email)
}
