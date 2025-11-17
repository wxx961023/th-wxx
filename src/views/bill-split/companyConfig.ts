// 公司信息配置文件
export interface CompanyInfo {
  shortName: string
  contact: string
  phone: string
  otherFullName?: string

}

export interface PersonInfo {
  fullName: string // 公司全称
  shortName: string
  contact: string
  phone: string
}

export interface CompanyConfig {
  nameMapping: Record<string, CompanyInfo>
  getCompanyInfo(fullName: string): CompanyInfo
}

export interface PersonConfig {
  nameMapping: Record<string, PersonInfo>
  getPersonInfo(personName: string): PersonInfo
}

const companyConfig: CompanyConfig = {
  // 公司名称映射配置
  nameMapping: {
    深圳市宝辰鑫激光科技有限公司苏州分公司: {
      shortName: "宝辰鑫激光-苏州分公司",
      contact: "王宛平",
      phone: "15083407402"
    },
    深圳市宝辰鑫激光科技有限公司: {
      shortName: "宝辰鑫激光",
      contact: "傅心源",
      phone: "18682006186"
    },
    深圳市创鑫激光股份有限公司: {
      shortName: "创鑫激光股份",
      contact: "曾玉娟",
      phone: "13620930071"
    },
    深圳市创鑫激光股份有限公司北京技术分公司: {
      shortName: "北京创鑫智造激光",
      otherFullName: "北京创鑫智造激光科技有限公司",
      contact: "冯萌",
      phone: "15033218468"
    },
    北京创鑫智造激光科技有限公司: {
      shortName: "北京创鑫智造激光",
      contact: "冯萌",
      phone: "15033218468"
    },
    深圳市桓日激光有限公司: {
      shortName: "桓日激光",
      contact: "韦雪娜",
      phone: "13430456543"
    },
    深圳市嘉鑫激光科技有限公司: {
      shortName: "嘉鑫激光",
      contact: "邵莹",
      phone: "18986300821"
    },
    深圳市欧亚激光智能科技有限公司: {
      shortName: "欧亚激光智能",
      contact: "钟颖",
      phone: "13750518967"
    },
    苏州创鑫激光科技有限公司: {
      shortName: "苏州创鑫激光",
      contact: "魏函",
      phone: "17356949230"
    },
    武汉创鑫激光科技有限公司: {
      shortName: "武汉创鑫激光",
      contact: "黄放",
      phone: "13986152289"
    }
  },

  // 获取公司映射信息
  getCompanyInfo(fullName: string): CompanyInfo {
    return (
      this.nameMapping[fullName] || {
        shortName: fullName, // 如果没有配置，使用原名
        contact: "",
        phone: ""
      }
    )
  }
}

// 大梦龙途人员配置
const personConfig: PersonConfig = {
  // 人员名称映射配置
  nameMapping: {
    陈敏铷: {
      fullName: "广州煋禾网络有限公司",
      shortName: "陈敏铷",
      contact: "陈敏铷",
      phone: "13538774228"
    },
    洪晴: {
      fullName: "湖南大梦龙途文化传播有限公司", // 暂时使用默认值，等用户补充
      shortName: "洪晴",
      contact: "洪晴",
      phone: "16670917363"
    },
    万语馨: {
      fullName: "深圳市大梦龙途文化传播有限公司", // 暂时使用默认值，等用户补充
      shortName: "万语馨",
      contact: "万语馨",
      phone: "17727420275"
    }
  },

  // 获取人员映射信息
  getPersonInfo(personName: string): PersonInfo {
    return (
      this.nameMapping[personName] || {
        fullName: personName, // 如果没有配置，使用原名
        shortName: personName,
        contact: personName,
        phone: ""
      }
    )
  }
}

export default companyConfig
export { personConfig }
