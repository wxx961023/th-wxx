// 公司信息配置文件
export interface CompanyInfo {
  shortName: string;
  contact: string;
  phone: string;
}

export interface CompanyConfig {
  nameMapping: Record<string, CompanyInfo>;
  getCompanyInfo(fullName: string): CompanyInfo;
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
      shortName: "创鑫激光股份-北京技术分公司",
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
    );
  }
};

export default companyConfig;
