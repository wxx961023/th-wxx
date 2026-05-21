import { http } from "@/utils/http";
import { formatToken } from "@/utils/auth";
import Cookies from "js-cookie";

export interface CreditBillsRequest {
  billDateStart?: string | null;
  corpNameLike?: string | null;
  status?: string | null;
  overdue?: boolean | null;
  pageNumber: number;
  pageSize: number;
}

export interface CreditBillItem {
  billId?: number | string;
  billNo?: string;
  billDate?: string;
  billStartDate?: string;
  billEndDate?: string;
  repaymentDate?: string;
  latestRepayDate?: string;
  repaidTime?: string;
  corpName?: string;
  billAmount?: number | string;
  paidAmount?: number | string;
  debtAmount?: number | string;
  status?: string;
  overdue?: boolean;
  creditBillFeeInfo?: Record<string, any>;
  [key: string]: any;
}

export interface CreditBillsResponse {
  success?: boolean;
  code?: number;
  message?: string;
  data?: {
    content: CreditBillItem[];
    totalElements: number;
    totalPages: number;
    size: number;
    number: number;
    first: boolean;
    last: boolean;
    empty: boolean;
  };
}

const DEFAULT_SEARCH_PARAMS: CreditBillsRequest = {
  billDateStart: null,
  corpNameLike: null,
  status: null,
  overdue: null,
  pageNumber: 1,
  pageSize: 20
};

const getAccessToken = (): string | null => {
  try {
    const cookieToken = Cookies.get("authorized-token");
    if (cookieToken && cookieToken.trim().startsWith("{")) {
      try {
        const parsedCookie = JSON.parse(cookieToken);
        if (parsedCookie?.accessToken) return parsedCookie.accessToken;
      } catch (error) {
        console.error("解析 Cookie token 失败:", error);
      }
    }

    const userInfo = localStorage.getItem("user-info");
    if (userInfo) {
      if (userInfo.trim().startsWith("{")) {
        try {
          const parsedUserInfo = JSON.parse(userInfo);
          if (parsedUserInfo?.accessToken) return parsedUserInfo.accessToken;
        } catch (error) {
          console.error("解析 localStorage user-info 失败:", error);
        }
      } else if (userInfo.length > 10) {
        return userInfo;
      }
    }

    const userToken = localStorage.getItem("userToken");
    if (userToken) return userToken;

    return null;
  } catch (error) {
    console.error("获取 token 失败:", error);
    return null;
  }
};

export const getCreditBills = (
  params: Partial<CreditBillsRequest>
): Promise<CreditBillsResponse> => {
  const searchParams = {
    ...DEFAULT_SEARCH_PARAMS,
    ...params
  };

  const token = getAccessToken();

  return http.request<CreditBillsResponse>(
    "post",
    "/admin/v1/finance/getCreditBills",
    {
      data: searchParams,
      baseURL: "https://staff-api-gateway.teyixing.com",
      headers: {
        Authorization: token ? formatToken(token) : undefined,
        "Content-Type": "application/json"
      }
    }
  );
};
