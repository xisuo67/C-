using Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace Tools
{
    /// <summary>
    /// Execl操作帮助类
    /// </summary>
    public static class ExcelUtil
    {
        /// <summary>
        /// 拓展方法,生成EXECL
        /// </summary>
        /// <param name="info">EXECL相关信息</param>
        /// <param name="token">用户认证令牌</param>
        /// <param name="guid">导出文件时，前端传入的标识,标记一次导出任务,为此为key,从Session中更新或读取该次导出任务的进度</param>
        /// <returns>Execl路径</returns>
        public static MemoryStream ExportExeclStream(this ExcelInfo info, string token, string guid = null)
        {
            if (!string.IsNullOrEmpty(guid))
            {
                HttpContext.Current.Session[guid] = 0;
            }
            //1.获取列表对应数据
            DataTable dt = GetGirdData(info, token);
            if (!string.IsNullOrEmpty(guid))
            {
                HttpContext.Current.Session[guid] = 30;//查询出结果,直接设置为30%
            }

            //2.创建文档
            MemoryStream ms = null;
            switch (info.FileFormat)
            {
                case ExportFileFormat.Excel:
                    //导出Excel
                    ms = NPOIHelper.Export(dt, info, guid);
                    break;
                case ExportFileFormat.Word:
                    //导出Word
                    ms = WordHelper.Export(dt, info, guid);
                    break;
                case ExportFileFormat.Pdf:
                    //导出Pdf
                    ms = PdfHelper.Export(dt, info, guid);
                    break;
            }
            return ms;
        }


        /// <summary>
        /// 从WebAPI中获取列表数据
        /// </summary>
        /// <param name="token">用户认证令牌</param>
        /// <returns></returns>
        private static DataTable GetGirdData(ExcelInfo info, string token)
        {
            if (info.IsExportSelectData)
            {
                if (info.Data == null)
                {
                    info.Data = new DataTable();
                }
                return info.Data;
            }
            try
            {
                using (var httpClient = new HttpClient())
                {
                    httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", token);
                    if (info.Type.Equals(HttpMethod.Post.ToString(), StringComparison.OrdinalIgnoreCase))
                    {
                        var responseJson = httpClient.PostAsync(info.Api, info.Condition,
                            new System.Net.Http.Formatting.JsonMediaTypeFormatter()).Result.Content.ReadAsAsync<ExcelApiResult>().Result;
                        if (!responseJson.HasError && responseJson.Message != "logout")
                        {
                            return info.ConvertDataEx2Data(responseJson.Result);
                        }
                        else
                        {
                            DataTable dt = new DataTable();
                            dt.Columns.Add(info.ColumnInfoList[0].Field);
                            DataRow dr = dt.NewRow();
                            //接口报错
                            if (responseJson.HasError)
                            {
                                dr[0] = responseJson.Message;
                            }
                            if (responseJson.Message == "logout")
                            {
                                dr[0] = "登录超时,请刷新页面重试";
                            }
                            dt.Rows.Add(dr);
                            return dt;
                        }
                    }
                    else
                    {
                        throw new ArgumentOutOfRangeException("不支持Get协议获取数据");
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
