using Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Tools;

namespace ImportDemo.Controllers
{
    public class ImportController : Controller
    {
        // GET: Import
        public ActionResult Index()
        {
            //var students = Query();
            //if (students.ErrorCode == 0)
            //{
            //    ViewBag.Data = JsonConvert.SerializeObject(students.Result);
            //}
            //else
            //{
            //    ViewBag.Data = null;
            //}
            return View();
        }
        /// <summary>
        /// 查询学生信息
        /// </summary>
        /// <returns></returns>
        public JsonResult Query()
        {
            ApiResult<List<Students>> apiResult = new ApiResult<List<Students>>();
            string sql = "select * from students";
            try
            {
                apiResult.ErrorCode = 0;
                var students = DBHelper.GetDataTable(sql).ToList<Students>();
                apiResult.Result = students;

            }
            catch (Exception ex)
            {
                apiResult.ErrorCode = 1001;
                apiResult.Message = ex.ToString();
            }
            return Json(apiResult);
        }
        public ActionResult Export()
        {
            string excelParam = "";
            ExcelInfo info = JsonConvert.DeserializeObject<ExcelInfo>(excelParam);
            string fileExt = info.GetFileExt();
            if (string.IsNullOrEmpty(info.FileName))
            {
                info.FileName = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + fileExt;
            }
            else
            {
                if (!info.FileName.EndsWith(fileExt))
                {
                    info.FileName = info.FileName + fileExt;
                }
            }
            //mimeType
            string mineType = MimeHelper.GetMineType(info.FileName);

            if (!info.IsExportSelectData)
            {
                //设置最大导出条数
                info.Condition.PageSize = 9999;
            }
            string token = string.Empty;
            //if (!info.IsExportSelectData)
            //{
            //    UserToken ctoken = TokenHelper.GetImsToken();
            //    if (ctoken != null)
            //    {
            //        token = TokenHelper.CreateToken(ctoken.UserGUID, ctoken.AppKey);
            //    }
            //}

            MemoryStream ms = info.ExportExeclStream(token, info.Guid);

            byte[] msbyte = ms.GetBuffer();
            ms.Dispose();
            ms = null;
            //记录进度
            if (!string.IsNullOrEmpty(info.Guid))
            {
                Session[info.Guid] = 100;
            }

            return File(msbyte, mineType, info.FileName);
        }
    }
}