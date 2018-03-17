using Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Tools;
using Utility;

namespace ImportDemo.Controllers
{
    public class ImportController : Controller
    {
        /// <summary>
        /// 系统所有批量导入业务
        /// </summary>
        [ImportMany(typeof(ExcelImport))]
        public IEnumerable<ExcelImport> AllImports { get; set; }
        // GET: Import
        public ActionResult Index()
        {
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
        public ActionResult Export(string excelParam)
        {
            ExcelInfo info = JsonConvert.DeserializeObject<ExcelInfo>(excelParam);
            if (info.Data == null)
            {
                string sql = "select * from students";
                info.Data = DBHelper.GetDataTable(sql);
            }
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
            string token = string.Empty;

            MemoryStream ms = info.ExportExeclStream(info.Guid);

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
        /// <summary>
        /// 导出Excel模版
        /// </summary>
        /// <returns></returns>
        public ActionResult ImportTemplate()
        {
            ImportResult result = new ImportResult();
            try
            {
                if (AllImports == null)
                {
                    throw new ArgumentNullException("系统不存在Excel批量导入业务处理模块");
                }
                string ywType = Request.QueryString["type"];
                if (string.IsNullOrEmpty(ywType))
                {
                    throw new ArgumentNullException("ywType");
                }
                //业务类型
                ExcelImportType type;
                int iType;
                if (int.TryParse(ywType, out iType))
                {
                    type = EnumHelper.IntToEnum<ExcelImportType>(iType);
                }
                else
                {
                    type = EnumHelper.StringToEnum<ExcelImportType>(ywType);
                }
                //文件
                HttpPostedFileBase file = Request.Files["file"];

                var handler = AllImports.FirstOrDefault(e => e.Type == type);
                if (handler == null)
                {
                    throw new Exception("未找到“" + type.ToString() + "”相应处理模块");
                }

                result = handler.ImportTemplate(file.InputStream, file.FileName);
                if (result.IsSuccess)
                {
                    //是否获取详细数据，决定后台是否返回 result.ExtraInfo
                    string ReturnDetailData = Request.QueryString["ReturnDetailData"];
                    if (string.IsNullOrEmpty(ReturnDetailData) || ReturnDetailData != "1")
                    {
                        result.ExtraInfo = null;
                    }
                }
                else
                {
                    //设置错误模版http路径
                    result.Message = "http://" + Request.Url.Authority + result.Message;
                }
            }
            catch (Exception ex)
            {
                result.IsSuccess = false;
                result.Message = ex.Message;
            }
            return Content(JsonConvert.SerializeObject(result));
        }

        public void DownLoadTemplate(ExcelImportType type)
        {
            if (AllImports == null)
            {
                throw new ArgumentNullException("系统不存在Excel批量导入业务处理模块");
            }
            var handler = AllImports.FirstOrDefault(e => e.Type == type);
            if (handler == null)
            {
                throw new Exception("未找到“" + type.ToString() + "”相应处理模块");
            }

            string path = ExcelImporMapper.GetTemplatePath(type);
            if (System.IO.File.Exists(path))
            {
                try
                {
                    string FileName = Path.GetFileName(path);
                    Response.Headers.Add("Content-Disposition", string.Format("attachment;filename={0}", FileName));
                    Response.ContentType = MimeHelper.GetMineType(path);
                    handler.GetExportTemplate(path, Response.OutputStream);
                    Response.OutputStream.Flush();
                    Response.End();
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
            else
            {
                throw new Exception("未找到“" + type.ToString() + "”对应模版文件");
            }
        }
    }
}