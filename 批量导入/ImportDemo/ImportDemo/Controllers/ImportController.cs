﻿using Models;
using System;
using System.Collections.Generic;
using System.Data;
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
            var a = Query();
            return View();
        }
        /// <summary>
        /// 查询学生信息
        /// </summary>
        /// <returns></returns>
        public ApiResult<List<Students>> Query()
        {
            ApiResult<List<Students>> apiResult = new ApiResult<List<Students>>();
            string sql = "select * from students";
            try
            {
                var students = DBHelper.GetDataTable(sql).ToList<Students>();
                apiResult.ErrorCode = 0;
                apiResult.Result = students;
                apiResult.TotalCount = students.Count();
            }
            catch (Exception ex)
            {
                apiResult.ErrorCode = 1001;
                apiResult.Message = ex.ToString();
            }
            return apiResult;
        }

    }
}