using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Models
{
    public class ApiResult<TResult>
    {
        public bool HasError { get; set; }

        private int _errorCode;

        /// <summary>
        /// 错误码，不同的 api 接口自己定义，调用方需要根据具体接口的 ErrorCode 来进行处理，后期会定义一系列标准错误码.
        /// 
        /// 通用错误码
        /// 错误码         错误信息
        /// 
        /// </summary>
        public int ErrorCode
        {
            get { return this._errorCode; }
            set
            {
                this._errorCode = value;
                if (value == 0)
                {
                    this.HasError = false;
                }
                else
                {
                    this.HasError = true;
                }

            }
        }
        private string _message;

        /// <summary>
        /// 执行返回消息
        /// </summary>
        /// <remarks></remarks>
        public string Message
        {
            get
            {
                if (this._message == null)
                {
                    switch (this.ErrorCode)
                    {
                        case -1: return "未知错误";
                        case 0: return "成功";
                        case 1: return "参数错误";
                        case 2: return "参数不能为空";
                        case 1001: return "连接数据库失败";
                    }
                }
                return this._message;
            }
            set { this._message = value; }
        }

        /// <summary>
        /// 异常信息
        /// </summary>
        public Exception Exception { get; set; }

        /// <summary>
        /// 返回的主要内容.
        /// </summary>
        public TResult Result { get; set; }

        /// <summary>
        /// 数据总条数
        /// </summary>
        public int? TotalCount { get; set; }

        /// <summary>
        /// 数据总页数
        /// </summary>
        public int? TotalPage { get; set; }

        public ApiResult()
        {
            Result = default(TResult);
            //HasError = false;
        }

    }
}
