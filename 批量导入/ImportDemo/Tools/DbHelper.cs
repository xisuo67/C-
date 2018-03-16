using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Tools
{
    public class DBHelper
    {
        //声明连接对象
        static SqlConnection conn;


        /// <summary>
        /// 连接属性
        /// </summary>
        public static SqlConnection Conn
        {
            get
            {
                if (conn == null)
                {
                    conn = new SqlConnection("Data Source=(local);Initial Catalog=Test;User ID=sa;password=1");
                    conn.Open();
                    return conn;
                }
                if (conn.State == ConnectionState.Closed)
                {
                    conn.Open();
                    return conn;
                }
                if (conn.State == ConnectionState.Broken)
                {
                    conn.Close();
                    conn.Open();
                    return conn;
                }
                return DBHelper.conn;

            }
        }


        #region 增删改共同工具方法,硬绑定
        /// <summary>
        /// 增删改共同工具方法
        /// </summary>
        /// <param name="sqlStr"></param>
        /// <returns></returns>
        public static int ExecuteNonQuery(string sqlStr)
        {
            SqlCommand cmd = new SqlCommand(sqlStr, DBHelper.Conn);
            return cmd.ExecuteNonQuery();
        }
        #endregion
        #region 增删改共同工具方法,软绑定
        /// <summary>
        /// 增删改共同工具方法
        /// </summary>
        /// <param name="sqlStr"></param>
        /// <returns></returns>                                                                                             
        public static int ExecuteNonQuery(string sqlStr, params SqlParameter[] para)
        {
            SqlCommand cmd = new SqlCommand(sqlStr, DBHelper.Conn);
            cmd.Parameters.AddRange(para);
            return cmd.ExecuteNonQuery();
        }
        #endregion



        #region 得到整体信息,硬绑定
        /// <summary>
        /// 得到整体信息
        /// </summary>
        /// <param name="sqlStr"></param>
        /// <returns></returns>
        public static object ExecuteScalar(string sqlStr)
        {
            SqlCommand cmd = new SqlCommand(sqlStr, DBHelper.Conn);
            return cmd.ExecuteScalar();
        }
        #endregion
        #region 得到整体信息,软绑定
        /// <summary>
        /// 得到整体信息
        /// </summary>
        /// <param name="sqlStr"></param>
        /// <returns></returns>
        public static object ExecuteScalar(string sqlStr, params SqlParameter[] para)
        {
            SqlCommand cmd = new SqlCommand(sqlStr, DBHelper.Conn);
            cmd.Parameters.AddRange(para);
            return cmd.ExecuteScalar();
        }
        #endregion


        #region 得到所有信息，直连式查询,硬绑定
        /// <summary>
        /// 得到所有信息
        /// </summary>
        /// <param name="sqlStr"></param>
        /// <returns></returns>
        public static SqlDataReader ExecuteReader(string sqlStr)
        {
            SqlCommand cmd = new SqlCommand(sqlStr, DBHelper.Conn);
            return cmd.ExecuteReader();
        }
        #endregion
        #region 得到所有信息，直连式查询,软绑定
        /// <summary>
        /// 得到所有信息
        /// </summary>
        /// <param name="sqlStr"></param>
        /// <returns></returns>
        public static SqlDataReader ExecuteReader(string sqlStr, params SqlParameter[] para)
        {
            SqlCommand cmd = new SqlCommand(sqlStr, DBHelper.Conn);
            cmd.Parameters.AddRange(para);
            return cmd.ExecuteReader();
        }
        #endregion


        #region 得到所有信息，断开式查询,硬绑定
        /// <summary>
        /// 得到所有信息
        /// </summary>
        /// <param name="sqlStr"></param>
        /// <returns></returns>
        public static DataSet ExecuteReaderDataSet(string sql)
        {
            SqlDataAdapter sda = new SqlDataAdapter(sql, DBHelper.Conn);
            DataSet ds = new DataSet();
            sda.Fill(ds);
            return ds;
        }
        #endregion
        #region 得到所有信息，断开式查询,软绑定
        /// <summary>
        /// 得到所有信息
        /// </summary>
        /// <param name="sqlStr"></param>
        /// <returns></returns>
        public static DataSet ExecuteReaderDataSet(string sqlStr, params SqlParameter[] para)
        {
            SqlDataAdapter sda = new SqlDataAdapter();
            SqlCommand cmd = new SqlCommand(sqlStr, DBHelper.Conn);
            cmd.Parameters.AddRange(para);
            sda.SelectCommand = cmd;
            DataSet ds = new DataSet();
            sda.Fill(ds);
            return ds;
        }
        #endregion
        //查询，获取DataTable
        public static DataTable GetDataTable(string sqlStr)
        {
            SqlCommand cmd = new SqlCommand(sqlStr, DBHelper.Conn);
            DataTable table = new DataTable();
            SqlDataAdapter dap = new SqlDataAdapter(sqlStr, Conn);
            dap.Fill(table);
            return table;
        }

    }
}
