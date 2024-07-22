using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace clinicReport
{
    public class DataAccess
    {
        public DataAccess()
        {
        }

        #region 配置数据库连接字符串
        /// <summary>
        /// 配置数据库连接字符串
        /// </summary>

        string connstr = ConfigurationManager.ConnectionStrings["str"].ConnectionString;

        #endregion

        //DataSet 门诊库获取数据
        public DataSet GetRows(string sql)
        {
            SqlConnection con = new SqlConnection(connstr);
            con.Open();
            SqlDataAdapter sda = new SqlDataAdapter(sql, con);
            DataSet ds = new DataSet();
            sda.Fill(ds);
            con.Close();
            return ds;
        }

        //DataSet 获取数据
        public void exec_sql(string sql)
        {
            SqlConnection con = new SqlConnection(connstr);
            con.Open();
            SqlCommand mycomm = new SqlCommand(sql, con);
            mycomm.ExecuteNonQuery();
            con.Close();
        }
    }
}
