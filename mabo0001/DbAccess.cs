using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OleDb;
using System.Data;

namespace mabo0001{

    internal class DbAccess
    {
        public static string ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=";
        public static string ErrorMessage = "";
        private static OleDbConnection objConnection = new OleDbConnection();
        private static OleDbCommand objCommand = new OleDbCommand();

        public DbAccess(string path)
        {
            ConnectionString = ConnectionString + path;
        }

        /// <summary>
        /// 打开数据库
        /// </summary>
        /// <returns></returns>
        private static bool OpenConn()
        {
            objConnection.ConnectionString = ConnectionString;
            try
            {
                objConnection.Open();
                return true;
            }
            catch (Exception ex)
            {
                ErrorMessage = ex.ToString();
                return false;
            }
        }



        /// <summary>
        /// 关闭数据库
        /// </summary>
        private static void CloseConn()
        {
            if (objConnection.State == System.Data.ConnectionState.Open)
            {
                try
                {
                    objConnection.Close();
                }
                catch { }
            }
        }

        /// <summary>
        /// 执行数据查询
        /// </summary>
        /// <param name="strSql"></param>
        /// <param name="intCount"></param>
        /// <returns></returns>
        public static bool ExecuteNonQuery(string strSql, out int intCount)
        {
            intCount = -1;
            objCommand.CommandType = System.Data.CommandType.Text;
            objCommand.CommandText = strSql;
            if (OpenConn())
            {
                objCommand.Connection = objConnection;
                intCount = objCommand.ExecuteNonQuery();
                CloseConn();
                return true;
            }
            else
            {
                return false;
            }
        }


        /// <summary>
        /// 执行数据sql
        /// </summary>
        /// <param name="strSql"></param>
        /// <param name="i"></param>
        /// <returns></returns>
        public static bool ExecuteScalar(string strSql, out int i)
        {
            i = -1;

            objCommand.CommandType = System.Data.CommandType.Text;
            objCommand.CommandText = strSql;
            if (OpenConn())
            {
                objCommand.Connection = objConnection;
                i = (int)objCommand.ExecuteScalar();
                CloseConn();
                return true;
            }
            else
            {
                return false;
            }
        }



        /// <summary>
        /// 执行数据sql
        /// </summary>
        /// <param name="strSql"></param>
        /// <param name="str"></param>
        /// <returns></returns>
        public static bool ExecuteScalar(string strSql, out string str)
        {
            str = null;

            objCommand.CommandType = System.Data.CommandType.Text;
            objCommand.CommandText = strSql;
            if (OpenConn())
            {
                objCommand.Connection = objConnection;
                Object o = objCommand.ExecuteScalar();
                if (o is DBNull)
                {
                    str = "";
                }
                else
                {
                    str = o.ToString();
                }
                CloseConn();
                return true;
            }
            else
            {
                return false;
            }
        }


        /// <summary>
        /// 执行sql语句 得到datatable
        /// </summary>
        /// <param name="sql"></param>
        /// <returns></returns>
        public static DataTable GetDataSet(string sql)
        {
            DataSet dataset = new DataSet();

            objCommand.CommandType = CommandType.Text;
            objCommand.CommandText = sql;
            if (OpenConn())
            {
                objCommand.Connection = objConnection;
                OleDbDataAdapter da = new OleDbDataAdapter(objCommand);

                da.Fill(dataset);
                CloseConn();
            }
            return dataset.Tables[0];
        }

    }
}




