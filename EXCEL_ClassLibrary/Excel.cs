using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EXCEL_ClassLibrary
{
    public class Excel
    {
        //获取数据库
        public static DataTable GetDataTable(string sql,string path)
        {
            //构建连接数据库的字符串       
            string sconnectstring = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + path + ";" + "Extended Properties='Excel 8.0;HDR=Yes;IMEX=0'";
            // string SConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + path + ";" + "Extended Properties='Excel 8.0;HDR=Yes;IMEX=0'";
            //IMEX=0 为汇出模式，这个模式Excle只能用作"写入"用途
            //IMEX=1 为汇入模式，这个模式Excle只能用作"读取"用途
            //IMEX=2 为链接模式, 这个模式Excle同时支持"读写"用途
            //HDR=Yes 创建表头//Excel 8.0代表版本。2003以上可以用8.0，低于则用7.0
            //连接数据库
            using (OleDbConnection ole_cnn = new OleDbConnection(sconnectstring))
            {
                //打开数据库-Access
                ole_cnn.Open();
                //创建操作对象
                using (OleDbCommand ole_cmd = ole_cnn.CreateCommand())
                {
                    //执行SQL语句
                    ole_cmd.CommandText = sql;
                    using (OleDbDataAdapter dapter = new OleDbDataAdapter(ole_cmd))
                    {
                        DataSet dr = new DataSet();
                        dapter.Fill(dr);
                        return dr.Tables[0];
                    }
                }
            }
           

        }
        //更新数据
        public static  int Upatate( string sql,string path)
        {
            //构建连接数据库的字符串
            //string SConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + path + ";" + "Extended Properties='Excel 8.0;HDR=Yes;IMEX=0'";          
            string sconnectstring = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + path + ";" + "Extended Properties='Excel 8.0;HDR=Yes;IMEX=0'";
            //IMEX=0 为汇出模式，这个模式Excle只能用作"写入"用途
            //IMEX=1 为汇入模式，这个模式Excle只能用作"读取"用途
            //IMEX=2 为链接模式, 这个模式Excle同时支持"读写"用途
            //HDR=Yes 创建表头
            //Excel 8.0代表版本。2003以上可以用8.0，低于则用7.0
            //连接数据库
            using (OleDbConnection ole_cnn = new OleDbConnection(sconnectstring))
            {
                //打开数据库-Access
                ole_cnn.Open();
                //创建操作对象
                using (OleDbCommand ole_cmd = ole_cnn.CreateCommand())
                {
                    //执行SQL语句
                    ole_cmd.CommandText = sql;
                    return ole_cmd.ExecuteNonQuery();
                }
            }

        }

    }
}
