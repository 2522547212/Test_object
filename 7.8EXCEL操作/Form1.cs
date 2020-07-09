using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using EXCEL_ClassLibrary;

namespace _7._8EXCEL操作
{
   
    public partial class Form1 : Form
    {
        //布尔值->判断是插入还是修改
        bool insert_into;
        public Form1()
        {
            InitializeComponent();
        }
        //********************************************以下问题是动态连接库与UI程序项目不在64位状态下生成的**************************************************
        //未能加载文件或程序集“EXCEL_ClassLibrary, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null”或它的某一个依赖项。试图加载格式不正确的程序
        //未在本地计算机上注册“Microsoft.ACE.OLEDB.12.0”提供程序
        //********************************************在属性-生成-选择X64-重新生成一下***********************************************************************
        //查询
        private void button5_Click(object sender, EventArgs e)
        {
            if (this.textBox1.Text=="")
            {
                var filePath = "./Excel表格.xls";
                string sql = "select 学号,姓名, 班级,电话号码 from[学生信息$]where 状态 = '正常'";
                this.DGV.DataSource = Excel.GetDataTable(sql, filePath);
            }
            else
            {
                var filePath = "Excel表格.xls";
                string sql = "select 学号,姓名,班级,电话号码 from[学生信息$] where 状态='正常' "
                 + " and " + "学号=" + this.textBox1.Text;              
                //注意：and的前后要空格
                this.DGV.DataSource = Excel.GetDataTable(sql, filePath);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var filePath = "Excel表格.xls";
            string sql = "CREATE TABLE 学生信息([学号]INT,[姓名]VarChar,[班级]VarChar,[电话号码]VarChar,[状态]VarChar)";          
            //调用更新数据库
            Excel.Upatate(sql, filePath);//        
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var filePath = "./Excel表格.xls";
            string sql;
            if (insert_into)
            {              
                //SQL语句
                sql = "insert into [学生信息$](学号,姓名,班级,电话号码,状态) values({0},'{1}','{2}','{3}','{4}')";
                sql = string.Format(sql, this.textBox5.Text, this.textBox2.Text, this.textBox3.Text, this.textBox4.Text, "正常");              
            }
            else
            {
                //SQL语句
                sql = "update [学生信息$] set 姓名='{0}',班级='{1}',电话号码='{2}',状态='正常' where 学号={3}";
                sql = string.Format(sql,  this.textBox2.Text, this.textBox3.Text, this.textBox4.Text, this.textBox1.Text, "正常");
            }
            this.DGV.DataSource = Excel.Upatate(sql, filePath);
            //查询一下      
            string sql1 = "select 学号,姓名,班级,电话号码 from[学生信息$] where 状态='正常' ";                  
            this.DGV.DataSource = Excel.GetDataTable(sql1, filePath);
        }    
        //插入
        private void button4_Click(object sender, EventArgs e)
        {
            this.groupBox3.Enabled = true;
            insert_into = true;
        }    
        private void Form1_Load(object sender, EventArgs e)
        {
            this.groupBox3.Enabled = false;
        }
        //修改
        private void button6_Click(object sender, EventArgs e)
        {
            this.groupBox3.Enabled = true;
            insert_into = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {

            this.textBox2.Clear();
            this.textBox3.Clear();
            this.textBox4.Clear();
            this.textBox5.Clear();

           

        }

        private void button7_Click(object sender, EventArgs e)
        {
            var filePath = "./Excel表格.xls";
            string sql;
            //SQL语句
            sql = "update [学生信息$] set 状态='删除' where 学号={0}";
            sql = string.Format(sql, this.textBox1.Text);
            this.DGV.DataSource = Excel.Upatate(sql, filePath);
            //查询一下      
            string sql1 = "select 学号,姓名,班级,电话号码 from[学生信息$] where 状态='正常' ";
            this.DGV.DataSource = Excel.GetDataTable(sql1, filePath);
        }
    }
}
