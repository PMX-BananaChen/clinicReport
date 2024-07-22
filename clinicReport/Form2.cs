using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace clinicReport
{
    public partial class Form2 : Form
    {
        DataAccess DA = new DataAccess();

        public Form2()
        {
            InitializeComponent();
        }
        private void Form2_Load(object sender, EventArgs e)
        {
            DataTable List = DA.GetRows("SELECT Ailment_Type FROM dbo.Treatment_Ailment ORDER BY Ailment_ID").Tables[0];
            foreach (DataRow ailment in List.Rows)
            {
                comboBox1.Items.Add(ailment[0]);
            }
            comboBox1.Text = comboBox1.Items[0].ToString();
        }

        //查询员工信息
        private void button1_Click(object sender, EventArgs e)
        {
            String dayFlag = DateTime.Now.ToString("yyyyMMdd");
            DataTable empInfo = DA.GetRows("SELECT * FROM HREmp.HRDataSync.dbo.v_hr_emp_app WHERE Emp_No ='" + textBox1.Text + "' AND Emp_OutDate >'" + dayFlag + "'").Tables[0];
            //姓名-部门-身份证号
            textBox2.Text = empInfo.Rows[0]["Emp_Name"].ToString();
            textBox3.Text = empInfo.Rows[0]["Dept_No"].ToString() + ";" + empInfo.Rows[0]["Factory"].ToString() + " / " + empInfo.Rows[0]["Dept_Name"].ToString();
            textBox5.Text = empInfo.Rows[0]["Emp_Serial_ID"].ToString();
        }

       
        //监听诊费
        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            //判断按键是不是要输入的类型。
            if (((int)e.KeyChar < 48 || (int)e.KeyChar > 57) && (int)e.KeyChar != 8 && (int)e.KeyChar != 46)
                e.Handled = true;

            //小数点的处理。
            if ((int)e.KeyChar == 46)                           //小数点
            {
                if (this.textBox6.Text.Length <= 0)
                    e.Handled = true;   //小数点不能在第一位
                else
                {
                    float f;
                    float oldf;
                    bool b1 = false, b2 = false;
                    b1 = float.TryParse(textBox6.Text, out oldf);
                    b2 = float.TryParse(textBox6.Text + e.KeyChar.ToString(), out f);
                    if (b2 == false)
                    {
                        if (b1 == true)
                            e.Handled = true;
                        else
                            e.Handled = false;
                    }
                }
            }
        }
        //诊费计算
        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            decimal result;
            if (decimal.TryParse(textBox6.Text, out result))
            {
                if (result > 0)
                {
                    decimal companyCost = result / 10;
                    decimal empCost = companyCost * 3;
                    textBox7.Text = empCost.ToString();
                }
            }
            if (textBox6.Text.Length == 0)
            {
                textBox7.Clear();
            }
        }

        //补录查询员工信息键盘回车事件
        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            textBox1.Text = textBox1.Text.Trim();
            if (e.KeyCode == Keys.Enter)
            {
                String dayFlag = DateTime.Now.ToString("yyyyMMdd");
                DataTable empInfo = DA.GetRows("SELECT * FROM HREmp.HRDataSync.dbo.v_hr_emp_app WHERE Emp_No ='" + textBox1.Text + "' AND Emp_OutDate >'" + dayFlag + "'").Tables[0];
                //工号
                //textBox1.Text = empInfo.Rows[0]["Emp_No"].ToString();
                //姓名-部门-身份证号
                textBox2.Text = empInfo.Rows[0]["Emp_Name"].ToString();
                textBox3.Text = empInfo.Rows[0]["Dept_No"].ToString() + ";" + empInfo.Rows[0]["Factory"].ToString() + " / " + empInfo.Rows[0]["Dept_Name"].ToString();
                textBox5.Text = empInfo.Rows[0]["Emp_Serial_ID"].ToString();
            }
        }

        //保存补录
        private void button2_Click(object sender, EventArgs e)
        {
            bool b = true;

            //String dayFlag = DateTime.Now.ToString("yyyyMMdd");
            String dayFlag = dateTimePicker1.Value.ToString("yyyyMMdd");
            String Treatment_ID = dayFlag + "000";
            string DEPT_ID = textBox3.Text.Substring(0, textBox3.Text.IndexOf(";"));
            string DEPT_NM = textBox3.Text.Substring(textBox3.Text.IndexOf(";") + 1);
            string EMP_ID = textBox1.Text.ToString();
            string EMP_NM = textBox2.Text.ToString();
            EMP_NM = ToTraditionalChinese(EMP_NM);
            string EMP_IC_ID = textBox5.Text.ToString();
            string Ailment_Type = comboBox1.Text.ToString();
            decimal Treatment_Cost = Convert.ToDecimal(textBox6.Text);
            decimal Self_Cost = Convert.ToDecimal(textBox7.Text);
            string Treatment_Type = "Unlimited";//排队无限制
            
            if (Treatment_Cost > Convert.ToDecimal(200))
            {
                DialogResult r = MessageBox.Show("诊疗费超过200元,确定保存吗?", "提示", MessageBoxButtons.OKCancel);
                if (r == DialogResult.Cancel)
                {
                    b = false;
                }
            }

            if (b)
            {
                //插入就诊记录表
                DA.exec_sql("INSERT INTO dbo.Treatment_List (Treatment_ID,DEPT_ID,DEPT_NM,EMP_ID,EMP_NM,EMP_IC_ID,Treatment_Date,Ailment_Type,Treatment_Type,Treatment_Cost,Self_Cost) VALUES ('" + Treatment_ID + "','" + DEPT_ID + "','" + DEPT_NM + "','" + EMP_ID + "',N'" + EMP_NM + "','" + EMP_IC_ID + "','" + dateTimePicker1.Value.ToString("yyyy-MM-dd HH:mm:ss:fff") + "','" + Ailment_Type + "','" + Treatment_Type + "'," + Treatment_Cost + "," + Self_Cost + ")");
                //更新排队就诊表
                //DA.exec_sql("UPDATE dbo.Treatment_Wait_List SET FLAG = 'Y',STATUS = 'Treatment End',Treatment_End_Time = '" +
                //                                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss:fff") +
                //                                    "' WHERE DayFlag ='" + dayFlag + "'AND Emp_ID ='" + EMP_ID +
                //                                    "' AND FLAG <> 'Y' ");

                textBox1.ReadOnly = false;
                textBox1.Focus();
                textBox2.Clear();
                textBox3.Clear();
                textBox5.Clear();
                textBox6.Clear();
                textBox7.Clear();
                comboBox1.Text = comboBox1.Items[0].ToString();
            }
        }

        //退出
        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //简体转繁体
        //解决Treatment_Wait_List表的插入简体中文乱码
        public static string ToTraditionalChinese(string strSimple)
        {
            string strTraditional = Microsoft.VisualBasic.Strings.StrConv(strSimple, Microsoft.VisualBasic.VbStrConv.TraditionalChinese, 1033);
            return strTraditional;
        }
        //繁体转简体
        public static string ToSimpleChinese(string strTraditional)
        {
            string strSimple = Microsoft.VisualBasic.Strings.StrConv(strTraditional, Microsoft.VisualBasic.VbStrConv.SimplifiedChinese, 1033);
            return strSimple;
        }
    }
}
