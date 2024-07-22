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
    
    public partial class Form3 : Form
    {
        DataAccess DA = new DataAccess();

        private DataGridViewRow currentRow1;

        public Form3()
        {
            
        }

        public Form3(DataGridViewRow currentRow)
        {
            InitializeComponent();
            this.currentRow1 = currentRow;
            getValue();
        }

        public void getValue()                           //给控件赋值
        {
                textBox1.Text = Convert.ToString(currentRow1.Cells[4].Value);
                textBox2.Text = Convert.ToString(currentRow1.Cells[5].Value);
                textBox3.Text = Convert.ToString(currentRow1.Cells[6].Value);
                textBox4.Text = Convert.ToString(currentRow1.Cells[8].Value);
                textBox5.Text = Convert.ToString(currentRow1.Cells[9].Value);
            this.textBox4.Select();
            this.textBox4.Focus();
        }

        //监听诊费
        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            //判断按键是不是要输入的类型。
            if (((int)e.KeyChar < 48 || (int)e.KeyChar > 57) && (int)e.KeyChar != 8 && (int)e.KeyChar != 46)
                e.Handled = true;

            //小数点的处理。
            if ((int)e.KeyChar == 46)                           //小数点
            {
                if (this.textBox4.Text.Length <= 0)
                    e.Handled = true;   //小数点不能在第一位
                else
                {
                    float f;
                    float oldf;
                    bool b1 = false, b2 = false;
                    b1 = float.TryParse(textBox4.Text, out oldf);
                    b2 = float.TryParse(textBox4.Text + e.KeyChar.ToString(), out f);
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
        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            decimal result;
            if (decimal.TryParse(textBox4.Text, out result))
            {
                if (result > 0)
                {
                    decimal companyCost = result / 10;
                    decimal empCost = companyCost * 3;
                    textBox5.Text = empCost.ToString();
                }
            }
            if (textBox4.Text.Length == 0)
            {
                textBox5.Clear();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DA.exec_sql("UPDATE dbo.Treatment_List SET Treatment_Cost = '" +
                                                textBox4.Text + "',Self_Cost ='" + textBox5.Text + "'WHERE EMP_ID ='" + textBox1.Text + "' AND EMP_NM = '" + textBox2.Text + "' AND Treatment_Date between '" + textBox3.Text + " 00:00:00" + "' and '" + textBox3.Text + " 23:59:59'");


            //UPDATE dbo.Treatment_List SET Treatment_Cost = '" +
            //                                    textBox4.Text + "' AND Self_Cost ='" + textBox5.Text + "'
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
