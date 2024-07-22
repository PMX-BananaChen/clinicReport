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
    public partial class Form1 : Form
    {
        DataAccess DA = new DataAccess();

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            DataTable List = DA.GetRows("select factory as 產品線,cost_center as 成本中心,dept_no as 部門編號,dept_name as 部門名稱,emp_no as 工號,emp_name as 姓名,emp_indate as 日期,emp_outdate as 類型,emp_title as 診費合計,emp_name as 自費部份 from HREmp.HRDataSync.dbo.v_hr_emp_app where emp_no = '" + "00000000" + "'").Tables[0];
            this.dataGridView1.DataSource = List;
            AutoSizeColumn(this.dataGridView1);
            this.textBox1.Select();
            this.textBox1.Focus();
        }

        /// <summary>
        /// 诊费明细
        /// </summary>
        private void button1_Click(object sender, EventArgs e)
        {
            //清空数据
            if (dataGridView1.DataSource != null)
            {
                DataTable dt = (DataTable)dataGridView1.DataSource;
                dt.Rows.Clear();
                dataGridView1.DataSource = dt;
            }
            else
            {
                dataGridView1.Rows.Clear();
            }

            this.textBox1.Text = textBox1.Text.Trim();
            if (textBox1.Text.Length == 0)
            {
                DataTable List = DA.GetRows("select a.factory as 產品線,a.cost_center as 成本中心,a.dept_no as 部門編號,a.dept_name as 部門名稱,"+
                      "a.emp_no as 工號,a.emp_name as 姓名,convert(char(10),b.treatment_date,120) as 日期,b.treatment_date as 时间," +
                      "b.ailment_type as 類型,b.treatment_cost as 診費合計,b.self_cost as 自費部份 " +
                      " from HREmp.HRDataSync.dbo.v_hr_emp_app a,treatment_list b " +
                      " where a.emp_no = b.emp_id " +
                      " and b.treatment_date between '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + " 00:00:00" +"' and '" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + " 23:59:59" + 
                      "' and a.emp_group_id not in (" + "'2'" + "," + "'7'" + ") " +
                      " order by 7").Tables[0];

                DA.GetRows("select a.factory as 產品線,a.cost_center as 成本中心,a.dept_no as 部門編號,a.dept_name as 部門名稱,"+
                      "a.emp_no as 工號,a.emp_name as 姓名,convert(char(10),b.treatment_date,120) as 日期,b.treatment_date as 时间," +
                      "b.ailment_type as 類型,b.treatment_cost as 診費合計,b.self_cost as 自費部份 " +
                      " from HREmp.HRDataSync.dbo.v_hr_emp_app a,treatment_list b " +
                      " where a.emp_no = b.emp_id " +
                      " and b.treatment_date between '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + " 00:00:00" +"' and '" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + " 23:59:59" + 
                      "' and a.emp_group_id not in (" + "'2'" + "," + "'7'" + ") " +
                      " order by 7").Tables[0];

                this.dataGridView1.DataSource = List;

                //查询汇总金额
                DataTable List1 = DA.GetRows("select sum(b.treatment_cost) as treatment_cost from HREmp.HRDataSync.dbo.v_hr_emp_app a,treatment_list b" +
                     " where a.emp_no = b.emp_id " +
                     " and b.treatment_date between '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + " 00:00:00" + "' and '" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + " 23:59:59" +
                     "' and a.emp_group_id not in (" + "'2'" + "," + "'7'" + ")").Tables[0];
                this.textBox3.Text = List1.Rows[0]["treatment_cost"].ToString();
            }
            else
            {
                DataTable List = DA.GetRows("select a.factory as 產品線,a.cost_center as 成本中心,a.dept_no as 部門編號,a.dept_name as 部門名稱," +
                     "a.emp_no as 工號,a.emp_name as 姓名,convert(char(10),b.treatment_date,120) as 日期,b.treatment_date as 时间," +
                     "b.ailment_type as 類型,b.treatment_cost as 診費合計,b.self_cost as 自費部份 " +
                     " from HREmp.HRDataSync.dbo.v_hr_emp_app a,treatment_list b " +
                     " where a.emp_no = b.emp_id " +
                     " and a.emp_no ='" + textBox1.Text +
                     "' and b.treatment_date between '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + " 00:00:00" + "' and '" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + " 23:59:59" +
                     "' and a.emp_group_id not in (" + "'2'" + "," + "'7'" + ") " +
                     " order by 7").Tables[0];
                this.dataGridView1.DataSource = List;

                //查询汇总金额
                DataTable List1 = DA.GetRows("select sum(b.treatment_cost) as treatment_cost from HREmp.HRDataSync.dbo.v_hr_emp_app a,treatment_list b" +
                     " where a.emp_no = b.emp_id " +
                     " and a.emp_no ='" + textBox1.Text +
                     "' and b.treatment_date between '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + " 00:00:00" + "' and '" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + " 23:59:59" +
                     "' and a.emp_group_id not in (" + "'2'" + "," + "'7'" + ")").Tables[0];
                this.textBox3.Text = List1.Rows[0]["treatment_cost"].ToString();
            }
            int i;
            if (dataGridView1.Rows.Count > 0)
            {
                i = dataGridView1.Rows.Count - 1;

            }
            else
            {
                i = dataGridView1.Rows.Count;
            }
            this.textBox2.Text = Convert.ToString(i);
            AutoSizeColumn(this.dataGridView1);
        }

        /// <summary>
        /// 个人汇总
        /// </summary>
        private void button2_Click(object sender, EventArgs e)
        {
            //清空数据
            if (dataGridView1.DataSource != null)
            {
                DataTable dt = (DataTable)dataGridView1.DataSource;
                dt.Rows.Clear();
                dataGridView1.DataSource = dt;
            }
            else
            {
                dataGridView1.Rows.Clear();
            }

            this.textBox1.Text = textBox1.Text.Trim();
            if (textBox1.Text.Length == 0)
            {
                //查询列表
                DataTable List = DA.GetRows("select a.factory as 產品線,a.cost_center as 成本中心,a.dept_no as 部門編號,a.dept_name as 部門名稱," +
                      "a.emp_no as 工號,a.emp_name as 姓名,sum(b.treatment_cost) as 診費合計,sum(b.self_cost) as 自費部份 " +
                      " from HREmp.HRDataSync.dbo.v_hr_emp_app a,treatment_list b " +
                      " where a.emp_no = b.emp_id " +
                      " and b.treatment_date between '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + " 00:00:00" + "' and '" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + " 23:59:59" +
                      "' and a.emp_group_id not in (" + "'2'" + "," + "'7'" + ") " +
                      "group by a.factory,a.cost_center,a.dept_no,a.dept_name,a.emp_no,a.emp_name" +
                      " order by 1,2,3,5").Tables[0];

                this.dataGridView1.DataSource = List;

                //查询汇总金额
                DataTable List1 = DA.GetRows("select sum(b.treatment_cost) as treatment_cost from HREmp.HRDataSync.dbo.v_hr_emp_app a,treatment_list b" +
                     " where a.emp_no = b.emp_id " +
                     " and b.treatment_date between '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + " 00:00:00" + "' and '" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + " 23:59:59" +
                     "' and a.emp_group_id not in (" + "'2'" + "," + "'7'" + ")").Tables[0];
                this.textBox3.Text = List1.Rows[0]["treatment_cost"].ToString();
            }
            else
            {
                DataTable List = DA.GetRows("select a.factory as 產品線,a.cost_center as 成本中心,a.dept_no as 部門編號,a.dept_name as 部門名稱," +
                      "a.emp_no as 工號,a.emp_name as 姓名,sum(b.treatment_cost) as 診費合計,sum(b.self_cost) as 自費部份 " +
                      " from HREmp.HRDataSync.dbo.v_hr_emp_app a,treatment_list b " +
                      " where a.emp_no = b.emp_id " +
                      " and a.emp_no ='" + textBox1.Text +
                      "'and b.treatment_date between '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + " 00:00:00" + "' and '" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + " 23:59:59" +
                      "'and a.emp_group_id not in (" + "'2'" + "," + "'7'" + ") " +
                      " group by a.factory,a.cost_center,a.dept_no,a.dept_name,a.emp_no,a.emp_name" +
                      " order by 1,2,3,5").Tables[0];

                this.dataGridView1.DataSource = List;

                //查询汇总金额
                DataTable List1 = DA.GetRows("select sum(b.treatment_cost) as treatment_cost from HREmp.HRDataSync.dbo.v_hr_emp_app a,treatment_list b" +
                     " where a.emp_no = b.emp_id " +
                     " and a.emp_no ='" + textBox1.Text +
                     "' and b.treatment_date between '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + " 00:00:00" + "' and '" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + " 23:59:59" +
                     "' and a.emp_group_id not in (" + "'2'" + "," + "'7'" + ")").Tables[0];
                this.textBox3.Text = List1.Rows[0]["treatment_cost"].ToString();
            }
            int i;
            if (dataGridView1.Rows.Count > 0)
            {
                i = dataGridView1.Rows.Count - 1;

            }
            else
            {
                i = dataGridView1.Rows.Count;
            }
            this.textBox2.Text = Convert.ToString(i);
            AutoSizeColumn(this.dataGridView1);
        }

        /// <summary>
        /// 成本中心汇总
        /// </summary>
        private void button3_Click(object sender, EventArgs e)
        {
            //查询列表
            DataTable List = DA.GetRows("select a.factory as 產品線,a.cost_center as 成本中心,sum(b.treatment_cost) as 診費合計,sum(b.self_cost) as 自費部份" +
                  " from HREmp.HRDataSync.dbo.v_hr_emp_app a,treatment_list b " +
                  " where a.emp_no = b.emp_id " +
                  " and b.treatment_date between '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + " 00:00:00" + "' and '" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + " 23:59:59" +
                  "' and a.emp_group_id not in (" + "'2'" + "," + "'7'" + ") " +
                  "group by a.factory,a.cost_center" +
                  " order by 1,2").Tables[0];

            this.dataGridView1.DataSource = List;
            int i;
            if (dataGridView1.Rows.Count > 0)
            {
                i = dataGridView1.Rows.Count - 1;

            }
            else
            {
                i = dataGridView1.Rows.Count;
            }
            this.textBox2.Text = Convert.ToString(i);

            //查询汇总金额
            DataTable List1 = DA.GetRows("select sum(b.treatment_cost) as treatment_cost from HREmp.HRDataSync.dbo.v_hr_emp_app a,treatment_list b" +
                 " where a.emp_no = b.emp_id " +
                 " and b.treatment_date between '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + " 00:00:00" + "' and '" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + " 23:59:59" +
                 "' and a.emp_group_id not in (" + "'2'" + "," + "'7'" + ")").Tables[0];
            this.textBox3.Text = List1.Rows[0]["treatment_cost"].ToString();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string fileName = "";
            string saveFileName = "";
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.DefaultExt = "xlsx";
            saveDialog.Filter = "Excel File(*.xlsx)|*.xlsx";
            //saveDialog.DefaultExt = "xls";
            //saveDialog.Filter = "Excel文件|*.xls";
            saveDialog.FileName = fileName;
            saveDialog.ShowDialog();
            saveFileName = saveDialog.FileName;
            if (saveFileName.IndexOf(":") < 0) return; //被点了取消

            //int FormatNum;//保存excel文件的格式
            //string Version;//excel版本号
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            if (xlApp == null)
            {
                MessageBox.Show("无法创建Excel对象，您的电脑可能未安装Excel");
                return;
            }
            Microsoft.Office.Interop.Excel.Workbooks workbooks = xlApp.Workbooks;
            Microsoft.Office.Interop.Excel.Workbook workbook = workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];//取得sheet1 
            //Version = xlApp.Version;
            //if (Convert.ToDouble(Version) < 12)//You use Excel 97-2003
            //{
            //    FormatNum = -4143;
            //}
            //else//you use excel 2007 or later
            //{
            //    FormatNum = 56;
            //}
            //写入标题             
            for (int i = 0; i < this.dataGridView1.ColumnCount; i++)
            { worksheet.Cells[1, i + 1] = this.dataGridView1.Columns[i].HeaderText; }
            //写入数值
            for (int r = 0; r < this.dataGridView1.Rows.Count; r++)
            {
                for (int i = 0; i < this.dataGridView1.ColumnCount; i++)
                {
                    worksheet.Cells[r + 2, i + 1] = this.dataGridView1.Rows[r].Cells[i].Value;
                }
                System.Windows.Forms.Application.DoEvents();
            }
            worksheet.Columns.EntireColumn.AutoFit();//列宽自适应
            MessageBox.Show(fileName + "资料保存成功", "提示", MessageBoxButtons.OK);
            if (saveFileName != "")
            {
                try
                {
                    workbook.Saved = true;
                    workbook.SaveCopyAs((saveFileName));  //fileSaved = true;                 
                }
                catch (Exception ex)
                {//fileSaved = false;                      
                    MessageBox.Show("导出文件时出错,文件可能正被打开！\n" + ex.Message);
                }
            }
            xlApp.Quit();
            GC.Collect();//强行销毁           }
        }

        //补录诊费
        private void button6_Click(object sender, EventArgs e)
        {
            using (Form2 dlg = new Form2()) //caozuo是窗口类名，确保访问；后面的是构造函数
            {
                dlg.ShowDialog();
            }
        }

        //双击行修改诊费
        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            using (Form3 f3 = new Form3(dataGridView1.CurrentRow))
            {
                f3.ShowDialog();          //显示窗体
            }
                
        }

        //退出程序
        private void button5_Click(object sender, EventArgs e)
        {
            System.Environment.Exit(0);
        }

        /// <summary>
        /// 使DataGridView的列自适应宽度
        /// </summary>
        /// <param name="dgViewFiles"></param>
        private void AutoSizeColumn(DataGridView dgViewFiles)
        {
            int width = 0;
            //使列自使用宽度
            //对于DataGridView的每一个列都调整
            for (int i = 0; i < dgViewFiles.Columns.Count; i++)
            {
                //将每一列都调整为自动适应模式
                dgViewFiles.AutoResizeColumn(i, DataGridViewAutoSizeColumnMode.AllCells);
                //记录整个DataGridView的宽度
                width += dgViewFiles.Columns[i].Width;
            }
            //判断调整后的宽度与原来设定的宽度的关系，如果是调整后的宽度大于原来设定的宽度，
            //则将DataGridView的列自动调整模式设置为显示的列即可，
            //如果是小于原来设定的宽度，将模式改为填充。
            if (width > dgViewFiles.Size.Width)
            {
                dgViewFiles.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
            }
            else
            {
                dgViewFiles.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }
            //冻结某列 从左开始 0，1，2
            dgViewFiles.Columns[1].Frozen = true;
        }
        
    }
}
