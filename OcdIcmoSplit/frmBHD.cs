using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.ApplicationBlocks.Data;
using System.Collections;
using System.IO;
using System.Data.OleDb;
using System.Reflection;
using System.Threading;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.CSharp.RuntimeBinder;
using System.Diagnostics;

namespace OcdIcmoSplit
{
    public partial class frmBHD : Form
    {
        [DllImport("user32.dll", SetLastError = true)]
        private static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

        public void KillSpecialExcel(Microsoft.Office.Interop.Excel.Application m_objExcel)
        {
            try
            {
                if (m_objExcel != null)
                {
                    int processId;
                    frmBHD.GetWindowThreadProcessId(new IntPtr(m_objExcel.Hwnd), out processId);
                    Process.GetProcessById(processId).Kill();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        [CompilerGenerated]
        private static class o__SiteContainer0
        {
            public static CallSite<Func<CallSite, object, Microsoft.Office.Interop.Excel.Worksheet>> p__Site1;
        }

        // 数据保存
        DataTable myTable = new DataTable();
       // List<String> list1 = new List<string>();

        //string WorkShop;
        //DateTime beginDate;
        //DateTime endDate;


        public frmBHD(string myWorkShop, DateTime mybeginDate, DateTime myendDate)
        {
            InitializeComponent();
            this.comboBox1.Text = myWorkShop;
            this.dateTimePicker1.Value = mybeginDate;
            this.dateTimePicker2.Value = myendDate;
        }

        private void frmBHD_Load(object sender, EventArgs e)
        {
            InsertComShopID();
            LoadData();

        }

        private void LoadData()
        {
            System.TimeSpan t3 = this.dateTimePicker2.Value - this.dateTimePicker1.Value;  //两个时间相减 。默认得到的是 两个时间之间的天数   得到：365.00:00:00 

            double getDay = t3.TotalDays;

            if (getDay < 50)
            {

                System.Configuration.Configuration config = System.Configuration.ConfigurationManager.OpenExeConfiguration(System.Configuration.ConfigurationUserLevel.None);
                string strConn = config.AppSettings.Settings["connectionstring"].Value;//this.txtUrl.Text.Trim();
                DataTable dt = SqlHelper.ExecuteDataset(strConn, "pr_xc_StatPC", this.dateTimePicker1.Value.ToShortDateString(), this.dateTimePicker2.Value.ToShortDateString(), this.comboBox1.Text,this.textBoxPlaner.Text.Trim(),this.textBoxItemNumber.Text.Trim()).Tables[0];
                myTable = dt;
                this.dataGridView1.DataSource = dt;
            }
            else
            {
                MessageBox.Show( "为了保证查看统计的速度，统计只能看50天的数据","金蝶提示");
            }
            
        }

        public void InsertComShopID()
        {
            System.Configuration.Configuration config = System.Configuration.ConfigurationManager.OpenExeConfiguration(System.Configuration.ConfigurationUserLevel.None);
            string strConn = config.AppSettings.Settings["connectionstring"].Value;//this.txtUrl.Text.Trim();

            this.comboBox1.Items.Clear();//清空ComBox
            IDataReader dr = SqlHelper.ExecuteReader(strConn, CommandType.StoredProcedure, "pr_xc_getWorkShop");
            while (dr.Read())
            {
                this.comboBox1.Items.Add(dr[0].ToString());//循环读取数据
            }//end block while

            dr.Close();//  关闭数据集
            // DB.GetColse();//关闭数据库连接
        }

        private void buttonSerach_Click(object sender, EventArgs e)
        {
            LoadData();
        }

 


        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        

        /// <summary>
        /// 导出Excel
        /// </summary>
        /// <param name="saveFileDialog"></param>
        public void ToExcel(SaveFileDialog saveFileDialog)
        {
            // saveFileDialog = new SaveFileDialog();
            //Thread.Sleep(1000); 
            saveFileDialog.Filter = "Execl files (*.xls)|*.xls";
            saveFileDialog.FilterIndex = 0;
            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.CreatePrompt = false;
            saveFileDialog.Title = "导出文件保存路径";
            System.TimeSpan t3 = this.dateTimePicker2.Value - this.dateTimePicker1.Value;
            double days = t3.TotalDays;

            List<String> list = new List<string>();

            list.Add("产线");
            list.Add("产品编码");
            list.Add("产品");
            list.Add("规格型号");
            for (int i = 0; i < days; i++)
            {
                list.Add(this.dateTimePicker1.Value.AddDays(i).ToShortDateString());
            }

            DialogResult result = saveFileDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                string fileName = saveFileDialog.FileName;
                if (myTable.Rows.Count == 0)
                {
                    return;
                }
                if (fileName.Length != 0)
                {
                    Missing value = Missing.Value;
                    Microsoft.Office.Interop.Excel.Application application = (Microsoft.Office.Interop.Excel.Application)Activator.CreateInstance(Type.GetTypeFromCLSID(new Guid("00024500-0000-0000-C000-000000000046")));
                    application.Visible = false;
                    CultureInfo currentCulture = Thread.CurrentThread.CurrentCulture;
                    Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
                    Microsoft.Office.Interop.Excel.Workbook workbook = application.Workbooks.Add(value);
                    if (frmBHD.o__SiteContainer0.p__Site1 == null)
                    {
                        frmBHD.o__SiteContainer0.p__Site1 = CallSite<Func<CallSite, object, Microsoft.Office.Interop.Excel.Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof(Microsoft.Office.Interop.Excel.Worksheet), typeof(frmBHD)));
                    }
                    Microsoft.Office.Interop.Excel.Worksheet worksheet = frmBHD.o__SiteContainer0.p__Site1.Target(frmBHD.o__SiteContainer0.p__Site1, workbook.Worksheets.Add(value, value, value, value));
                    worksheet.Name = "data";
                    worksheet.Cells[1, 1] = "车间安排统计数据";
                    for (int i = 0; i < list.Count; i++)
                    {
                        worksheet.Cells[2, i + 1] = list[i];
                    }
                    Microsoft.Office.Interop.Excel.Range arg_1F8_0 = worksheet.get_Range("A2", value);
              
                    ((Microsoft.Office.Interop.Excel.Range)worksheet.Rows[3, Missing.Value]).ColumnWidth = 25;
                    Microsoft.Office.Interop.Excel.Range range = null;
                    int count = myTable.Rows.Count;
                    int j = 1;
                    int num = 1;
                    int count2 = myTable.Columns.Count;
                    object[,] array2 = new object[num, count2];
                    try
                    {
                        int num2 = num;
                        while (j < count)
                        {
                            if (count - j < num)
                            {
                                num2 = count - j;
                            }
                            for (int k = 0; k < num2; k++)
                            {
                                for (int l = 0; l < count2; l++)
                                {
                                    array2[k, l] = myTable.Rows[k + j][l].ToString();
                                }
                                System.Windows.Forms.Application.DoEvents();
                            }
                            range = worksheet.get_Range("A" + (j + 2).ToString(), ((char)(65 + count2 - 1)).ToString() + (j + num2 + 1).ToString());
                            range.Value2 = array2;
                            j += num2;
                        }
                        worksheet.SaveAs(fileName, value, value, value, value, value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, value, value, value);
                        Marshal.ReleaseComObject(range);
                        this.KillSpecialExcel(application);
                        MessageBox.Show("导出数据成功！", "温馨提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                        return;
                    }
                    Thread.CurrentThread.CurrentCulture = currentCulture;
                    //11
                }
            }
        }
        //点击导出
        private void btnexport_Click(object sender, EventArgs e)
        {
            ToExcel(new SaveFileDialog());
            //122
        }

    }
}
