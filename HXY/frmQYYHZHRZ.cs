using China_System.Common;
using QM.Buiness;
using QM.DB;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace HXY
{
    public partial class frmQYYHZHRZ : Form
    {
        string inputFileName;

        // 后台执行控件
        private BackgroundWorker bgWorker;
        // 消息显示窗体
        private frmMessageShow frmMessageShow;
        // 后台操作是否正常完成
        private bool blnBackGroundWorkIsOK = false;
        //后加的后台属性显
        private bool backGroundRunResult;

        List<clsDatabaseinfo> RResult;
        string savepath;
        public frmQYYHZHRZ()
        {
            InitializeComponent();
        }
        private void InitialBackGroundWorker()
        {
            bgWorker = new BackgroundWorker();
            bgWorker.WorkerReportsProgress = true;
            bgWorker.WorkerSupportsCancellation = true;
            bgWorker.RunWorkerCompleted +=
                new RunWorkerCompletedEventHandler(bgWorker_RunWorkerCompleted);
            bgWorker.ProgressChanged +=
                new ProgressChangedEventHandler(bgWorker_ProgressChanged);
        }

        private void bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                blnBackGroundWorkIsOK = false;
            }
            else if (e.Cancelled)
            {
                blnBackGroundWorkIsOK = true;
            }
            else
            {
                blnBackGroundWorkIsOK = true;
            }
        }

        private void bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (frmMessageShow != null && frmMessageShow.Visible == true)
            {
                //设置显示的消息
                frmMessageShow.setMessage(e.UserState.ToString());
                //设置显示的按钮文字
                if (e.ProgressPercentage == clsConstant.Thread_Progress_OK)
                {
                    frmMessageShow.setStatus(clsConstant.Dialog_Status_Enable);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //
            try
            {

                InitialBackGroundWorker();
                bgWorker.DoWork += new DoWorkEventHandler(Read_arir);

                bgWorker.RunWorkerAsync();
                Control.CheckForIllegalCrossThreadCalls = false;
                // 启动消息显示画面
                frmMessageShow = new frmMessageShow(clsShowMessage.MSG_001,
                                                    clsShowMessage.MSG_007,
                                                    clsConstant.Dialog_Status_Disable);
                frmMessageShow.ShowDialog();
          
                // 数据读取成功后在画面显示
                if (blnBackGroundWorkIsOK)
                {
                    chushihua();

                }
            }

            catch (Exception ex)
            {
                MessageBox.Show("错误" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);

                return;
                throw ex;
            }
        }

        private void chushihua()
        {
            this.dataGridView1.DataSource = null;
            this.dataGridView1.AutoGenerateColumns = false;
            this.dataGridView1.DataSource = RResult;
            label2.Text = "共计 : " + RResult.Count;
        }
        private void Read_arir(object sender, DoWorkEventArgs e)
        {
            DateTime oldDate = DateTime.Now;

            //初始化信息
            clsAllnew BusinessHelp = new clsAllnew();
            RResult = new List<clsDatabaseinfo>();
            BusinessHelp.bgWorker1 = bgWorker;
            RResult = BusinessHelp.ReadfindngFile(inputFileName);

            #region 读取网站信息
            //BusinessHelp.RResult = RResult;

            //BusinessHelp.run_type = 2;
            //BusinessHelp.ReadGeckoWEN(null);
            //RResult = BusinessHelp.RResult; 
            #endregion

            DateTime FinishTime = DateTime.Now;
            TimeSpan s = DateTime.Now - oldDate;
            string timei = s.Minutes.ToString() + ":" + s.Seconds.ToString();
            string Showtime = clsShowMessage.MSG_029 + timei.ToString();
            bgWorker.ReportProgress(clsConstant.Thread_Progress_OK, clsShowMessage.MSG_009 + "\r\n" + Showtime);

        }
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*.xlsx";
            ofd.FileName = "";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                inputFileName = ofd.FileName;
                this.textBox1.Text = inputFileName;

            }
            else
            {
                return;
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
              DataGridViewColumn column = dataGridView1.Columns[e.ColumnIndex];

              if (column == tijiaoanjian)
              {
                  var oids = GetOrderIdsBySelectedGridCell();
                  clsAllnew BusinessHelp = new clsAllnew();

                  var filtered = RResult.Find(s => s.zhanghao == oids[0]);
                  if (filtered != null)
                      BusinessHelp.Item_RResult = filtered;

                  moveFolder(filtered.kaihusheng, filtered.kaihushi, filtered.kaihuzhihang, filtered.kaihuhang);

                  BusinessHelp.run_type = 3;
                  BusinessHelp.ReadGeckoWEN(null);
              }
        }
        private void moveFolder(string qunmingcheng, string body, string kaihuzhihang,string kaihuhang)
        {
            wirite_txt(qunmingcheng, body, kaihuzhihang, kaihuhang);
            string path = AppDomain.CurrentDomain.BaseDirectory + "System";
            string dir = @"C:\Program Files (x86)\HXY\System";
            CopyFolder(path, dir);


        }
        public static void CopyFolder(string sourcePath, string destPath)
        {
            if (Directory.Exists(sourcePath))
            {
                if (!Directory.Exists(destPath))
                {
                    //目标目录不存在则创建
                    try
                    {
                        Directory.CreateDirectory(destPath);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("创建目标目录失败：" + ex.Message);
                    }
                }
                //获得源文件下所有文件
                List<string> files = new List<string>(Directory.GetFiles(sourcePath));
                files.ForEach(c =>
                {
                    string destFile = Path.Combine(new string[] { destPath, Path.GetFileName(c) });
                    File.Copy(c, destFile, true);//覆盖模式
                });
                //获得源文件下所有目录文件
                List<string> folders = new List<string>(Directory.GetDirectories(sourcePath));
                folders.ForEach(c =>
                {
                    string destDir = Path.Combine(new string[] { destPath, Path.GetFileName(c) });
                    //采用递归的方法实现
                    CopyFolder(c, destDir);
                });
            }
            else
            {
                throw new DirectoryNotFoundException("源目录不存在！");
            }
        }

        private void wirite_txt(string qunmingcheng, string body, string kaihuzhihang, string kaihuhang)
        {
            string A_Path = AppDomain.CurrentDomain.BaseDirectory + "System\\kaihusheng.txt";

            StreamWriter sw = new StreamWriter(A_Path);

            sw.WriteLine(qunmingcheng.Trim());

            sw.Flush();
            sw.Close();

            A_Path = AppDomain.CurrentDomain.BaseDirectory + "System\\kaihushi.txt";

            sw = new StreamWriter(A_Path);

            sw.WriteLine(body.Trim());


            sw.Flush();
            sw.Close();



            A_Path = AppDomain.CurrentDomain.BaseDirectory + "System\\kaihuzhihang.txt";

            sw = new StreamWriter(A_Path);

            sw.WriteLine(kaihuzhihang.Trim());


            sw.Flush();
            sw.Close();

            A_Path = AppDomain.CurrentDomain.BaseDirectory + "System\\kaihuhang.txt";

            sw = new StreamWriter(A_Path);

            sw.WriteLine(kaihuhang.Trim());


            sw.Flush();
            sw.Close();

        }

        private List<string> GetOrderIdsBySelectedGridCell()
        {

            List<string> order_ids = new List<string>();
            var rows = GetSelectedRowsBySelectedCells(dataGridView1);
            foreach (DataGridViewRow row in rows)
            {
                var Diningorder = row.DataBoundItem as clsDatabaseinfo;
                order_ids.Add((string)Diningorder.zhanghao);
            }

            return order_ids;
        }
        private IEnumerable<DataGridViewRow> GetSelectedRowsBySelectedCells(DataGridView dgv)
        {
            List<DataGridViewRow> rows = new List<DataGridViewRow>();
            foreach (DataGridViewCell cell in dgv.SelectedCells)
            {
                rows.Add(cell.OwningRow);

            }

            return rows.Distinct();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            if (this.dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("Sorry , No Data Output !", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.DefaultExt = ".csv";
            saveFileDialog.Filter = "csv|*.csv";
            string strFileName = "下载 信息" + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
            saveFileDialog.FileName = strFileName;
            if (saveFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                strFileName = saveFileDialog.FileName.ToString();
            }
            else
            {
                return;
            }
            FileStream fa = new FileStream(strFileName, FileMode.Create);
            StreamWriter sw = new StreamWriter(fa, Encoding.Unicode);
            string delimiter = "\t";
            string strHeader = "";
            for (int i = 0; i < this.dataGridView1.Columns.Count; i++)
            {
                strHeader += this.dataGridView1.Columns[i].HeaderText + delimiter;
            }
            sw.WriteLine(strHeader);

            //output rows data
            for (int j = 0; j < this.dataGridView1.Rows.Count; j++)
            {
                string strRowValue = "";

                for (int k = 0; k < this.dataGridView1.Columns.Count; k++)
                {
                    if (this.dataGridView1.Rows[j].Cells[k].Value != null)
                    {
                        strRowValue += this.dataGridView1.Rows[j].Cells[k].Value.ToString().Replace("\r\n", " ").Replace("\n", "") + delimiter;
                        if (this.dataGridView1.Rows[j].Cells[k].Value.ToString() == "LIP201507-35")
                        {

                        }

                    }
                    else
                    {
                        strRowValue += this.dataGridView1.Rows[j].Cells[k].Value + delimiter;
                    }
                }
                sw.WriteLine(strRowValue);
            }
            sw.Close();
            fa.Close();
            MessageBox.Show("下载完成 ！", "System", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }

        private void button4_Click(object sender, EventArgs e)
        {
            clsAllnew BusinessHelp = new clsAllnew();
            BusinessHelp.tsStatusLabel1 = toolStripLabel1;
            #region 读取网站信息
            BusinessHelp.RResult = RResult;

            BusinessHelp.run_type = 2;
            BusinessHelp.ReadGeckoWEN(null);
            RResult = BusinessHelp.RResult;

            chushihua();
            #endregion
        }
    }
}
