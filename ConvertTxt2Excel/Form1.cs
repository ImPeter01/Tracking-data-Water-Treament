using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Threading;
using System.Runtime.InteropServices;
using System.Data.OleDb;
using System.Globalization;
using System.Timers;
using System.Diagnostics;
using System.Configuration;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace ConvertTxt2Excel
{

    public partial class Form1 : Form
    {
        public enum SELECT_TOTAL
        {
            HOURS_RP,DAY_RP,MONTH_RP
        }
        public struct dataSelected
        {
            public string nameStation;
            public int index; 
        }
        public struct dataTracking
        {
            public string dataFlow1;
            public string dataTemp1;
            public string dataLevel1;
            public string dataTotal1;

            public string dataFlow2;
            public string dataTemp2;
            public string dataLevel2;
            public string dataTotal2;

            public string dataFlow3;
            public string dataTemp3;
            public string dataLevel3;
            public string dataTotal3;
        }
        public struct dataReportTotal
        {
            public double totalHours1;
            public double totalDays1;
            public double totalMonth1;

            public double totalHours2;
            public double totalDays2;
            public double totalMonth2;
        }
        static string[] is31DayInMont = { "01", "03", "05", "07", "08", "10", "12" };
        static dataSelected dataSelectedd = new dataSelected();
        static string getPath = "";
        static bool eventStop = false;
        static Queue<string> queue = new Queue<string>();
        dataTracking[] qDataTracking  = new dataTracking[5];
        List<dataReportTotal> dataReportTotals = new List<dataReportTotal>();
        static bool checkDistingueshExcel = false;
        static bool checkDistingueshSearch = false;
        static bool isClick = false;
        static bool oneTimes = false;
        static int saveHours;
        static int saveDays;
        static int saveMonth;
        public Form1()
        {
            InitializeComponent();
            init();
            string filePath = AppDomain.CurrentDomain.BaseDirectory + "\\source\\Logo.JPG";
            string[] s = { "\\bin" };
            string path = Application.StartupPath.Split(s, StringSplitOptions.None)[0] + "\\source\\Logo.JPG";
            //ShowMyImage(path, 77, 69);
        }
        public void killExcel(bool isAll = true)
        {
            if (isAll)
            {
                foreach (Process process in Process.GetProcessesByName("excel"))
                {
                    process.Kill();
                }
            }
            else
            {
                foreach (Process process in Process.GetProcessesByName("excel"))
                {
                    if(process.MainWindowTitle == "")
                    {
                        process.Kill();
                    }
             
                }
            }

            
        }
        public void init()
        {
            for (int i = 0; i < 5; i++)
            {
                qDataTracking[i].dataFlow1 = i.ToString();
                qDataTracking[i].dataTemp1 = i.ToString();
                qDataTracking[i].dataLevel1 = i.ToString();

                qDataTracking[i].dataFlow2 = i.ToString();
                qDataTracking[i].dataTemp2 = i.ToString();
                qDataTracking[i].dataLevel2 = i.ToString();

                qDataTracking[i].dataFlow3 = i.ToString();
                qDataTracking[i].dataTemp3 = i.ToString();
                qDataTracking[i].dataLevel3 = i.ToString();
            }
            killExcel();
            saveHours = System.DateTime.Now.Hour;
            saveDays = System.DateTime.Now.Day;
            saveMonth = System.DateTime.Now.Month;
            
            lbDateTime.Text = System.DateTime.Now.ToString("dddd , MMM dd yyyy,hh:mm:ss");
            cbSelectStation.SelectedIndex = 0;
            dataSelectedd.index = 0;
            btnSearch.Enabled = false;
            btnStart.Enabled = false;
            btnStop.Enabled = false;
            btnExcelPath.Enabled = true;
            btnTxtPath.Enabled = true;
            txtExcelPath.Enabled = true;
            txtPath.Enabled = true;
            lbStatus.Text = "STOP";
            lbStatus.ForeColor = System.Drawing.Color.Purple;
            dtFrom.Format = DateTimePickerFormat.Custom;
            dtFrom.CustomFormat = "yyyy/MM/dd HH:mm:ss";
            dtTo.Format = DateTimePickerFormat.Custom;
            dtTo.CustomFormat = "yyyy/MM/dd HH:mm:ss";
            //dtReport.Format = DateTimePickerFormat.Custom;
            // dtReport.CustomFormat = "yyyy/MM/dd ";   
            //FunctionCalculateToltal();
        }


        public DataTable ReadDataFromDateTime(string dateFrom, string dateTo, string path)
        {
            string connectString = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = '"+ path + "';" +
              "Extended Properties = 'Excel 12.0 Xml;HDR=YES'";
            DataTable dataTable = new DataTable();
            checkDistingueshSearch = true;
            if (!checkDistingueshExcel)
            {
                try
                { 
                    OleDbConnection oleDbConnection = new OleDbConnection(connectString);
                    //get table name
                    oleDbConnection.Open();
                    System.Data.DataTable dt = oleDbConnection.GetOleDbSchemaTable(
                        System.Data.OleDb.OleDbSchemaGuid.Tables, null);
                    oleDbConnection.Close();
                    ///get excel data
                    System.Data.OleDb.OleDbDataAdapter objAdapter = new System.Data.OleDb.OleDbDataAdapter
                        ("select * from [Sheet1$] Where [DateTime] >=" + dateFrom + "and" + "[DateTime]<=" + dateTo, oleDbConnection);

                    objAdapter.Fill(dataTable);
                    checkDistingueshSearch = false;
                }
                catch (Exception ex)
                {
                    checkDistingueshSearch = false;
                    MessageBox.Show(this,ex.Message,"Error!",MessageBoxButtons.OK,MessageBoxIcon.Error);
                }
            }
            else
            {
                checkDistingueshSearch = false;
                MessageBox.Show(this, "System are reading excel file! \n Please! Wating for few second!", "Information!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }         
            checkDistingueshSearch = false;
            return dataTable;
        }
        private void btnTxtPath_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog _txtFolder = new FolderBrowserDialog();
            if (_txtFolder.ShowDialog() == DialogResult.OK)
            {
                string _fileName = _txtFolder.SelectedPath;
                txtPath.Text = _fileName;
                if (txtExcelPath.Text != "")
                {
                    btnStart.Enabled = true;
                }
            }

        }

        private void btnExcelPath_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog _txtExcelOpenFileDialog = new FolderBrowserDialog();
            if (_txtExcelOpenFileDialog.ShowDialog() == DialogResult.OK)
            {
                string _fileName = _txtExcelOpenFileDialog.SelectedPath;
                txtExcelPath.Text = _fileName;
              
                if (txtPath.Text != "")
                {
                    btnStart.Enabled = true;
                    
                }
                btnSearch.Enabled = true;
            }
        }
        
        public string[] WriteSafeReadAllLines(String path)
        {
            using (var csv = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var sr = new StreamReader(csv))
            {
                List<string> file = new List<string>();
                while (!sr.EndOfStream)
                {
                    file.Add(sr.ReadLine());
                }

                return file.ToArray();
            }
        }
        public static string[] InternalReadAllLines(string path, Encoding encoding)
        {
            List<string> list = new List<string>();
            using (StreamReader streamReader = new StreamReader(path, encoding))
            {
                string str;
                while ((str = streamReader.ReadLine()) != null)
                    list.Add(str);
            }
            return list.ToArray();
        }
        public void TXT2EXCEL(string txtFile, string excelFile,int index)
        {
            checkDistingueshExcel = true;

            if (!checkDistingueshSearch)
            {
                try
                {
                    if(index <= 0)
                    {
                        index = 0;
                    }
                    else
                    {
                        index = index - 1;
                    }                
                    int i;
                    Excel.Application xlApp = null;
                    Excel.Workbook xlWorkBook = null;
                    Excel._Worksheet xlWorkSheet = null;
                    object misValue = System.Reflection.Missing.Value;
                    string[] lines = WriteSafeReadAllLines(txtFile);
                    xlApp = new Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Open(excelFile);
                    xlWorkSheet = xlWorkBook.Sheets[1];
                    char[] spearator = { ' ', '\t' };
                    Excel.Range range = xlWorkSheet.UsedRange;
                    int _lenRow = range.Rows.Count + 1;

                    var arrSlipt = lines[0].Split('\t');
                    var _lenLines = arrSlipt.Length;
                    if (_lenLines == 5)
                    {
                        for (i = 0; i < lines.Length; i++)
                        {
                            string[] _temp = lines[i].Split(spearator);
                            if (i == 0)
                            {
                                xlWorkSheet.Cells[range.Rows.Count + 1, 1] = _temp[3];
                                qDataTracking[index].dataFlow1 = _temp[1];
                            }
                            else if (i == 1)
                            {
                                qDataTracking[index].dataLevel1 = _temp[1];
                               
                            }
                            else if (i == 2)
                            {
                                qDataTracking[index].dataTemp1 = _temp[1];
                            }
                            qDataTracking[index].dataTemp1 = "---";
                            qDataTracking[index].dataFlow2 = "---";
                            qDataTracking[index].dataTemp2 = "---";
                            qDataTracking[index].dataLevel2 = "---";
                            qDataTracking[index].dataFlow3 = "---";
                            qDataTracking[index].dataTemp3 = "---";
                            qDataTracking[index].dataLevel3 = "---";
                            xlWorkSheet.Cells[range.Rows.Count + 1, i + 2] = _temp[1];

                        }
                    }
                    else
                    {
                        for (i = 0; i < lines.Length; i++)
                        {
                            string[] _temp = lines[i].Split(spearator);
                            if (i == 0)
                            {
                                xlWorkSheet.Cells[range.Rows.Count + 1, 1] = _temp[4];
                                qDataTracking[index].dataLevel1 = _temp[2];
                            }
                            else if (i == 1)
                            {
                                qDataTracking[index].dataTemp1 = _temp[2];
                                
                            }
                            else if (i == 2)
                            {
                                qDataTracking[index].dataFlow1 = _temp[2];
                            }
                            else if (i == 3)
                            {
                                qDataTracking[index].dataTotal1 = _temp[2];
                            }
                            else if (i == 4)
                            {
                                qDataTracking[index].dataLevel2 = _temp[2];
                                
                            }
                            else if (i == 5)
                            {
                                qDataTracking[index].dataTemp2 = _temp[2];
                            }
                            else if (i == 6)
                            {
                                qDataTracking[index].dataFlow2 = _temp[2];
                            }
                            else if (i == 7)
                            {
                                qDataTracking[index].dataTotal2 = _temp[2];
                            }
                            qDataTracking[index].dataFlow3 = "---";
                            qDataTracking[index].dataTemp3 = "---";
                            qDataTracking[index].dataLevel3 = "---";
                            qDataTracking[index].dataTotal3 = "---";
                            xlWorkSheet.Cells[range.Rows.Count + 1, i + 2] = _temp[2];
                        }
                    }
                    
                    if(index == dataSelectedd.index)
                    {
                        btnFlow1.Invoke(new Action(() =>
                        {
                            btnFlow1.Text = qDataTracking[index].dataFlow1;
                        }));
                        btnTemp1.Invoke(new Action(() =>
                        {
                            btnTemp1.Text = qDataTracking[index].dataTemp1;
                        }));
                        btnLevel1.Invoke(new Action(() =>
                        {
                            btnLevel1.Text = qDataTracking[index].dataLevel1;
                        }));

                        btnTotal1.Invoke(new Action(() =>
                        {
                            btnTotal1.Text = qDataTracking[index].dataTotal1;
                        }));
                        btnLevel2.Invoke(new Action(() =>
                        {
                            btnLevel2.Text = qDataTracking[index].dataLevel2;
                        }));
                        btnFlow2.Invoke(new Action(() =>
                        {
                            btnFlow2.Text = qDataTracking[index].dataFlow2;
                        }));

                        btnTemp2.Invoke(new Action(() =>
                        {
                            btnTemp2.Text = qDataTracking[index].dataTemp2;
                        }));
                        btnFlow3.Invoke(new Action(() =>
                        {
                            btnFlow3.Text = qDataTracking[index].dataFlow3;
                        }));
                        btnTotal2.Invoke(new Action(() =>
                        {
                            btnTotal2.Text = qDataTracking[index].dataTotal2;
                        }));



                        // Report Total hours
                        if(dataReportTotals[index].totalHours1 == 0.0)
                        {
                            btnFowHrs1.Invoke(new Action(() =>
                            {
                                btnFowHrs1.Text = "---";
                            }));
                        }
                        else
                        {
                            btnFowHrs1.Invoke(new Action(() =>
                            {
                                btnFowHrs1.Text = (double.Parse(qDataTracking[index].dataTotal1) - dataReportTotals[index].totalHours1).ToString();
                            }));
                        }

                        if (dataReportTotals[index].totalHours2 == 0.0)
                        {
                            btnFowHrs2.Invoke(new Action(() =>
                            {
                                btnFowHrs2.Text = "---";
                            }));
                        }
                        else
                        {
                            btnFowHrs2.Invoke(new Action(() =>
                            {
                                btnFowHrs2.Text = (double.Parse(qDataTracking[index].dataTotal2) - dataReportTotals[index].totalHours2).ToString();
                            }));
                        }

                        // Report Total days
                        if (dataReportTotals[index].totalDays1 == 0.0)
                        {
                            btnDay1.Invoke(new Action(() =>
                            {
                                btnDay1.Text = "---";
                            }));
                        }
                        else
                        {
                            btnDay1.Invoke(new Action(() =>
                            {
                                btnDay1.Text = (double.Parse(qDataTracking[index].dataTotal1) - dataReportTotals[index].totalDays1).ToString();
                            }));
                        }

                        if (dataReportTotals[index].totalDays2 == 0.0)
                        {
                            btnDay2.Invoke(new Action(() =>
                            {
                                btnDay2.Text = "---";
                            }));
                        }
                        else
                        {
                            btnDay2.Invoke(new Action(() =>
                            {
                                btnDay2.Text = (double.Parse(qDataTracking[index].dataTotal2) - dataReportTotals[index].totalDays2).ToString();
                            }));
                        }

                        // Report Total month
                        if (dataReportTotals[index].totalMonth1 == 0.0)
                        {
                            btnMonth1.Invoke(new Action(() =>
                            {
                                btnMonth1.Text = "---";
                            }));
                        }
                        else
                        {
                            btnMonth1.Invoke(new Action(() =>
                            {
                                btnMonth1.Text = (double.Parse(qDataTracking[index].dataTotal1) - dataReportTotals[index].totalMonth1).ToString();
                            }));
                        }

                        if (dataReportTotals[index].totalMonth2 == 0.0)
                        {
                            btnMonth2.Invoke(new Action(() =>
                            {
                                btnMonth2.Text = "---";
                            }));
                        }
                        else
                        {
                            btnMonth2.Invoke(new Action(() =>
                            {
                                btnMonth2.Text = (double.Parse(qDataTracking[index].dataTotal2) - dataReportTotals[index].totalMonth2).ToString();
                            }));
                        }

                        btnFowHrs3.Invoke(new Action(() =>
                        {
                            btnFowHrs3.Text = "---";
                        }));

                        btnMonth3.Invoke(new Action(() =>
                        {
                            btnMonth3.Text = "---";
                        }));

                        btnDay3.Invoke(new Action(() =>
                        {
                            btnDay3.Text = "---";
                        }));

                        dataReportTotal dataTemp = dataReportTotals[index];
                        //Update value in list Total
                        if (saveHours != System.DateTime.Now.Hour)
                        {
                            saveHours = System.DateTime.Now.Hour;
                            dataTemp.totalHours1 = double.Parse(qDataTracking[index].dataTotal1);
                            dataTemp.totalHours2 = double.Parse(qDataTracking[index].dataTotal2);
                        }
                        else if (saveDays != System.DateTime.Now.Day)
                        {
                            saveDays = System.DateTime.Now.Day;
                            dataTemp.totalDays1 = double.Parse(qDataTracking[index].dataTotal1);
                            dataTemp.totalDays2 = double.Parse(qDataTracking[index].dataTotal2);
                        }
                        else if( saveMonth != System.DateTime.Now.Month)
                        {
                            saveMonth = System.DateTime.Now.Month;
                            dataTemp.totalMonth1 = double.Parse(qDataTracking[index].dataTotal1);
                            dataTemp.totalMonth2 = double.Parse(qDataTracking[index].dataTotal2);
                        }
                        dataReportTotals[index] = dataTemp;
                    }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    if (xlWorkSheet != null) Marshal.ReleaseComObject(xlWorkSheet);
                    xlWorkBook.Save();
                    xlWorkBook.Close();
                    if (xlWorkBook != null) Marshal.ReleaseComObject(xlWorkBook);
                    xlApp.Quit();                                   
                    if (xlApp != null) Marshal.ReleaseComObject(xlApp);


                    checkDistingueshExcel = false;
                }
                catch (Exception ex)
                {
                    checkDistingueshExcel = false;
                    MessageBox.Show(ex.Message + ".");
                }
            }
            else
            {
                MessageBox.Show(this, "System are searching excel file! \n Please! Wating for few second!", "Information!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            
            checkDistingueshExcel = false;
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            try
            {                
                txtExcelPath.Enabled = false;
                txtPath.Enabled = false;
                btnStart.Enabled = false;
                btnExcelPath.Enabled = false;
                btnTxtPath.Enabled = false;
                lbStatus.Text = "Running";
                lbStatus.ForeColor = System.Drawing.Color.Lime;               
                if ((txtPath.Text != "") && (txtExcelPath.Text != ""))
                {
                    eventStop = false;
                    btnStop.Enabled = true;
                    if (!oneTimes)
                    {
                        EventChangeFolder(txtPath.Text);
                        oneTimes = true;
                    }                   
                }
                Thread threadReport = new Thread(FunctionCalculateToltal);
                threadReport.Start();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }
        
        public void EventChangeFolder(string path)
        {
            try
            {
                FileSystemWatcher watcher = new FileSystemWatcher(path);
                watcher.NotifyFilter = NotifyFilters.Attributes
                                     | NotifyFilters.CreationTime
                                     | NotifyFilters.DirectoryName
                                     | NotifyFilters.FileName
                                     | NotifyFilters.LastAccess
                                     | NotifyFilters.LastWrite
                                     | NotifyFilters.Security
                                     | NotifyFilters.Size;

                watcher.Created += OnCreate;

                watcher.Filter = "*.txt";
                watcher.IncludeSubdirectories = true;
                watcher.EnableRaisingEvents = true;
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }
        public void OnCreate(object sender, FileSystemEventArgs e)
        {
            getPath = e.FullPath;
            getPath = getPath.Replace("\\", "/");
            queue.Enqueue(getPath);
            if(queue.Count() == 1)
            {
                Thread thread = new Thread(runApp);
                thread.Start();
            }
        }
        string tempExcelPath = "";
        public void runApp()
        {
            string tempathsave = "";
            while (queue.Count != 0)
            {
                try
                {
                    if (eventStop)
                    {
                        continue;
                    }
                    else if (queue.Count > 0 && !isClick && !checkDistingueshSearch)
                    {
                        tempathsave = queue.Dequeue();
                        tempExcelPath = getNameStation(tempathsave);
                        var tempNumberStation = getNumberStation(tempExcelPath);
                        tempExcelPath = txtExcelPath.Text + "\\" + tempExcelPath + ".xlsx";
                        //killExcel();
                        Thread.Sleep(2000);
                        TXT2EXCEL(tempathsave, tempExcelPath, Int32.Parse(tempNumberStation));                     
                        
                        tempExcelPath = "";
                    }
                    else
                    {
                        continue;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);

                }
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
            // killExcel(false);
        }
        private void btnStop_Click(object sender, EventArgs e)
        {
            //getPath = "";
            //btnStop.Enabled = false;
            //btnStart.Enabled = true;
            //btnTxtPath.Enabled = true;
            //btnExcelPath.Enabled = true;
            //txtPath.Enabled = true;
            //txtExcelPath.Enabled = true;
            //eventStop = true;
            //lbStatus.Text = "STOP";
            //lbStatus.ForeColor = System.Drawing.Color.Purple;
            Environment.Exit(0);
            Application.Exit();
        }
        public string getNameStation(string path)
        {
            var items = path.Split('/');
            int _lenItems = items.Length;
            return items[_lenItems - 2];
        }
        public string getNumberStation(string path)
        {
            var temp = path.Split(' ');
            var rs = temp[0].Replace('.',' ');
            return rs;
        }
        public static void Release(object obj)
        {
            // Errors are ignored per Microsoft's suggestion for this type of function:
            // http://support.microsoft.com/default.aspx/kb/317109
            try
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(obj);
            }
            catch
            {
            }
        }
        public void copyExcel(string pathFileSource, string pathFileDestination)
        {
            try
            {
                Excel.Application excel = new Excel.Application();
                Excel.Workbook wbSource = excel.Workbooks.Open(pathFileSource, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                Excel.Workbook wbDestination = excel.Workbooks.Open(pathFileDestination, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                Excel.Worksheet WorksheetSource = wbSource.Sheets[1];
                //Copy all range in this worksheet
                object misValue = System.Reflection.Missing.Value;
                WorksheetSource.UsedRange.Copy(misValue);
                Excel.Worksheet WorksheetDestination = wbDestination.Sheets[1];

                // Select used Range, paste value only
                //WorksheetDestination.PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationAdd, false, false);
                WorksheetDestination.UsedRange.PasteSpecial();

                GC.Collect();
                GC.WaitForPendingFinalizers();
                Marshal.ReleaseComObject(WorksheetDestination);
                Marshal.ReleaseComObject(WorksheetSource);
                wbDestination.Save();
                wbSource.Close();
                Marshal.ReleaseComObject(wbSource);

                wbDestination.Close();
                Marshal.ReleaseComObject(wbDestination);
                //Quit application
                excel.Quit();
                Marshal.ReleaseComObject(excel);
                WorksheetDestination = null;
                WorksheetSource = null;
                wbDestination = null;
                wbSource = null;
                excel = null;
                //killExcel(false);
            }
            catch (Exception)
            {

                throw;
            }
           
        }
        private void btnSearch_Click(object sender, EventArgs e)
        {
            isClick = true;
            btnSearch.Enabled = false;
            Thread threadSearch = new Thread(runSearch);
            threadSearch.Start();
        }
        private void runSearch()
        {
            try
            {
                if (isClick)
                {
                    //killExcel(false);
                    string dateFrom = dtFrom.Text.Replace(" ", "").Replace(":", "").Replace("/", "");
                    string dateTo = dtTo.Text.Replace(" ", "").Replace(":", "").Replace("/", "");
                    string path = txtExcelPath.Text + "\\"+ getNameExcelNow() + ".xlsx";
                    string pathDes = txtExcelPath.Text +"\\" + "temp"+"\\"+ "temp1.xlsx";
                    if(tempExcelPath == path)
                    {
                        isClick = false;
                        btnSearch.Invoke(new Action(() => {
                            btnSearch.Enabled = true;
                        }));                  
                        //MessageBox.Show(this,"File Excel is running. \nPlease waiting for few second","Information");
                        return;
                    }
                    //copyExcel(path, pathDes);                  
                    dtgvData.Invoke(new Action(() => {
                        dtgvData.DataSource = null;
                        dtgvData.Rows.Clear();
                        dtgvData.Refresh();
                    }));                       
                    DataTable dataTable = ReadDataFromDateTime(dateFrom, dateTo, path);
                    isClick = false;
                    int counter = 0;
                    foreach (DataRow row in dataTable.Rows)
                    {
                        var flow1 = dataTable.Rows[counter]["Luu luong 01"].ToString();
                        var lvWater1 = dataTable.Rows[counter]["Muc nuoc 01"].ToString();
                        var temperature1 = dataTable.Rows[counter]["Nhiet do 01"].ToString();
                        var total1 = dataTable.Rows[counter]["Total 01"].ToString();
                        var flow2 = dataTable.Rows[counter]["Luu luong 02"].ToString();
                        var lvWater2 = dataTable.Rows[counter]["Muc nuoc 02"].ToString();
                        var temperature2 = dataTable.Rows[counter]["Nhiet do 02"].ToString();
                        var total2 = dataTable.Rows[counter]["Total 02"].ToString();
                        counter += 1;
                        var dtime = row["DateTime"].ToString().Substring(0, 4) + "-" +
                                        row["DateTime"].ToString().Substring(4, 2) + "-" +
                                        row["DateTime"].ToString().Substring(6, 2) + " " +
                                        row["DateTime"].ToString().Substring(8, 2) + "-" +
                                        row["DateTime"].ToString().Substring(10, 2);
                        dtgvData.Invoke(new Action(() => {                              
                            dtgvData.Rows.Add(counter,dtime,flow1, lvWater1,total1, temperature1, flow2, lvWater2, temperature2,total2);
                        }));                              
                    }
                    btnSearch.Invoke(new Action(() => {
                        btnSearch.Enabled = true;
                    }));                                              
                }
               
            }
            catch (Exception ex)
            {
                isClick = false;
                btnSearch.Invoke(new Action(() => {
                    btnSearch.Enabled = true;
                }));
                MessageBox.Show(ex.Message);
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Environment.Exit(0);
            Application.Exit();
        }

        private void btnViewchart_Click(object sender, EventArgs e)
        {
            ViewChart chartForm = new ViewChart();
            chartForm.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }
        public void updateDataTracking(dataTracking data)
        {
            btnFlow1.Text = data.dataFlow1;
            btnTemp1.Text = data.dataTemp1;
            btnLevel1.Text = data.dataLevel1;
            btnTotal1.Text = data.dataTotal1;

            btnFlow2.Text = data.dataFlow2;
            btnTemp2.Text = data.dataTemp2;
            btnLevel2.Text = data.dataLevel2;
            btnTotal2.Text = data.dataTotal2;

            btnFlow3.Text = data.dataFlow3;

        }
        public string getNameExcelNow()
        {
            string rs = "";
            string[] files = Directory.GetFiles(txtExcelPath.Text);
            string[] justFiless = files.Select(f => Path.GetFileNameWithoutExtension(f)).ToArray();
            for(int idx = 0;idx < justFiless.Length;idx++)
            {
                if(Int32.Parse(getNumberStation(justFiless[idx])) == (dataSelectedd.index +1))
                {
                    rs = justFiless[idx];
                    break;
                }    
            }
            return rs; 
        }
        public string[] getAllNameExcel()
        {
            return Directory.GetFiles(txtExcelPath.Text); ;
        }
        public void selectStation(dataSelected selected ,bool isdiff)
        {
            lbTitle.Text = selected.nameStation.ToUpper();
            var index = selected.index;
            updateDataTracking(qDataTracking[index]);
            dtgvData.DataSource = null;
            dtgvData.Rows.Clear();
            dtgvData.Refresh();
        }
        private void cbSelectStation_SelectedValueChanged(object sender, EventArgs e)
        {
            dataSelectedd.index = cbSelectStation.SelectedIndex;
            dataSelectedd.nameStation = cbSelectStation.SelectedItem.ToString();
            if(dataSelectedd.index == 0)
            {
                
                panel1_1.Show();
                panel1_2.Show();
                panel1_3.Hide();
                panel1_4.Hide();
                panel2_2.Hide();
                panel2_1.Hide();
                panel2_3.Hide();
                panel3_1.Hide();
                panel2_4.Hide();

                panel21_1.Show();
                panel21_2.Hide();
                panel21_3.Hide();
            }
            else if (dataSelectedd.index == 1)
            {

                panel1_1.Show();
                panel1_3.Show();
                panel1_2.Show();
                panel1_4.Show();
                panel2_2.Show();
                panel2_1.Show();
                panel2_3.Show();
                panel2_4.Show();
                panel3_1.Hide();

                panel21_1.Show();
                panel21_2.Show();
                panel21_3.Hide();
            }
            selectStation(dataSelectedd, false);
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            lbDateTime.Text = System.DateTime.Now.ToString("dddd , MMM dd yyyy,hh:mm:ss");
        }

        public void CreateFileExcelForDay()
        {
            var workbook = new XSSFWorkbook();
            var sheet = (XSSFSheet)workbook.CreateSheet("Daily Report");

            XSSFCellStyle headStyle = workbook.CreateCellStyle() as XSSFCellStyle;
            headStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            
            headStyle.FillPattern = FillPattern.SolidForeground;
            headStyle.FillForegroundColor = IndexedColors.Grey25Percent.Index;
            headStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            headStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            headStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            headStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;

            XSSFFont font = workbook.CreateFont() as XSSFFont;
            font.FontHeightInPoints = 12;
            headStyle.SetFont(font);

            XSSFCellStyle normalStyle = workbook.CreateCellStyle() as XSSFCellStyle;
            normalStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            //headStyle.SetFillBackgroundColor(new XSSFColor(new byte[] { 128, 128, 192 }));
            normalStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            normalStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            normalStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            normalStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            XSSFFont fontt = workbook.CreateFont() as XSSFFont;
            fontt.FontHeightInPoints = 12;
            normalStyle.SetFont(fontt);

            XSSFCellStyle titleStyle = workbook.CreateCellStyle() as XSSFCellStyle;
            titleStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            //headStyle.SetFillBackgroundColor(new XSSFColor(new byte[] { 128, 128, 192 }));
            XSSFFont fonttt = workbook.CreateFont() as XSSFFont;
            fonttt.FontHeightInPoints = 12;
            titleStyle.SetFont(fonttt);

            XSSFCellStyle timeStyle = workbook.CreateCellStyle() as XSSFCellStyle;
            timeStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            timeStyle.ShrinkToFit = true;
            //headStyle.SetFillBackgroundColor(new XSSFColor(new byte[] { 128, 128, 192 }));
            XSSFFont fontttt = workbook.CreateFont() as XSSFFont;
            fontttt.FontHeightInPoints = 12;
            fontttt.IsItalic = true;
            timeStyle.SetFont(fontttt);


            XSSFCellStyle timeRPStyle = workbook.CreateCellStyle() as XSSFCellStyle;
            timeRPStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            //headStyle.SetFillBackgroundColor(new XSSFColor(new byte[] { 128, 128, 192 }));
            XSSFFont fonttttt = workbook.CreateFont() as XSSFFont;
            fonttttt.FontHeightInPoints = 12;
            fonttttt.IsBold = true;
            timeRPStyle.SetFont(fonttttt);



            var row1 = sheet.CreateRow(0);
            row1.CreateCell(0).SetCellValue("Generator time create:");
            row1.GetCell(0).CellStyle = titleStyle;

            var row2 = sheet.CreateRow(1);
            //var createTimeRP = dtReport.Text;
            var time = System.DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            var createTime = "UTC + 07:00 " + " " + time;
            row2.CreateCell(0).SetCellValue(time);
            row2.GetCell(0).CellStyle = timeStyle;



            var row3 = sheet.CreateRow(3);
            var strRP = "Daily Report: " + time;
            row3.CreateCell(0).SetCellValue(strRP);
            row3.GetCell(0).CellStyle = timeRPStyle;

            var row4 = sheet.CreateRow(4);
            row4.CreateCell(0).SetCellValue("DateTime/Items");
            row4.GetCell(0).CellStyle = headStyle;
            for (int i = 1; i < 24; i++)
            {
                string temp = i.ToString() + ":00";
                row4.CreateCell(i).SetCellValue(temp);
                row4.GetCell(i).CellStyle = headStyle;
            }
            row4.CreateCell(24).SetCellValue("Total");
            row4.GetCell(24).CellStyle = headStyle;




            var row5 = sheet.CreateRow(5);
            row5.CreateCell(0).SetCellValue("Total flow (m3)");
            row5.GetCell(0).CellStyle = headStyle;
            for (int i = 1; i < 24; i++)
            {

                double db = 0.0;
                string temp = db.ToString();
                row5.CreateCell(i).SetCellValue(temp);
                row5.GetCell(i).CellStyle = normalStyle;
            }
            row5.CreateCell(24).SetCellValue("0.0");
            row5.GetCell(24).CellStyle = normalStyle;

            sheet.AutoSizeColumn(0);
      
            FileStream sw = File.Create(txtExcelPath.Text + "\\" + dataSelectedd.nameStation +"_" + "DailyReport.xlsx");
            workbook.Write(sw);
            workbook.Close();

        }
        public int CONSTAN_HOURS = 60;
        public void FunctionCalculateToltal()
        {
            var nameExcel = getAllNameExcel();
            
            dataReportTotal dataReportTotal;
            for (int i = 0; i < nameExcel.Length; i++)
            {
                double[] result = { 0.0,0.0 } ;
                //Hours SELECT MODE
                RunFunctionTotalReport(nameExcel[i], SELECT_TOTAL.HOURS_RP, ref result);
                dataReportTotal.totalHours1 = result[0];
                dataReportTotal.totalHours2 = result[1];

                //Days SELECT MODE
                double[] result1 = { 0.0,0.0 };
                RunFunctionTotalReport(nameExcel[i], SELECT_TOTAL.DAY_RP, ref result1);
                dataReportTotal.totalDays1 = result1[0];
                dataReportTotal.totalDays2 = result1[1];
    

                //Hours SELECT MODE
                double[] result2 = { 0.0, 0.0 };
                RunFunctionTotalReport(nameExcel[i], SELECT_TOTAL.MONTH_RP, ref result2);

                dataReportTotal.totalMonth1 = result2[0];
                dataReportTotal.totalMonth2 = result2[1];
                dataReportTotals.Add(dataReportTotal);
            }

        }

        public void RunFunctionTotalReport(string path, SELECT_TOTAL modeTotal, ref double[] result)
        {

            var dataTotalRP = SelectTotalHourBase(modeTotal, 0, path, false);
            var countRow = dataTotalRP.Rows.Count - 1;
            if(countRow < 0)
            {
                result[0] = 0.0;
                result[1] = 0.0;
                return;
            }
            var total1 = dataTotalRP.Rows[countRow]["Total 01"].ToString();
            var total2 = dataTotalRP.Rows[countRow]["Total 02"].ToString();
            result[0] = double.Parse(total1);
            result[1] = double.Parse(total2);

            //int counter_check = 0;
            //bool isTwoTimesSearch = false;
            //DataTable dataTotalRP = null;
            //while (counter_check < CONSTAN_HOURS)
            //{
            //    if (!isTwoTimesSearch)
            //    {
            //        dataTotalRP = SelectTotalHourBase(modeTotal, counter_check, path);
            //    }
            //    else
            //    {
            //        dataTotalRP = SelectTotalHourBase(modeTotal, counter_check, path, false);
            //    }

            //    if (dataTotalRP != null)
            //    {

            //        var total1 = dataTotalRP.Rows[0]["Total 01"].ToString();
            //        var total2 = dataTotalRP.Rows[0]["Total 02"].ToString();
            //        result[0] = double.Parse(total1);
            //        result[1] = double.Parse(total2);
            //        break;
            //    }
            //    if (counter_check == CONSTAN_HOURS - 1)
            //    {
            //        if (!isTwoTimesSearch)
            //        {
            //            isTwoTimesSearch = true;
            //            counter_check = 0;
            //            continue;
            //        }
            //        result[0] = 0.0;
            //        result[1] = 0.0;
            //    }
            //    counter_check += 1;
            //}

        }
        public DataTable SelectTotalHourBase(SELECT_TOTAL modeSelect ,int tol, string path, bool positive = true)
        {
            var years = System.DateTime.Now.Year;
            string year = "";
            string month = "";
            string day = "";
            string hour = "";
            string min = "";
            string second = "";
            string minn = "";
            if (modeSelect == SELECT_TOTAL.HOURS_RP)
            {
                var months = System.DateTime.Now.Month;
                var days = System.DateTime.Now.Day;
                var hours = System.DateTime.Now.Hour;
                int mins;            

                if(positive)
                {
                    mins =  tol;
                }
                else
                {
                    mins = 59 - tol;
                    hours = hours - 1;
                    if (hours < 0)
                    {
                        days = days - 1;
                        if (days <= 0)
                        {
                            months = months - 1;
                            if(months <=0)
                            {
                                years = years - 1;
                            }
                            if (is31DayInMont.Contains(months.ToString()))
                            {
                                days = 31;
                            }
                            else
                            {
                                days = 30;
                            }
                        }
                        hours = 23;
                    }
                }
                year = years.ToString();
                month = months.ToString();
                day = days.ToString();
                hour = hours.ToString();
                min = mins.ToString();
                minn = (mins - 59).ToString();
                if (month.Length == 1)
                {
                    month = "0" + month;
                }
                if (day.Length == 1)
                {
                    day = "0" + day;
                }
                if (min.Length == 1)
                {
                    min = "0" + min;
                }
                if (minn.Length == 1)
                {
                    minn = "0" + minn;
                }
                if (hour.Length == 1)
                {
                    hour = "0" + hour;
                }
                second = "00";
               
            }
            else if(modeSelect == SELECT_TOTAL.DAY_RP)
            {              
                var months = System.DateTime.Now.Month;
                var days = System.DateTime.Now.Day;
                var hours = 0;
                var mins = 0;
                if(positive)
                {
                    hours = 0;
                    mins = tol;
                }
                else
                {                    
                    days = days - 1;
                    if(days <=0)
                    {
                        months = months - 1;
                        if (months <= 0)
                        {
                            years = years - 1;
                        }
                        if (is31DayInMont.Contains(months.ToString()))
                        {
                            days = 31;
                        }
                        else
                        {
                            days = 30;
                        }
                    }
                    mins = 59 - tol;
                    hours = 23;
                }
                year = years.ToString();
                month = months.ToString();
                day = days.ToString();
                hour = hours.ToString();
                min = mins.ToString();
                minn = (mins - 59).ToString();
                if (month.Length == 1)
                {
                    month = "0" + month;
                }
                if (day.Length == 1)
                {
                    day = "0" + day;
                }
                if (min.Length == 1)
                {
                    min = "0" + min;
                }
                if (minn.Length == 1)
                {
                    minn = "0" + minn;
                }

                if (hour.Length == 1)
                {
                    hour = "0" + hour;
                }

                second = "00";
            }
            else if(modeSelect == SELECT_TOTAL.MONTH_RP)
            {
                var months = System.DateTime.Now.Month;
                int mins = 0;
                int days = 1;
                int hours = 0;
                if(positive)
                {
                    mins = tol;
                }
                else
                {
                    days = days - 1;
                    if (days <= 0)
                    {
                        months = months - 1;
                        if (months <= 0)
                        {
                            years = years - 1;
                        }
                        string temp = months.ToString(); ;
                        if (months.ToString().Length == 1)
                            temp = "0" + months.ToString();
                        if (is31DayInMont.Contains(temp))
                        {
                            days = 31;
                        }
                        else
                        {
                            days = 30;
                        }
                    }
                    mins = 59 - tol;
                    hours = 23;
                }

                year = years.ToString();
                month = months.ToString();
                day = days.ToString();
                hour = hours.ToString();
                min = mins.ToString();
                minn = (mins - 59).ToString();
                if (month.Length == 1)
                {
                    month = "0" + month;
                }
                if (day.Length == 1)
                {
                    day = "0" + day;
                }
                if (min.Length == 1)
                {
                    min = "0" + min;
                }
                if (minn.Length == 1)
                {
                    minn = "0" + minn;
                }
                if (hour.Length == 1)
                {
                    hour = "0" + hour;
                }
            }
            var keymin = year + month  + day + hour + minn+ second;
            var keymax = year + month  + day + hour + min + second;
            return FunctionSearch(keymin,keymax, path);
        }      
        public DataTable FunctionSearch(string keyMin, string keyMax, string path)
        {
            string connectString = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = '" + path + "';" +
              "Extended Properties = 'Excel 12.0 Xml;HDR=YES'";
            DataTable dataTable = new DataTable();
            try
            {
                OleDbConnection oleDbConnection = new OleDbConnection(connectString);
                //get table name
                oleDbConnection.Open();
                System.Data.DataTable dt = oleDbConnection.GetOleDbSchemaTable(
                    System.Data.OleDb.OleDbSchemaGuid.Tables, null);
                oleDbConnection.Close();
                ///get excel data
                System.Data.OleDb.OleDbDataAdapter objAdapter = new System.Data.OleDb.OleDbDataAdapter
                    ("select * from[Sheet1$] Where[DateTime] >= " + keyMin + "and" + "[DateTime] <= " + keyMax, oleDbConnection);
                objAdapter.Fill(dataTable);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return dataTable;
        }
    }
}
