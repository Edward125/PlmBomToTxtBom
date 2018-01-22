using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Edward;
using System.Runtime.InteropServices;

namespace PlmBomToTxtBom
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        // help
        //http://www.xiaoten.com/operation-excel-by-c-sharp-language.html
        //https://www.cnblogs.com/WarBlog/articles/5646906.html

        public static string appFolder = System.Windows.Forms. Application.StartupPath + @"\PLMBOM";

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Text = "PLM Exeel BOM To Text BOM ,Ver:" + System.Windows.Forms.Application.ProductVersion + ",Author:Edward Song"; 
            //
            createAppFolder();

            txtExcelFile.SetWatermark("Double click here to select the excel file download from PLM");
        }




        #region DataSet


        static bool  DataSetParse(string fileName , out DataSet ds)
        {
            // string connectionString = string.Format("provider=Microsoft.Jet.OLEDB.4.0; data source={0};Extended Properties=Excel 8.0;", fileName);


            ////2003（Microsoft.Jet.Oledb.4.0）
            //string strConn = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'", excelFilePath);
            ////2010（Microsoft.ACE.OLEDB.12.0）
            //string strConn = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'", excelFilePath);

            string connectionString = string.Empty;
            System.IO.FileInfo fi = new System.IO.FileInfo(fileName);
            //MessageBox.Show(fi.Extension);
            DataSet data = new DataSet();
            try
            {
                if (fi.Extension == ".xls")
                    connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'", fileName);
                if (fi.Extension == ".xlsx")
                    connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'", fileName);
            }
            catch (Exception  ex)
            {

                MessageBox.Show(ex.Message);
                ds = data;
                return false;
            }

            foreach (var sheetName in GetExcelSheetNames(connectionString))
            {
                using (OleDbConnection con = new OleDbConnection(connectionString))
                {
                    Console.WriteLine(sheetName);
                    var dataTable = new System.Data.DataTable(sheetName);
                    string query = string.Format("SELECT * FROM [{0}]", sheetName);
                    con.Open();
                    OleDbDataAdapter adapter = new OleDbDataAdapter(query, con);
                    adapter.Fill(dataTable);
                    data.Tables.Add(dataTable);

                }
            }

            ds = data;

            return true;

        }

        static string[] GetExcelSheetNames(string connectionString)
        {
            OleDbConnection con = null;
            System.Data.DataTable dt = null;
            con = new OleDbConnection(connectionString);
            con.Open();
            dt = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            if (dt == null)
            {
                return null;
            }

            String[] excelSheetNames = new String[dt.Rows.Count];
            int i = 0;

            foreach (DataRow row in dt.Rows)
            {
                excelSheetNames[i] = row["TABLE_NAME"].ToString();
                i++;
            }

            return excelSheetNames;
        }

        #endregion

        #region checkFolder


        private void createAppFolder()
        {

           //if (Directory.Exists(appFolder))
            // Directory.Delete(appFolder, true);
            if (!Directory.Exists (appFolder ))
               Directory.CreateDirectory(appFolder);

        }

        #endregion

        private void txtExcelFile_DoubleClick(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "(Excel files)|*.xls;*.xlsx";
            open.Multiselect = false;
            if (open.ShowDialog() == DialogResult.OK)
            {
                FileInfo fi = new FileInfo(open.FileName);
                if ((fi.Extension == ".xls") || (fi.Extension == ".xlsx"))
                {
                    txtExcelFile.Text = open.FileName;
                }
                else
                {
                    MessageBox.Show("you select file is not excel file...", "File Not Excel", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }
            }
        }

        private void btnGo_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty (txtExcelFile.Text.Trim ()))
            {
                return;
            }

            if (fileIsExcel(txtExcelFile.Text.Trim()))
            {
                btnGo.Enabled = false;   


              // DateTime  startTime=DateTime. Now.AddSeconds(-1);
                Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                Workbooks wkbs = app.Workbooks;
                _Workbook _wbk = wkbs.Add(txtExcelFile.Text.Trim()); //open the excel file
               // DateTime endTime = DateTime.Now.AddSeconds(1);

                //
                Sheets shs = _wbk.Sheets;
                _Worksheet _wsh = (_Worksheet)shs.get_Item(1);
                // _Worksheet _wsh = (_Worksheet)_wbk.Sheets[0];
                ((Range)_wsh.Rows[1, Type.Missing]).Delete(XlDeleteShiftDirection.xlShiftUp);

                // _wsh.get_Range(_wsh.Cells[1, 2], _wsh.Cells[_wsh.Rows.Count, 2]).Delete(XlDeleteShiftDirection.xlShiftToLeft);
                //  _wsh.get_Range(_wsh.Cells[1, 2],_wsh.Cells[_wsh.Rows.Count, 2]).Delete(XlDeleteShiftDirection.xlShiftToLeft);
                int _usedColumns = _wsh.UsedRange.Columns.Count;
                int _usedRows = _wsh.UsedRange.Rows.Count;

                int _start = -1;
                int _Level = 1; //value = 1,default
                int _PartNumber = -1;
                int _PartDescription = -1;
                int _Qty = -1;
                int _Location = -1;



                string TXTBOM = appFolder + @"\" + txtExcelFile.Text.Trim().Substring(txtExcelFile.Text.Trim().LastIndexOf(@"\") + 1, txtExcelFile.Text.Trim().Length - txtExcelFile.Text.Trim().LastIndexOf(@"\") - 1) + @".txt";

                if (File.Exists(TXTBOM))
                {
                    try
                    {
                        File.Delete(TXTBOM);
                    }
                    catch (Exception ex)
                    {

                        btnGo.Enabled = true;
                    }
                }


              

                for (int i = 1; i < _usedRows ; i++)
                {
                    if (_wsh.Range["A" + i].Value.ToString().Trim().ToUpper() == "LEVEL".ToUpper ())
                    {
                        _start = i;
                        
                        for (int j = 1; j < _usedColumns; j++)
                        {
                          
                            if (_wsh.Range[Other.Chr(64 + j) + i].Value.ToString().Trim().ToUpper() == "Part Number".ToUpper())
                                _PartNumber = j;
                            if (_wsh.Range[Other.Chr(64 + j) + i].Value.ToString().Trim().ToUpper() == "Part Description".ToUpper())
                               _PartDescription  = j;
                            if (_wsh.Range[Other.Chr(64 + j) + i].Value.ToString().Trim().ToUpper() == "Qty".ToUpper())
                                _Qty = j;
                            if (_wsh.Range[Other.Chr(64 + j) + i].Value.ToString().Trim().ToUpper() == "Location".ToUpper())
                                _Location = j;
                        }
                        break;
                    }             
                }


                //
                string _sLevel = "";
                string _sPartNumber = "";
                string _sPartDescription = "";
                string _sQty = "";
                string _sLocation = "";
                string _sLine = "";           

                
                for (int i = _start  + 1; i < _usedRows; i++)
                {
                    try
                    {
                        _sLevel = _wsh.Range[Other.Chr(64 + _Level) + i].Value.ToString();
                        _sPartNumber = _wsh.Range[Other.Chr(64 + _PartNumber) + i].Value.ToString();
                        _sPartDescription = _wsh.Range[Other.Chr(64 + _PartDescription) + i].Value.ToString();
                        _sQty = _wsh.Range[Other.Chr(64 + _Qty) + i].Value.ToString();
                        _sLocation = _wsh.Range[Other.Chr(64 + _Location) + i].Value.ToString();

                        StreamWriter sw = new StreamWriter(TXTBOM, true, Encoding.UTF8);

                        if (!string.IsNullOrEmpty (_sLevel )  && !string.IsNullOrEmpty(_sLocation))
                        {

                            if (_sLocation.Contains(","))
                            {
                                foreach (string item in _sLocation.Split(','))
                                {
                                    _sLine = _sPartNumber.PadRight(36) + _sPartDescription.PadRight(65) + _sQty.PadRight(10) + item;
                                    sw.WriteLine(_sLine);
                                }
                            }
                            else
                            {
                                _sLine = _sPartNumber.PadRight(36) + _sPartDescription.PadRight(65) + _sQty.PadRight(10) + _sLocation;
                                sw.WriteLine(_sLine);
                            }

                            //FileStream fs = File.Create(TXTBOM);
                            //fs.Close(); 
                           
                          
                          
                            
                        }

                        sw.Close();


                    }
                    catch (Exception)
                    {
                        
                       // throw;
                        btnGo.Enabled = true;
                    }

                }

                
            
                //屏蔽掉系统跳出的Alert
                app.AlertBeforeOverwriting = false;
                app.DisplayAlerts = false;
                app.Visible = false;
                //保存到指定目录
                // _wbk.SaveAs(appFolder + @"\.1.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                //     Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                //_wbk.SaveAs(appFolder + @"\1.xlsx");


                app.Quit();
                //释放掉多余的excel进程
               //   System.Runtime.InteropServices.Marshal.ReleaseComObject(app);            
                
                //app = null;
                KillProcess(app);

                //foreach (System.Diagnostics.Process theProc in System.Diagnostics.Process.GetProcessesByName("EXCEL"))
                //{
                //    if (theProc.StartTime.CompareTo(startTime) > 0 && theProc.StartTime.CompareTo(endTime) < 0)
                //    {
                //        theProc.Kill();
                //        break;
                //    }
                //}


                MessageBox.Show("Complete!,File save in'" + TXTBOM + "'", "Save OK",MessageBoxButtons.OK, MessageBoxIcon.Information);
                btnGo.Enabled = true;
            }
            else
            {
                txtExcelFile.SelectAll();
                txtExcelFile.Focus();
            }
        }



        /// <summary>
        /// 导出Excel后，杀死Excel进程
        /// </summary>
        /// <param name="app"></param>
        private static void KillProcess(Microsoft.Office.Interop.Excel.Application   app)
        {
            IntPtr t = new IntPtr(app.Hwnd);
            int k = 0;
            GetWindowThreadProcessId(t, out k);
            System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);
            p.Kill();
        }

        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out   int ID);



        private void useDataSet()
        {

            //=================

            DataSet ds = new DataSet();
            DataSetParse(txtExcelFile.Text.Trim(), out ds);
            // dataGridView1.DataSource = ds.Tables[0];

            //column index

            int _Level = -1; //value = 1,default
            int _PartNumber = -1;
            int _PartDescription = -1;
            int _Qty = -1;
            int _Location = -1;
            string TXTBOM = appFolder + @"\" + txtExcelFile.Text.Trim().Substring(txtExcelFile.Text.Trim().LastIndexOf(@"\") + 1, txtExcelFile.Text.Trim().Length - txtExcelFile.Text.Trim().LastIndexOf(@"\") - 1) + @".txt";

            //get ColumnName 
            for (int i = 0; i < ds.Tables[0].Columns.Count; i++)
            {
                // MessageBox.Show(ds.Tables[0].Columns[i].ToString());
                if (ds.Tables[0].Columns[i].ToString().ToUpper().Trim() == "LEVEL")
                    _Level = i;
                if (ds.Tables[0].Columns[i].ToString().ToUpper().Trim() == "Part Number".ToUpper())
                    _PartNumber = i;
                if (ds.Tables[0].Columns[i].ToString().ToUpper().Trim() == "Part Description".ToUpper())
                    _PartDescription = i;
                if (ds.Tables[0].Columns[i].ToString().ToUpper().Trim() == "Qty".ToUpper())
                    _Qty = i;
                if (ds.Tables[0].Columns[i].ToString().ToUpper().Trim() == "Location".ToUpper())
                    _Location = i;
            }


            //
            string _sLevel = "";
            string _sPartNumber = "";
            string _sPartDescription = "";
            string _sQty = "";
            string _sLocation = "";
            string _sLine = "";





            if (File.Exists(TXTBOM))
            {
                try
                {
                    File.Delete(TXTBOM);
                }
                catch (Exception ex)
                {


                }
            }


            for (int i = 0; i < ds.Tables[0].Columns.Count; i++)
            {
                for (int j = 0; j < ds.Tables[0].Rows.Count; j++)
                {
                    _sLevel = ds.Tables[0].Rows[j][_Level].ToString();
                    _sLocation = ds.Tables[0].Rows[j][_Location].ToString();
                    if (_sLevel == "1" && !string.IsNullOrEmpty(_sLocation))
                    {
                        _sPartNumber = ds.Tables[0].Rows[j][_PartNumber].ToString();
                        _sPartDescription = ds.Tables[0].Rows[j][_PartDescription].ToString();
                        _sQty = ds.Tables[0].Rows[j][_Qty].ToString();

                        if (_sLocation.Contains(","))
                        {
                            foreach (string item in _sLocation.Split(','))
                            {
                                _sLine = _sPartNumber.PadRight(36) + _sPartDescription.PadRight(65) + _sQty.PadRight(10) + item;
                            }
                        }
                        else
                        {
                            _sLine = _sPartNumber.PadRight(36) + _sPartDescription.PadRight(65) + _sQty.PadRight(10) + _sLocation;
                        }
                        StreamWriter sw = new StreamWriter(TXTBOM, true, Encoding.UTF8);
                        sw.WriteLine(_sLine);
                        sw.Close();
                    }
                }

            }

        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="filepath"></param>
        /// <returns></returns>
        private bool fileIsExcel(string filepath)
        {
            FileInfo fi = new FileInfo(filepath);
            if ((fi.Extension == ".xls") || (fi.Extension == ".xlsx"))
            {
                return true;
            }
            else
            {
                MessageBox.Show("you select file is not excel file...", "File Not Excel", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
        }

        private void btnHelp_Click(object sender, EventArgs e)
        {
            if (this.Height == 422)
            {
                this.Height = 122;
                lblInfo.Visible = false;
            }
            else
            {
                this.Height = 422;
                lblInfo.Visible = true;
            }
                  
        }

        private void txtExcelFile_TextChanged(object sender, EventArgs e)
        {

        }


    }
}
