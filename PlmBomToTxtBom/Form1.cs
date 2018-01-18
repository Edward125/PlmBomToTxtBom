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


              //  Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
              //  Workbooks wkbs = app.Workbooks;
              //  _Workbook _wbk = wkbs.Add(txtExcelFile.Text.Trim()); //open the excel file

              //  //
              //  Sheets shs = _wbk.Sheets;
              //  _Worksheet _wsh = (_Worksheet)shs.get_Item(1);
              // // _Worksheet _wsh = (_Worksheet)_wbk.Sheets[0];

              //  ((Range)_wsh.Rows[1,Type.Missing ]).Delete(XlDeleteShiftDirection.xlShiftUp);
              // // _wsh.get_Range(_wsh.Cells[1, 2], _wsh.Cells[_wsh.Rows.Count, 2]).Delete(XlDeleteShiftDirection.xlShiftToLeft);
              ////  _wsh.get_Range(_wsh.Cells[1, 2],_wsh.Cells[_wsh.Rows.Count, 2]).Delete(XlDeleteShiftDirection.xlShiftToLeft);
              //  //string ss = _wsh.Range["A1"].Value.ToString();
              // // MessageBox.Show(ss);
              //  //MessageBox.Show(_wsh.UsedRange.Columns.Count.ToString());
              //  //MessageBox.Show(_wsh.UsedRange.Rows.Count.ToString());

              //  string TXTBOM = appFolder + @"\" + txtExcelFile.Text.Trim().Substring(txtExcelFile.Text.Trim().LastIndexOf(@"\") + 1, txtExcelFile.Text.Trim().Length - txtExcelFile.Text.Trim().LastIndexOf(@"\") - 1) + @".txt";

              //  MessageBox.Show(TXTBOM);
              //  return;




              //  //屏蔽掉系统跳出的Alert
              //  app.AlertBeforeOverwriting = false;
              //  //保存到指定目录
              // // _wbk.SaveAs(appFolder + @"\.1.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
              // //     Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
              //  _wbk.SaveAs(appFolder + @"\1.xlsx");


              //  app.Quit();
              //  //释放掉多余的excel进程
              //  System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
              //  app = null;
              //  MessageBox.Show("OK");


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


                //get ColumnName 
                for (int i = 0; i < ds.Tables[0].Columns.Count; i++)
                {
                   // MessageBox.Show(ds.Tables[0].Columns[i].ToString());
                    if (ds.Tables[0].Columns[i].ToString().ToUpper().Trim() == "LEVEL")
                        _Level = i;
                    if (ds.Tables[0].Columns[i].ToString().ToUpper().Trim() == "Part Number".ToUpper ())
                        _PartNumber  = i;
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


                string TXTBOM = appFolder + @"\" + txtExcelFile.Text.Trim().Substring(txtExcelFile.Text.Trim().LastIndexOf(@"\") + 1, txtExcelFile.Text.Trim().Length - txtExcelFile.Text.Trim().LastIndexOf(@"\") - 1) + @".txt";
                

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


                for (int i = 0; i < ds.Tables[0].Columns.Count ; i++)
                {
                    for (int j = 0; j < ds.Tables[0].Rows.Count ; j++)
                    {
                        _sLevel = ds.Tables[0].Rows[j][_Level].ToString();
                        _sLocation = ds.Tables[0].Rows[j][_Location].ToString();
                        if ( _sLevel  == "1"  && !string.IsNullOrEmpty (_sLocation ))
                        {
                            _sPartNumber = ds.Tables[0].Rows[j][_PartNumber ].ToString();
                            _sPartDescription = ds.Tables[0].Rows[j][_PartDescription].ToString();
                            _sQty = ds.Tables[0].Rows[j][_Qty].ToString();
                                                     
                            if (_sLocation.Contains(","))
                            {
                                foreach (string  item in _sLocation.Split (','))
                                {
                                    _sLine = _sPartNumber.PadRight(36) + _sPartDescription.PadRight(85) + _sQty.PadRight(5) + item;
                                }
                            }
                            else
                            {
                                _sLine = _sPartNumber.PadRight(36) + _sPartDescription.PadRight(85) + _sQty.PadRight(5) + _sLocation;
                            }
                            StreamWriter sw = new StreamWriter(TXTBOM, true, Encoding.UTF8);
                            sw.WriteLine(_sLine);
                            sw.Close();
                        }
                    }

                }


                MessageBox.Show("Complete!,File save in'" + TXTBOM + "'", "Save OK",MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                txtExcelFile.SelectAll();
                txtExcelFile.Focus();
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
            if (this.Height == 411)
                this.Height = 122;
            else
                this.Height = 411;
                  
        }


    }
}
