using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using System.Configuration;
using System.IO;

namespace UploadBio
{
    public partial class frmUploadBio : Form
    {
        Cls_Upload oUpload = new Cls_Upload();

        public frmUploadBio()
        {
            InitializeComponent();
        }

        private void btnUpload_Click(object sender, EventArgs e)
        {
            UploadFTP();
        }

        private void frmUploadBio_Load(object sender, EventArgs e)
        {
            //Export();
            UploadFTP();

            Application.Exit();
        }

        private void UploadFTP()
        {
            try
            {
                //Method 1
                Upload(ConfigurationSettings.AppSettings["ftplocation"], "administrator", "HpMl350", ConfigurationSettings.AppSettings["filelocation"]);


                //Method 2

                ////////Get the object used to communicate with the server.
                //////FtpWebRequest request = (FtpWebRequest)WebRequest.Create("ftp://192.168.1.2/BioDB/MCT/att2000.mdb");
                //////request.Method = WebRequestMethods.Ftp.UploadFile;

                ////////This assumes the FTP site uses anonymous logon.
                ////////request.Credentials = new NetworkCredential("anonymous", "renju@premierlogisticsgroup.com");

                ////////Copy the contents of the file to the request stream.           
                //////byte[] b = File.ReadAllBytes("D://BioDB/att2000.mdb");

                //////request.ContentLength = b.Length;
                //////using (Stream s = request.GetRequestStream())
                //////{
                //////    s.Write(b, 0, b.Length);
                //////}

                //////FtpWebResponse ftpResp = (FtpWebResponse)request.GetResponse();

                //////if (ftpResp != null)
                //////{
                //////    lblMsg.Text=ftpResp.StatusDescription;
                //////}

                //////ftpResp.Close();
            }
            catch (Exception ex)
            {
                if (ConfigurationSettings.AppSettings["ftplocation"].ToString() == "Yes")
                {
                    MessageBox.Show("Error at UploadFTP - " + ex.ToString());
                }
            }
            finally
            { }
        }

        private static void Upload(string ftpServer, string userName, string password, string filename)
        {
            try
            {
                using (System.Net.WebClient client = new System.Net.WebClient())
                {
                    client.Credentials = new System.Net.NetworkCredential(userName, password);
                    client.UploadFile(ftpServer + "/" + new FileInfo(filename).Name, "STOR", filename);
                }
            }
            catch (Exception ex)
            {
                if (ConfigurationSettings.AppSettings["ftplocation"].ToString() == "Yes")
                {
                    MessageBox.Show("Error at Upload- " + ex.ToString());
                }
            }
            finally
            { }
        }

        private void Export()
        {
            try
            {
                //Get the last log time when the log was fetched last time
                //string text = System.IO.File.ReadAllText(@"D:\BioDB\LastLog.txt");
                string readText = System.IO.File.ReadAllText(ConfigurationSettings.AppSettings["lastlogfilelocation"] + "LastLog.txt");

                //Get the Last Log DateTime.
                DateTime lastLog = DateTime.Parse(readText);

                DataTable dtBioLogs = oUpload.getEmpLogs();
                dgvwLog.DataSource = dtBioLogs;

                ExportToExcel();

                //Write the last log time to the text file
                string writeText = oUpload.getLastLogTime().ToString();
                // Creates a file, writes the specified string to the file, and then closes the file.
                System.IO.File.WriteAllText(ConfigurationSettings.AppSettings["lastlogfilelocation"] + "LastLog.txt", writeText);
            }
            catch (Exception ex)
            {
                if (ConfigurationSettings.AppSettings["ftplocation"].ToString() == "Yes")
                {
                    MessageBox.Show("Error - " + ex.ToString());
                }
            }
            finally
            { }
        }

        private void ExportToExcel()
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            ExcelApp.Application.Workbooks.Add(Type.Missing);

            // Change properties of the Workbook 
            //ExcelApp.Columns.ColumnWidth = 20;

            // Storing header part in Excel
            for (int i = 1; i < dgvwLog.Columns.Count + 1; i++)
            {
                ExcelApp.Cells[1, i] = dgvwLog.Columns[i - 1].HeaderText;
            }

            // Storing Each row and column value to excel sheet
            for (int i = 0; i < dgvwLog.Rows.Count; i++)
            {
                for (int j = 0; j < dgvwLog.Columns.Count; j++)
                {
                    ExcelApp.Cells[i + 2, j + 1] = dgvwLog.Rows[i].Cells[j].Value.ToString();
                }
            }
            //MessageBox.Show("Excel file created , you can find the file D:\\Employee Data.xlsx");
            //ExcelApp.Visible = true;

            //To save to a destination
            //ExcelApp.ActiveWorkbook.SaveCopyAs("D:\\BioDB\\EmpBio.xlsx");
            ExcelApp.ActiveWorkbook.SaveCopyAs(ConfigurationSettings.AppSettings["savelocation"] + "EmpBio.xls");
            ExcelApp.ActiveWorkbook.Saved = true;

            ExcelApp.Quit();
        } 
    }
}
