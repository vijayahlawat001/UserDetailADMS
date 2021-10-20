using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace EPSDS_ADMS
{
    public partial class Service1 : ServiceBase
    {
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            try
            {
                WriteToFile("Service is started at " + DateTime.Now);
                System.Threading.ThreadStart job = new System.Threading.ThreadStart(ReportSchedulerJob);
                System.Threading.Thread thread = new System.Threading.Thread(job);
                thread.Start();
            }
            catch (Exception ex)
            {
                WriteToFile("Error Occured :- " + ex.Message + " at time " + DateTime.Now);
            }
            finally
            {
                System.Threading.Thread.Sleep(1000 * 60);
            }
        }
        protected override void OnStop()
        {
            WriteToFile("Service is stopped at " + DateTime.Now);
        }
        private void ReportSchedulerJob()
        {
            try
            {
                try
                {
                    WriteToFile("OnElapsedTime Started at " + DateTime.Now);
                    String ssFile = "D:/Book1.xlsx";
                    var workbook = SpreadsheetGear.Factory.GetWorkbook(ssFile);
                    var worksheet = workbook.Worksheets[0];
                    SpreadsheetGear.IRange usedRange = worksheet.Range;
                    DataTable dt = new DataTable();
                    dt.Columns.Add("USER_ID", typeof(System.String));
                    dt.Columns.Add("DisplayName", typeof(System.String));
                    dt.Columns.Add("FIRST_NAME", typeof(System.String));
                    dt.Columns.Add("LAST_NAME", typeof(System.String));
                    dt.Columns.Add("EMAIL_ID", typeof(System.String));
                    dt.Columns.Add("DEPARTMENT", typeof(System.String));
                    dt.Columns.Add("DESIGNATION", typeof(System.String));
                    dt.Columns.Add("PHONE_NO", typeof(System.String));
                    dt.Columns.Add("ACTIVE", typeof(System.String));
                    dt.Columns.Add("BLOCKED", typeof(System.String));
                    long rows = worksheet.UsedRange.Rows.Count;
                    long cells = worksheet.UsedRange.ColumnCount;
                    DataRow dr;
                    for (int i = 1; i < rows - 1; i++)
                    {
                        dr = dt.NewRow();
                        for (int j = 0; j < cells; j++)
                        {
                            dr[j] = usedRange[i, j].Text;
                        }
                        dt.Rows.Add(dr);
                    }
                    DBAccess db = new DBAccess();
                    SqlCommand cmd = new SqlCommand();
                    cmd.CommandText = "FETCH_DATA_FROM_ADMS";
                    string data;
                    dt.TableName = "DATA";
                    using (StringWriter sw = new StringWriter())
                    {
                        dt.WriteXml(sw);
                        data = sw.ToString();
                    }
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@P_DATA", data);
                    db.GetRcdSetByCmdTrans(cmd);
                    WriteToFile("Updates updated at " + DateTime.Now);
                }
                catch (Exception ex)
                {
                    WriteToFile("Error Occured inner catch:- " + ex.Message + " at time " + DateTime.Now);
                }
                finally
                {
                    System.Threading.Thread.Sleep(1000 * 60 * 1);
                }
            }
            catch (Exception ex)
            {
                WriteToFile("Error Occured :- " + ex.Message + " at time " + DateTime.Now);
            }
        }
        public void WriteToFile(string Message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            if (!File.Exists(filepath))
            {
                // Create a file to write to.   
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
        }
    }
}
