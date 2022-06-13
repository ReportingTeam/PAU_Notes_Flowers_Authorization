using System;
using System.Data;
using System.Data.SqlClient;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using CommonFunctions;
using System.Linq;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using System.Text.Json;
using System.Collections.Generic;

namespace PAU_Notes_Flowers_Authorization
{
    class Program
    {

        static void Main(string[] args)
        {
            string strSql = null;
            string strExcelReport = CF.Folder.Reports + "\\PAU_Notes_Flowers_Authorization\\" +
                "PAU_Notes_Flowers_Authorization_" + DateTime.Now.ToString("yyyy-MM-dd_HHmmss") + ".xlsx";
            string strTemp = null;
            int intLastFlowersNoteId;
            SqlConnection cnnProductivity = null;
            SqlCommand cmdProductivity = null;
            SqlDataAdapter daProductivity = null;
            DataTable dt = null;
            DataTable dtTemp = null;
            Excel.Application appExcel = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            System.IO.FileInfo fi = null;
            HttpClient hc = new HttpClient();
            Task<HttpResponseMessage> task = null;
            CF.Email eml = null;

            try
            {
                //log start
                CF.WriteToLog("START", "Report started.", CF.DatabaseName.COMMON);

                //create excel app
                appExcel = new Excel.Application();
                Thread.Sleep(1000);

                //open template
                wb = appExcel.Workbooks.Add();
                Thread.Sleep(1000);

                //get worksheet
                ws = wb.Worksheets[1];

                //open connections
                cnnProductivity = CF.OpenSqlConnectionWithRetry(CF.GetConnectionString(CF.DatabaseName.PRODUCTIVITY), 10);

                //get last note id
                strSql = "SELECT LASTFLOWERSNOTEID FROM PAU_LAST_FLOWERS_NOTE_ID";
                daProductivity = new SqlDataAdapter(strSql, cnnProductivity);
                daProductivity.Fill(dt = new DataTable());
                intLastFlowersNoteId = Convert.ToInt32(dt.Rows[0]["LASTFLOWERSNOTEID"]);

                //get data
                strSql =
                    "SELECT " +

                    "ID AS 'NoteId', " +
                    "ACCOUNT AS 'Account', " +
                    "HOSPITALPATIENTNAME AS 'Patient Name', " +
                    "INSURANCETYPE AS 'Insurance Type', " +
                    "INSURANCENAME AS 'Insurance Name', " +
                    "INSURANCEPLAN AS 'Insurance Plan', " +
                    "'' AS 'CPT Code' " +

                    "FROM PAU_NOTE " +

                    "WHERE " +
                    "ID>'" + intLastFlowersNoteId + "' AND " +
                    "FACILITYNUMBER=65 AND " +
                    "AUTHORIZATIONREQUIRED='Y' " +
                    "ORDER BY ID";
                daProductivity = new SqlDataAdapter(strSql, cnnProductivity);
                daProductivity.Fill(dt = new DataTable());

                //get last flowers note id
                if (dt.Rows.Count > 0) 
                {
                    intLastFlowersNoteId = dt.AsEnumerable().Max(x => x.Field<int>("NOTEID"));
                }

                //add cpt code
                foreach (DataRow dr in dt.Rows)
                {
                    strTemp = "";
                    strSql =
                        "SELECT " +
                        "CPTCODE " +
                        "FROM PAU_CPT_CODE " +
                        "WHERE NOTEID='" + dr["NOTEID"].ToString() + "'";
                    daProductivity = new SqlDataAdapter(strSql, cnnProductivity);
                    daProductivity.Fill(dtTemp = new DataTable());
                    foreach(DataRow drTemp in dtTemp.Rows)
                    {
                        strTemp = strTemp + drTemp["CPTCODE"].ToString() + ",";
                    }
                    if (strTemp.Length > 0)
                    {
                        dr["CPT CODE"] = strTemp.Substring(0, strTemp.Length - 1);
                    }
                }

                //remove note id
                dt.Columns.RemoveAt(0);
                dt.AcceptChanges();

                //export to excel
                CF.DataTableToExcel(dt, ws, true, 1, 1);

                //autofit
                ws.Columns.AutoFit();

                //save excel
                wb.SaveAs(strExcelReport);
                wb.Close();

                //update last note id
                strSql =
                   "UPDATE PAU_LAST_FLOWERS_NOTE_ID " +
                   "SET LASTFLOWERSNOTEID=" + intLastFlowersNoteId;
                cmdProductivity = new SqlCommand(strSql, cnnProductivity);
                cmdProductivity.ExecuteNonQuery();

                //readonly
                fi = new System.IO.FileInfo(strExcelReport);
                fi.IsReadOnly = true;

                //email
                eml = new CF.Email();

                //from
                eml.From = "NashvilleReportingTeam@ssc-nashville.com";

                //to
                eml.To.Add("8175_CI_NashRptTeam_" + System.Reflection.Assembly.GetCallingAssembly().GetName().Name +
                   "@ssc-nashville.com");

                //subject
                eml.Subject = "PAU Notes Authorization Report for Flowers " +
                    DateTime.Now.ToString("yyyy-MM-dd_HHmmss");

                //body
                eml.Body = CF.GetEmailMessage("PAU Notes Authorization Report for Flowers", null, null);

                //attachments
                eml.Attachments.Add(new CF.Email.Attachment
                {
                    FileName = Path.GetFileName(strExcelReport),
                    Bytes = File.ReadAllBytes(strExcelReport)
                });

                //send it
                task = hc.PostAsJsonAsync<string>("http://10.5.72.172:108/api/webemail",
                    JsonSerializer.Serialize(eml));
                task.Wait();
                if (!task.Result.IsSuccessStatusCode)
                {
                    //error
                    CF.WriteToLog("ERROR", task.Result.ReasonPhrase, CF.DatabaseName.COMMON);
                    Environment.Exit(0);
                }

                //write to log
                CF.WriteToLog("FINISH", "Report finished.", CF.DatabaseName.COMMON);

            }
            catch (Exception ex)
            {
                CF.WriteToLog("ERROR", ex.ToString(), CF.DatabaseName.COMMON);
                Environment.Exit(0);
            }
            finally
            {
                appExcel.Quit();
                CF.KillApp(appExcel.Hwnd);
                if (cnnProductivity != null) cnnProductivity.Dispose();
                if (daProductivity != null) daProductivity.Dispose();
                if (dt != null) dt.Dispose();
            }

        }

    }

}
