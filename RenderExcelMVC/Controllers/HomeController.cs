using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace RenderExcelMVC.Controllers
{
    public class HomeController : Controller
    {
        DataTable dtSheet3 = new DataTable();
        public ActionResult Index()
        {
            return View();
        }
        [HttpPost]
        public ActionResult Index(HttpPostedFileBase postedFile, HttpPostedFileBase importedFile) {

            //Declare variables to be used
            string path = Server.MapPath("~/uploads/");
            string filePath1 = string.Empty;
            string filePath2 = string.Empty;
            string extension1 = string.Empty;
            string extension2 = string.Empty;
            ViewBag.Message = string.Empty;
            DataTable dtSheet1 = new DataTable();
            DataTable dtSheet2 = new DataTable();
            DataTable dtSheet4 = new DataTable();
            DataSet ExcelData = new DataSet();
            //Check whether file 1 is posted and store that in server
            if (postedFile != null) {
                {
                    if (!Directory.Exists(path)) {
                        Directory.CreateDirectory(path);
                    }
                    filePath1 = path + Path.GetFileName(postedFile.FileName);
                    extension1 = Path.GetExtension(postedFile.FileName);
                    postedFile.SaveAs(filePath1);
                }
            }
            //Check whether file 2 is posted and store that in server
            if (importedFile != null)
            {
                {
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }
                    filePath2 = path + Path.GetFileName(importedFile.FileName);
                    extension2 = Path.GetExtension(importedFile.FileName);
                    importedFile.SaveAs(filePath2);
                }
            }
            //
            dtSheet1 = getFileData(filePath1);
            dtSheet2 = getFileData(filePath2);
            dtSheet4 = getmergedTable(dtSheet1, dtSheet2);
            //  dtSheet3 = getTableData(dtSheet1, dtSheet2);
            dtSheet3 = getTableData(dtSheet1, dtSheet4);
            ExcelData.Tables.Add(dtSheet3);
            
            if (ExcelData.Tables.Count > 0) {
                ViewBag.Message = "File Uploaded Successfully!";
            }
            return View(ExcelData);
            
        }

        [HttpGet]
        public ActionResult Reset() {
            dtSheet3 = null;
             return View(dtSheet3);
           // return null;
        }

        //Read the excel file on filePath
        public DataTable getFileData(string filePath) {
            string connectionString = String.Empty;
            DataTable dtSheet = new DataTable();
            connectionString = ConfigurationManager.ConnectionStrings["Excelload"].ConnectionString;

            connectionString = String.Format(connectionString, filePath);

            using (OleDbConnection connExcel = new OleDbConnection(connectionString))
            {
                using (OleDbCommand cmdExcel = new OleDbCommand())
                {
                    using (OleDbDataAdapter odaExcel = new OleDbDataAdapter())
                    {
                        cmdExcel.Connection = connExcel;
                        connExcel.Open();
                        DataTable dtExcelSchema;
                        dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                        string sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                        connExcel.Close();

                        connExcel.Open();
                        cmdExcel.CommandText = "SELECT * FROM [" + sheetName + "]";
                        odaExcel.SelectCommand = cmdExcel;
                        odaExcel.Fill(dtSheet);
                        connExcel.Close();

                    }
                }

            }
            return dtSheet;
        }

        // Get the table datas from both the table and populate them as key-value pair
        public DataTable getTableData(DataTable table1, DataTable table2) { 
        DataTable dtTable = new DataTable();
            //Create Dictionary to retrieve FirmName from FirmID
            Dictionary<string, string> dict = new Dictionary<string, string>();
            foreach (DataRow dr in table1.Rows) {

                dict.Add(dr["Firm ID"].ToString(), dr["Firm Name"].ToString());
            }
        dtTable.Clear();
        dtTable.Columns.Add(new DataColumn("Firm Name",typeof(string)));
        dtTable.Columns.Add(new DataColumn("AGS", typeof(string)));
        dtTable.Columns.Add(new DataColumn("BTC", typeof(string)));
        dtTable.Columns.Add(new DataColumn("CL", typeof(string)));
        dtTable.Columns.Add(new DataColumn("COAL", typeof(string)));
        dtTable.Columns.Add(new DataColumn("ELECTRICITY", typeof(string)));
        dtTable.Columns.Add(new DataColumn("EQ", typeof(string)));
        dtTable.Columns.Add(new DataColumn("FX", typeof(string)));
        dtTable.Columns.Add(new DataColumn("IR", typeof(string)));
        dtTable.Columns.Add(new DataColumn("METALS", typeof(string)));
        dtTable.Columns.Add(new DataColumn("NG", typeof(string)));
            

            
            int index = table2.Columns.IndexOf("Asset Class Name");

            var matchingrows = from r in table2.Rows.OfType<DataRow>()
                               group r by r["Firm ID"] into g
                               select new { Firm = g.Key, Data = g };

            foreach (var matchingrow in matchingrows) {
                string ags = string.Empty;
                string btc = string.Empty; string cl = string.Empty; string coal = string.Empty; string electricity = string.Empty; string eq = string.Empty;
                string fx = string.Empty; string ir = string.Empty; string metals = string.Empty; string ng = string.Empty;
                foreach (var item in matchingrow.Data) {
                    if(item.ItemArray[index].ToString().ToUpper().Equals("AGS"))
                     ags =  " X ";
                    else if(item.ItemArray[index].ToString().ToUpper().Equals("BTC")) 
                     btc = " X " ;
                    else if(item.ItemArray[index].ToString().ToUpper().Equals("CL")) 
                     cl = " X ";
                    else if(item.ItemArray[index].ToString().ToUpper().Equals("COAL")) 
                     coal = " X ";
                    else if(item.ItemArray[index].ToString().ToUpper().Equals("ELECTRICITY")) 
                     electricity = " X " ;
                    else if(item.ItemArray[index].ToString().ToUpper().Equals("EQ")) 
                     eq =   " X " ;
                    else if(item.ItemArray[index].ToString().ToUpper().Equals("FX")) 
                     fx =  " X " ;
                    else if(item.ItemArray[index].ToString().ToUpper().Equals("IR")) 
                     ir =   " X " ;
                    else if (item.ItemArray[index].ToString().ToUpper().Equals("METALS"))
                     metals = " X ";
                    else if(item.ItemArray[index].ToString().ToUpper().Equals("NG")) 
                     ng =  " X " ;
                }
                var matchingfirm = dict[matchingrow.Firm.ToString()];
                dtTable.Rows.Add(matchingfirm, ags, btc, cl, coal, electricity, eq, fx,ir,metals,ng);
            }
            dtTable.DefaultView.Sort = "Firm Name";
            dtTable=dtTable.DefaultView.ToTable();
        return dtTable;
        }
        //get the datatable for display.
        public DataTable getmergedTable(DataTable table1, DataTable table2) {
            var dtblResult = new DataTable();
            dtblResult.Columns.Add(new DataColumn("Firm ID", typeof(string)));
            dtblResult.Columns.Add(new DataColumn("Firm Name", typeof(string)));
            dtblResult.Columns.Add(new DataColumn("Asset Class ID", typeof(string)));
            dtblResult.Columns.Add(new DataColumn("Asset Class Name", typeof(string)));
            dtblResult.Columns.Add(new DataColumn("Interested Firms (ID)", typeof(string)));
            //dtblResult.Columns.Add(new DataColumn("Firm Name", typeof(string)));

            var result = from rowLeft in table1.AsEnumerable()
                         join rowRight in table2.AsEnumerable() on rowLeft["Firm ID"] equals rowRight["Interested Firms (ID)"] into gj
                         from subRight in gj.DefaultIfEmpty()
                         select rowLeft.ItemArray.Concat((subRight == null) ? (table2.NewRow().ItemArray) : subRight.ItemArray).ToArray();


            foreach (var dataRow in result)
            {
                dtblResult.Rows.Add(dataRow);
            }
            return dtblResult;
        }
    }
}