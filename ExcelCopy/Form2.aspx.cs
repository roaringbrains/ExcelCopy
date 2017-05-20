using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;


namespace ExcelCopy
{
    public partial class Form2 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void btnUpload_Click(object sender, EventArgs e)
        {
            if (FileUploadControl.HasFile)
            {
                string Ext = Path.GetExtension(FileUploadControl.PostedFile.FileName);
                if (Ext == ".xls" || Ext == ".xlsx")
                {
                    lblErrorMessage.Visible = false;
                    string Name = Path.GetFileName(FileUploadControl.PostedFile.FileName);
                    string FolderPath = ConfigurationManager.AppSettings["FolderPath"];
                    string FilePath = Server.MapPath(FolderPath + Name);
                    FileUploadControl.SaveAs(FilePath);
                    FillGridFromExcelSheet(FilePath, Ext, ddlIsHeaderExists.SelectedValue);
                }
                else
                {
                    lblErrorMessage.Visible = true;
                    lblErrorMessage.InnerText = "Please upload valid Excel File";
                    ExcelGridView.DataSource = null;
                    ExcelGridView.DataBind();
                }
            }

        }
        private void FillGridFromExcelSheet(string FilePath, string ext, string isHader)
        {
            string connectionString = "";
            if (ext == ".xls")
            {   //For Excel 97-03
                connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = 'D:\\Test1.xlsx' ; Extended Properties = 'Excel 8.0;HDR={1}'";
            }
            else if (ext == ".xlsx")
            {    //For Excel 07 and greater
                connectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source ='D:\\Test1.xlsx' ; Extended Properties = 'Excel 8.0;HDR={1}'";
            }
            connectionString = String.Format(connectionString, FilePath, isHader);
            OleDbConnection conn = new OleDbConnection(connectionString);
            OleDbCommand cmd = new OleDbCommand();
            OleDbDataAdapter dataAdapter = new OleDbDataAdapter();
            DataTable dt = new DataTable();
            cmd.Connection = conn;
            //Fetch 1st Sheet Name
            conn.Open();
            DataTable dtSchema;
            dtSchema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            string ExcelSheetName = dtSchema.Rows[0]["TABLE_NAME"].ToString();
            conn.Close();
            //Read all data of fetched Sheet to a Data Table
            conn.Open();
            cmd.CommandText = "SELECT * From [Sheet1$]";
            dataAdapter.SelectCommand = cmd;
            dataAdapter.Fill(dt);
            conn.Close();
            //Bind Sheet Data to GridView
            ExcelGridView.Caption = Path.GetFileName(FilePath);
            ExcelGridView.DataSource = dt;
            ExcelGridView.DataBind();
        }



    }
}