using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.Data;

namespace ExcelCopy
{
    public partial class UploadExcel : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack == true)
            {
                FileUpload2.Visible = false;
                UploadButton2.Visible = false;
                FileUpload3.Visible = false;
                UploadButton3.Visible = false;
                Calculate.Visible = false;
            }

        }

        protected void UploadButton1_Click(object sender, EventArgs e)
        {
            if (FileUpload1.HasFile)
            {
                try
                {
                    if (FileUpload1.PostedFile.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    {
                        string filename1 = Path.GetFileName(FileUpload1.FileName);
                        FileUpload1.SaveAs(Server.MapPath("~\\Files\\Old\\") + filename1);
                        StatusLabel1.Text = "Report 1 Upload successfully!!!";
                        UploadButton1.Enabled = false;
                        FileUpload1.Enabled = false;
                        FileUpload2.Visible = true;
                        UploadButton2.Visible = true;

                    }
                    else
                        StatusLabel1.Text = "Sorry You can upload only .xlsx Files !!!";

                }
                catch (Exception Ex)
                {

                    StatusLabel1.Text = "Sorry file wasn't uploaded because of " + Ex.Message;

                }
            }
        }

        protected void UploadButton2_Click(object sender, EventArgs e)
        {
            try
            {
                if (FileUpload2.HasFile)
                {
                    if (FileUpload2.PostedFile.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    {
                        string filename2 = Path.GetFileName(FileUpload2.FileName);
                        FileUpload2.SaveAs(Server.MapPath("~\\Files\\Old\\") + filename2);
                        StatusLabel2.Text = "Report 2 Upload successfully!!!";
                        FileUpload2.Enabled = false;
                        UploadButton2.Enabled = false;
                        FileUpload3.Visible = true;
                        UploadButton3.Visible = true;
                    }
                    else
                        StatusLabel2.Text = "Sorry You can upload only .xlsx Files !!!";
                }
            }
            catch (Exception Ex)
            {
                StatusLabel2.Text = "Sorry file wasn't uploaded because of " + Ex.Message;
            }
        }


        protected void UploadButton3_Click(object sender, EventArgs e)
        {
            try
            {
                if (FileUpload3.HasFile)
                {
                    if (FileUpload3.PostedFile.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    {
                        string filename3 = Path.GetFileName(FileUpload3.FileName);
                        FileUpload3.SaveAs(Server.MapPath("~\\Files\\Old\\") + filename3);
                        StatusLabel3.Text = "DevTools ATTRI file Upload successfully!!!";
                        FileUpload3.Enabled = false;
                        UploadButton3.Enabled = false;
                        Calculate.Visible = true;
                    }
                    else
                        StatusLabel2.Text = "Sorry You can upload only .xlsx Files !!!";
                }
            }
            catch (Exception Ex)
            {
                StatusLabel2.Text = "Sorry file wasn't uploaded because of " + Ex.Message;
            }

        }

        protected void Calculate_Click(object sender, EventArgs e)
        {

            //Read data from DevTool file - Change the file path according to your server config
            string conStrFile3 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =C:\\Users\\Alpha\\Documents\\visual studio 2015\\Projects\\ExcelCopy\\ExcelCopy\\Files\\Old\\DevTools ATTRI.xlsx; Extended Properties = 'Excel 8.0;HDR=Yes'";

            OleDbConnection con3 = new OleDbConnection(conStrFile3);
            OleDbCommand cmd3 = new OleDbCommand();
            OleDbDataAdapter dA3 = new OleDbDataAdapter();
            DataTable dT3 = new DataTable();
            cmd3.Connection = con3;
            con3.Open();
            cmd3.CommandText = "SELECT * From [EAM$] ";
            dA3.SelectCommand = cmd3;
            dA3.Fill(dT3);
            //GridView1.DataSource = dT3;
            //GridView1.DataBind();
            //FinalStatus.Text = dT3.Rows[1][1].ToString();

            con3.Close();


            //Read data from Report 2 file - Change the file path according to your server config
            string conStrFile2 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =C:\\Users\\Alpha\\Documents\\visual studio 2015\\Projects\\ExcelCopy\\ExcelCopy\\Files\\Old\\Report2.xlsx; Extended Properties = 'Excel 8.0;HDR=Yes'";
            // On hold as of now 


            //Read data from  Report 1 file - Change the file path according to your server config
            string conStrFile1 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =C:\\Users\\Alpha\\Documents\\visual studio 2015\\Projects\\ExcelCopy\\ExcelCopy\\Files\\Old\\Report1.xlsx; Extended Properties = 'Excel 8.0;HDR=Yes'";
            OleDbConnection con1 = new OleDbConnection(conStrFile1);
            OleDbCommand cmd1 = new OleDbCommand();
            OleDbDataAdapter dA1 = new OleDbDataAdapter();
            DataTable dT1 = new DataTable();
            cmd1.Connection = con1;
            con1.Open();
            cmd1.CommandText = "SELECT * From [Sheet1$] ";
            dA1.SelectCommand = cmd1;
            dA1.Fill(dT1);
            con1.Close();


            //Working on: "DevTools ATTRI"[0][0] Sheet "EAM" to Fill data from "Report 1"[2][0] for Column "DataProvider"
            //starting row column number = "DevTools ATTRI"[0][0]
            //starting row column number = "Report 1"[2][0]
            //Start : column mapping DataProvider = lev1admins
            for (int i = 0; i < (dT3.Rows.Count); i++)
            {
                for (int j = 2; j < (dT1.Rows.Count); j++)
                {
                    if (dT3.Rows[i][0].ToString() == dT1.Rows[j][0].ToString())
                    {
                        //Here you can write the logic to copy the data in column(lev1admin=dT1.Rows[j][42].ToString()) to any where
                        //i am tranfering it to data table 3 you can again open the DEVTOOL file and write this data to 
                        //particular row and column
                        dT3.Rows[i][2] = dT1.Rows[j][42].ToString();

                    }
                }
            }
            //End : column mapping DataProvider = lev1admins

            //----------------------------------------------------------------------------------------------------------------//
            
            //Start : column mapping Name = EIName
            for (int i = 0; i < (dT3.Rows.Count); i++)
            {
                for (int j = 2; j < (dT1.Rows.Count); j++)
                {
                    if (dT3.Rows[i][0].ToString() == dT1.Rows[j][0].ToString())
                    {
                        //Here you can write the logic to copy the data in column(EINAME=dT1.Rows[j][3].ToString()) to any where
                        //i am tranfering it to data table 3 you can again open the DEVTOOL file and write this data to 
                        //particular row and column

                        dT3.Rows[i][4] = dT1.Rows[j][3].ToString();

                    }
                }
            }
            //End : column mapping DataProvider = lev1admins




            GridView1.DataSource = dT3;
            GridView1.DataBind();




        }
    }
}