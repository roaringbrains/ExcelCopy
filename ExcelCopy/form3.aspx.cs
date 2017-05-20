using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using System.Data;
using System.Data.OleDb;


namespace ExcelCopy
{
    public partial class form3 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void btnImport_Click(object sender, EventArgs e)
        {
            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =C:\\Users\\Alpha\\documents\\visual studio 2015\\Projects\\ExcelCopy\\ExcelCopy\\Files\\may2.xlsx; Extended Properties = 'Excel 8.0;HDR=Yes'";
            OleDbConnection conn = new OleDbConnection(connectionString);
            OleDbCommand cmd = new OleDbCommand();
            OleDbDataAdapter dataAdapter = new OleDbDataAdapter();
            System.Data.DataTable dt = new System.Data.DataTable();
            DataSet ds = new DataSet();
            cmd.Connection = conn;
            //Fetch 1st Sheet Name
            conn.Open();
            System.Data.DataTable dtSchema;
            dtSchema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            string ExcelSheetName = dtSchema.Rows[0]["TABLE_NAME"].ToString();
            conn.Close();
            //Read all data of fetched Sheet to a Data Table
            conn.Open();
            cmd.CommandText = "SELECT * From [" + ExcelSheetName + "] ";
            dataAdapter.SelectCommand = cmd;
            dataAdapter.Fill(dt);
            conn.Close();

            System.Data.DataTable dt1 = CopyTable(dt);

            GridView1.DataSource =  dt1;
            GridView1.DataBind();


            StreamWriter wr = new StreamWriter("D:\\Book1.xls");

            try
            {
                for (int i = 0; i < dt1.Columns.Count; i++)
                {
                    wr.Write(dt1.Columns[i].ToString().ToUpper() + "\t");
                   
                }

                wr.WriteLine();
                for (int i = 0; i < (dt1.Rows.Count); i++)
                {
                    for (int j = 0; j < dt1.Columns.Count; j++)
                    {
                        if (j == 3)
                        {
                            if (dt1.Rows[i][j].ToString() == "False")
                            {
                                wr.Write(Convert.ToString("F") + "\t");
                            }
                            else
                            {
                                wr.Write(Convert.ToString("T") + "\t");
                            }
                        }
                        else
                        {
                            if (dt1.Rows[i][j] != null)
                            {
                                wr.Write(Convert.ToString(dt1.Rows[i][j]) + "\t");
                            }
                            else
                            {
                                wr.Write("\t");
                            }
                        }
                    }
                    wr.WriteLine();
                }
                wr.Close();
            }

            catch (Exception ex)
            {
                throw ex;
            }

           // demo();
        }


        //public void demo()
        //{
        //    string connectionString1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source =C:\\Users\\pujakumx\\Downloads\\latest\\Book1.xls; Extended Properties = 'Excel 8.0;HDR=Yes'";
        //    OleDbConnection conn1 = new OleDbConnection(connectionString1);
        //    OleDbDataAdapter dad = new OleDbDataAdapter();
        //    System.Data.DataTable dta = new System.Data.DataTable();
        //    conn1.Open();
        //    System.Data.DataTable dtSch;

        //    dtSch = conn1.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
        //    string ExcelSheet = dtSch.Rows[0]["TABLE_NAME"].ToString();
        //    conn1.Close();
        //    //Read all data of fetched Sheet to a Data Table
        //    conn1.Open();

        //    using (OleDbCommand cmd1 = new OleDbCommand("UPDATE [" + ExcelSheet + "] SET EI MPQ = 'F' WHERE EID = '893231'", conn1))
        //    {
        //        conn1.Open();
        //        dad.SelectCommand = cmd1;
        //        dad.Fill(dta);

        //        cmd1.ExecuteNonQuery();
        //        conn1.Close();
        //    }
        //}


        public System.Data.DataTable CopyTable(System.Data.DataTable dt)
        {

            System.Data.DataTable tblFormat = new System.Data.DataTable();

            tblFormat.Columns.Add("EI ID");
            tblFormat.Columns.Add("EI Path");
            tblFormat.Columns.Add("EI Count User");
            tblFormat.Columns.Add("EI Count RR");
            tblFormat.Columns.Add("EI MPQ");
            tblFormat.Columns.Add("RR");
            tblFormat.Columns.Add("RR Counter User");
            tblFormat.Columns.Add("RR Count EID");
            tblFormat.Columns.Add("RR AD IDSID");
            tblFormat.Columns.Add("Srv AD online");
            tblFormat.Columns.Add("Srv AD Offline");
            tblFormat.Columns.Add("Srv Cust online");
            tblFormat.Columns.Add("Srv Cust offline");
            tblFormat.Columns.Add("RR Migration Flag");
            tblFormat.Columns.Add("RR AGS Aggregated");
            tblFormat.Columns.Add("Domain");
            tblFormat.Columns.Add("AD CORP CN");

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                tblFormat.Rows.Add(new string[] { dt.Rows[i][0].ToString(), dt.Rows[i][1].ToString(),
                dt.Rows[i][5].ToString(), dt.Rows[i][12].ToString()});
            }
            return tblFormat;
        }
    }
}
 