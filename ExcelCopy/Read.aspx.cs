using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Data.OleDb;
using System.Data;
using System.IO;

namespace ExcelCopy
{
    public partial class Read : System.Web.UI.Page
    {
        OleDbConnection con;
        OleDbCommand com;
        OleDbDataAdapter oda;
        DataSet ds;
        DataTable dataTable;

        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source =D:\\Test1.xlsx; Extended Properties = 'Excel 8.0;HDR=Yes'");
            com = new OleDbCommand("Select * from [Sheet1$]", con);
            con.Open();
            oda = new OleDbDataAdapter(com);
            ds = new DataSet();
            oda.Fill(ds);
            con.Close();
            FileData.DataSource = ds;
            FileData.DataBind();
            Label1.Text = "Completed";
            
        }
        public override void VerifyRenderingInServerForm(Control control)
        {
            //required to avoid the runtime error "  
            //Control 'GridView1' of type 'GridView' must be placed inside a form tag with runat=server."  
        }
        protected void Button2_Click(object sender, EventArgs e)
        {
            //Method : 1
            Response.Clear();
            Response.Buffer = true;
            Response.ClearContent();
            Response.ClearHeaders();
            Response.Charset = "";
            string FileName = "Exported.xls";
            StringWriter strwritter = new StringWriter();
            HtmlTextWriter htmltextwrtter = new HtmlTextWriter(strwritter);
            Response.Cache.SetCacheability(HttpCacheability.NoCache);

            Response.ContentType = "application/vnd.ms-excel";
            Response.AddHeader("Content-Disposition", "attachment;filename=" + FileName);
            FileData.GridLines = GridLines.Both;
            FileData.HeaderStyle.Font.Bold = true;
            FileData.RenderControl(htmltextwrtter);
            Response.Write(strwritter.ToString());
            Response.End();



            ////Method : 2 
            //Response.Clear();
            //Response.Buffer = true;
            //Response.AddHeader("content-disposition",
            // "attachment;filename=DataTable.xls");
            //Response.Charset = "";
            //Response.ContentType = "application/vnd.ms-excel";
            //StringWriter sw = new StringWriter();
            //HtmlTextWriter hw = new HtmlTextWriter(sw);
            //for (int i = 0; i < FileData.Rows.Count; i++)
            //{
            //    //Apply text style to each Row
            //    FileData.Rows[i].Attributes.Add("class", "textmode");
            //}
            //FileData.RenderControl(hw);
            ////style to format numbers to string
            //string style = @"<style> .textmode { mso-number-format:\@; } </style>";
            //Response.Write(style);
            //Response.Output.Write(sw.ToString());
            //Response.Flush();
            //Response.End();



        }
    }
}
