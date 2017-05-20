using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using System.Data.SqlClient;
using System.Configuration;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;


namespace ExcelCopy
{
    public partial class getDataFromSql : System.Web.UI.Page
    {
        
        protected void Page_Load(object sender, EventArgs e)
        {
            SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["MoviesDb"].ConnectionString);

            SqlCommand cmd = new SqlCommand("SELECT * FROM Rental", con);
            DataSet ds = new DataSet();
            SqlDataAdapter da = new SqlDataAdapter();

            da.SelectCommand = cmd;
            con.Open();

            da.Fill(ds, "Rental");

            GridView1.DataSource = ds;
            GridView1.DataBind();

            con.Close();
            
            //SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["MoviesDB"].ConnectionString);

            //con.Open();

            //SqlCommand cmd = new SqlCommand();

            //cmd.Connection = con;
            //cmd.CommandType = CommandType.Text;
            //cmd.CommandText = "SELECT * FROM Rental";

            //SqlDataAdapter da = new SqlDataAdapter(cmd);

            //DataSet ds = new DataSet();

            //da.Fill(ds);

            //GridView1.DataSource = ds;
            //GridView1.DataBind();

            //con.Close();

        }

        public void Button1_Click(object sender, EventArgs e)
        {
                
        }
    }
}