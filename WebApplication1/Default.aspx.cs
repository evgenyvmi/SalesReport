using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.IO;
using System.Configuration;
using System.Data.SqlClient;

namespace WebApplication1
{
    public partial class Default : System.Web.UI.Page
    {
        private SqlConnection con;
        private SqlCommand com;
        private string constr, query;
        private void connection()
        {
            constr = ConfigurationManager.ConnectionStrings["getconn"].ToString();
            con = new SqlConnection(constr);
            con.Open();

        }
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                Bindgrid();

            }
        }


        private void Bindgrid()
        {
            connection();
            query = @"select 
                            OrderDetail.OrderID,""Order"".OrderDate, Product.Name, Product.UnitsOnOrder, Product.UnitPrice
                    from 
                            OrderDetail, Product, ""Order"" 
                    where 
                        (
                        OrderDetail.ProductID=Product.ID 
                        and OrderDetail.OrderID=""Order"".ID
                        and Product.UnitsOnOrder>0
                        )"
                                                                                    ;//not recommended this i have written just for example,write stored procedure for security  
            com = new SqlCommand(query, con);
            SqlDataAdapter da = new SqlDataAdapter(query, con);
            DataSet ds = new DataSet();
            da.Fill(ds);
            
            GridView1.DataSource = ds;
            GridView1.DataBind();
            con.Close();
            ViewState["DataTable"] = ds.Tables[0];
        }


        public void CreateExcelFile(DataTable Excel)
        {

            //Clears all content output from the buffer stream.  
            Response.ClearContent();
            //Adds HTTP header to the output stream  
            Response.AddHeader("content-disposition", string.Format("attachment; filename=SalesReport.xls"));

            // Gets or sets the HTTP MIME type of the output stream  
            Response.ContentType = "application/vnd.ms-excel";
            string space = "";

            foreach (DataColumn dcolumn in Excel.Columns)
            {

                Response.Write(space + dcolumn.ColumnName);
                space = "\t";
            }
            Response.Write("\n");
            int countcolumn;
            foreach (DataRow dr in Excel.Rows)
            {
                space = "";
                for (countcolumn = 0; countcolumn < Excel.Columns.Count; countcolumn++)
                {

                    Response.Write(space + dr[countcolumn].ToString());
                    space = "\t";

                }

                Response.Write("\n");


            }
            Response.End();
        }


        protected void Button1_Click(object sender, EventArgs e)
        {
            //getting datatable from viewstate  
            DataTable dt = (DataTable)ViewState["DataTable"];

            //calling create Excel File Method and ing dataTable   
            CreateExcelFile(dt);
        }
    }
}