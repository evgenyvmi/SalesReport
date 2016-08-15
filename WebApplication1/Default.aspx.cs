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
using ClosedXML.Excel;
using DocumentFormat.OpenXml;

namespace WebApplication1
{
    public partial class Default : System.Web.UI.Page
    {
        private SqlConnection con;
        private SqlCommand com;
        private string constr, query;
        private void connection()
        {
            constr = ConfigurationManager.ConnectionStrings["Northwind"].ToString();
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
                            OrderDetail.OrderID,""Order"".OrderDate, Product.Name, OrderDetail.Quantity, OrderDetail.UnitPrice
                    from 
                            OrderDetail, Product, ""Order"" 
                    where 
                        (
                        OrderDetail.ProductID=Product.ID 
                        and OrderDetail.OrderID=""Order"".ID
                        )"
                                                                                    ;
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
            XLWorkbook wb = new XLWorkbook();
            var ws = wb.Worksheets.Add(Excel, "WorksheetName");
            
            HttpResponse httpResponse = Response;
            httpResponse.Clear();
            httpResponse.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            httpResponse.AddHeader("content-disposition", "attachment;filename=\"SalesReport.xlsx\"");

            // Flush the workbook to the Response.OutputStream
            using (MemoryStream memoryStream = new MemoryStream())
            {
                wb.SaveAs(memoryStream);
                memoryStream.WriteTo(httpResponse.OutputStream);
                memoryStream.Close();
            }

            httpResponse.End();
        }


        protected void Button1_Click(object sender, EventArgs e)
        {
            //getting datatable from viewstate  
            DataTable dt = (DataTable)ViewState["DataTable"];
            //adding new column which calculates total amount of money that is paid for an order
            DataColumn TotalAmount = new DataColumn();
            TotalAmount.DataType = System.Type.GetType("System.Decimal");
            TotalAmount.ColumnName = "TotalAmount";
            TotalAmount.Expression = "UnitPrice * Quantity";
            dt.Columns.Add(TotalAmount);
            //calling create Excel File Method and ing dataTable   
            CreateExcelFile(dt);
        }
    }
}