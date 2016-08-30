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
using System.Net.Mail;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using System.Text;
using System.Net;

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
            int totalRows = Excel.Rows.Count;
            int currentRow = 2;
            XLWorkbook wb = new XLWorkbook();
            var ws = wb.Worksheets.Add(Excel, "WorksheetName");
            for (int i = 0; i < totalRows; i++)
            {
                string currentCell = "F" + currentRow;
                string Formula = "=D" + currentRow + "*" + "E" + currentRow;
                ws.Cells(currentCell).FormulaA1 = Formula;
                currentRow++;
            }
            HttpResponse httpResponse = Response;
            httpResponse.Clear();
            httpResponse.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            httpResponse.AddHeader("content-disposition", "attachment;filename=\"SalesReport.xlsx\"");

            // Flush the workbook to the Response.OutputStream
            using (MemoryStream memoryStream = new MemoryStream())
            {
                wb.SaveAs(memoryStream);
                memoryStream.WriteTo(httpResponse.OutputStream);
                byte[] bytes = memoryStream.ToArray();
                memoryStream.Close();
                MailMessage mm = new MailMessage("anruevinswork@gmail.com", "anruevinsnew@gmail.com");
                mm.Subject = "Sales Report";
                mm.Body = "HI";
                mm.Attachments.Add(new Attachment(new MemoryStream(bytes), "SalesReport.xlsx"));
                mm.IsBodyHtml = true;
                SmtpClient smtp = new SmtpClient();
                smtp.Host = "smtp.gmail.com";
                smtp.EnableSsl = true;
                smtp.UseDefaultCredentials = false;
                NetworkCredential NetworkCred = new NetworkCredential();
                NetworkCred.UserName = "anruevinswork@gmail.com";
                NetworkCred.Password = "ZZHFZ3lZhDb3bioCiqXA";

                smtp.Credentials = NetworkCred;
                smtp.Port = 587;
                smtp.Send(mm);
            }
            httpResponse.End();
        }


        protected void Button1_Click(object sender, EventArgs e)
        {
            DateTime datefrom = DateTime.Parse(Request.Form[TextBox1.UniqueID]);
            DateTime dateto = DateTime.Parse(Request.Form[TextBox2.UniqueID]);
            //getting datatable from viewstate  
            DataTable dt = (DataTable)ViewState["DataTable"];
            string filterExp = "OrderDate  > '" + datefrom.Date + "' and OrderDate < '" + dateto.Date + "'";
            DataRow[] foundrows;
            foundrows = dt.Select(filterExp);
            DataTable newdt = new DataTable();
            newdt.Columns.Add("OrderID", typeof(Int32));
            newdt.Columns.Add("OrderDate", typeof(DateTime));
            newdt.Columns.Add("Name", typeof(string));
            newdt.Columns.Add("Quantity", typeof(Int32));
            newdt.Columns.Add("UnitPrice", typeof(decimal));
            foreach (DataRow row in foundrows)
            {
                newdt.ImportRow(row);
            }
            //adding new column which calculates total amount of money that is paid for an order
            DataColumn TotalAmount = new DataColumn();
            TotalAmount.DataType = System.Type.GetType("System.Decimal");
            TotalAmount.ColumnName = "TotalAmount";
            //TotalAmount.Expression = "UnitPrice * Quantity";
            newdt.Columns.Add(TotalAmount);
            //calling create Excel File Method and ing dataTable   
            CreateExcelFile(newdt);
        }
    }
}