using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using System.Data.SqlClient;
using System.Net;
using System.Configuration;
using System.Windows.Forms;
using System.Timers;
using System.Data.Common;

namespace AutomatedEmailAlert_
{
    public partial class AutomatedEmailAlert : Form
    {
        public AutomatedEmailAlert()
        {
            InitializeComponent();
        }
        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            this.button1_Click(null, null);
        }
        private static void Timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            Environment.Exit(0);
        }
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                using (SqlConnection con = new SqlConnection("Server=192.9.200.129; Database=inforerpdb; User Id=dev1; Password=Dev1@12345;"))
                {
                    string MailToBuyer_Indenter = "";
                    string sRecord = @"select distinct a.Buyer, a.PONo,a.Indenter,a.Division, COUNT(*) from AutomatedEmailAlert a 
                                      inner join AutomatedEmailAlert b on a.PONo= b.PONo
                                      where a.Indenter= b.Indenter and a.POLine=b.POLine
                                      group by a.PONo,a.Indenter,a.Buyer,a.Division
                                      order by a.PONo,a.Indenter,a.Buyer,a.Division";
                     if (con.State != ConnectionState.Open)
                    {
                        con.Open();
                    }
                    SqlTransaction tran = con.BeginTransaction();
                    SqlCommand cmd = new SqlCommand(sRecord, con, tran);
                    var dt = new DataTable();
                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        dt.Load(dr);
                    }
                    foreach (DataRow dr in dt.Rows)
                    {
                        string sBuyer = dr["Buyer"].ToString();
                        string sIndenter = dr["Indenter"].ToString();
                        string sPONo = dr["PONo"].ToString();
                        string sDivision = dr["Division"].ToString();
                     //   string sHODEmail = "";
                        //string sHODmail = "";

                        string sRecord2 = @"Select * from AutomatedEmailAlert where PONo='" + sPONo + "' and Indenter ='" + sIndenter + "' and Buyer= '" + sBuyer + "' order by POLine ";
                        if (con.State != ConnectionState.Open)
                        {
                            con.Open();
                        }
                        SqlCommand cmd2 = new SqlCommand(sRecord2, con, tran);
                        var dt4 = new DataTable();
                        using (SqlDataReader dr5 = cmd2.ExecuteReader())
                        {
                            dt4.Load(dr5);
                        }

                        using (SqlConnection con2 = new SqlConnection("Server=192.9.200.150; Database=IJTPerks; User Id=sa; Password=isgec12345;"))
                        {
                            if (con2.State != ConnectionState.Open)
                            {
                                con2.Open();
                            }
                             SqlTransaction tran1 = con2.BeginTransaction();
                            //string sHODCode = "select HODEmployeeCode from ProjectHODsAutoAlert where Division = '" + sDivision + "' ";
                            //SqlCommand cmdHODCode = new SqlCommand(sHODCode, con2, tran1);
                            //var dt3 = new DataTable();
                            //using (SqlDataReader dr2 = cmdHODCode.ExecuteReader())
                            //{
                            //    dt3.Load(dr2);
                            //    // return tb;
                            //}
                            string sBuyermail = "select EMailID from HRM_Employees where CardNo = " + sBuyer + " ";
                            SqlCommand cmdBuyerEmail = new SqlCommand(sBuyermail, con2, tran1);
                            string sBuyerEmail = cmdBuyerEmail.ExecuteScalar().ToString();
                            string sIndentermail = "select EMailID from HRM_Employees where CardNo = " + sIndenter + "";
                            SqlCommand cmdIndenterEmail = new SqlCommand(sIndentermail, con2, tran1);
                            string sIndenterEmail = cmdIndenterEmail.ExecuteScalar().ToString();
                            //if (dt3.Rows.Count == 1)
                            //{
                            //    sHODmail = "select EMailID from HRM_Employees where CardNo = " + dt3.Rows[0]["HODEmployeeCode"] + "";
                            //    SqlCommand cmdHODEmail = new SqlCommand(sHODmail, con2, tran1);
                            //    sHODEmail += cmdHODEmail.ExecuteScalar().ToString();
                            //    sHODEmail += ";";
                            //}
                            //else if (dt3.Rows.Count > 1)
                            //{
                            //    foreach (DataRow datarow1 in dt3.Rows)
                            //    {
                            //        sHODmail = "select EMailID from HRM_Employees where CardNo = " + datarow1["HODEmployeeCode"] + "";
                            //        SqlCommand cmdHODEmail = new SqlCommand(sHODmail, con2, tran1);
                            //        sHODEmail += cmdHODEmail.ExecuteScalar().ToString();
                            //        sHODEmail += ";";
                            //    }
                            //}
                            //else
                            //{
                            //    sHODEmail = ";";
                            //}
                            MailToBuyer_Indenter = sBuyerEmail + ";" + sIndenterEmail; 
                            //"sagar.shukla@isgec.co.in";
                            // sHODEmail = "pankaj.gupta@isgec.co.in";
                            //sBuyerEmail + ";" + sIndenterEmail;
                        }
                        //SendMail(MailToBuyer_Indenter, sHODEmail, dt4);
                        SendMail(MailToBuyer_Indenter, dt4);
                    }

                }
            }
            catch (Exception ex)
            {
                System.Timers.Timer timer = new System.Timers.Timer(5000);
                timer.Elapsed += Timer_Elapsed;
                timer.Start();
            }
            finally
            {
                System.Timers.Timer timer = new System.Timers.Timer(5000);
                timer.Elapsed += Timer_Elapsed;
                timer.Start();
            }
        }

         //;public void SendMail(string MailToBuyer_Indenter, string sHODEmail, DataTable dt4)
         public void SendMail(string MailToBuyer_Indenter, DataTable dt4)
        {
            try
            {
                MailMessage mM = new MailMessage();
                mM.From = new MailAddress("baansupport@isgec.co.in");
                foreach (var address in MailToBuyer_Indenter.Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries))
                {
                    mM.To.Add(address);
                }
                //foreach (var address in sHODEmail.Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries))
                //{
                //    mM.CC.Add(address);
                //}
                // Below two Cc Added to check the proper email functionality
                mM.CC.Add("baansupport@isgec.co.in");
                mM.CC.Add("veena@isgec.co.in");
                string Project = "";
                //for (int i = 0; i < dt4.Rows.Count; i++)
                //{

                //    if (i == 0)
                //    {
                //        Project = dt4.Rows[0]["Project"].ToString() + " - " + dt4.Rows[0]["ProjectName"].ToString();
                //    }
                //    else
                //    {
                //        //if (dt4.Rows[i - 1]["Project"].ToString() != dt4.Rows[i]["Project"].ToString())
                //        //{
                //            Project += "/'" + dt4.Rows[i]["Project"].ToString() + " - " + dt4.Rows[i]["ProjectName"].ToString();
                //        //}
                //    }

                //}

                // Display ProjectCode-ProjectName in EMail Subject line --- Change 20-09-18 
                Project = dt4.Rows[0]["Project"].ToString() + " - " + dt4.Rows[0]["ProjectName"].ToString();
                mM.Subject = "Deviation in Plan Delivery Date :- " + Project;

               // mM.Subject = "Deviation in Plan Delivery Date";

                mM.IsBodyHtml = true;
                mM.Body = "In following purchase Orders Plan delivery Date is more than requested in Indent –" + @"
               <br />
               <table class ='table' style='width:100%'  border = " + 2 + " cellpadding =" + 0 + " cellspacing = " + 1 + " width = " + 900 + " >" +
              "<thead align='center' style='border-color:grey'><tr>" +
              "<td><b>PO.No.</b></td>" +
              "<td><b>Line No</b></td>" +
              "<td><b>Supplier Code</b></td>" +
              "<td><b>Supplier Name</b></td>" +
              "<td><b>PO Plan Delievery Date(a)</b> </td>" +
               "<td><b>Indent</b></td>" +
               "<td><b>Line</b></td>" +
                "<td><b>Item Code</b></td>" +
               "<td><b>Item Description</b></td>" +
                "<td><b>Indent Plan Delievry Date(b)</b></td>" +
                "<td><b>PO Delievery Date Delayed by (a-b Days) </td>" +
              "</tr></thead>";
                foreach (DataRow dr in dt4.Rows)
                {
                    mM.Body += "<tr>" +
                     "<td>" + dr["PONo"].ToString() + "</td>" +
                     "<td>" + dr["POLine"].ToString() + "</td>" +
                     "<td>" + dr["SupplierCode"].ToString() + "</td>" +
                     "<td>" + dr["SupplierName"].ToString() + "</td>" +
                     "<td>" + dr["PlannedDeliveryDate"].ToString() + "</td>" +
                     "<td>" + dr["IndentNo"].ToString() + "</td>" +
                     "<td>" + dr["IndentLineNo"].ToString() + "</td>" +
                     "<td>" + dr["ItemCode"].ToString() + "</td>" +
                     "<td>" + dr["ItemDesc"].ToString() + "</td>" +
                     "<td>" + dr["IndentPlannedDeliveryDate"].ToString() + "</td>" +
                     "<td>" + dr["DayDiff"].ToString() + "</td>" +
                    "</tr>";
                }
                mM.Body += "</table>";

                mM.Body += "<br /><br /><br /><br /><br />Note- This is an autogenerated e-mail";
                mM.Body = mM.Body.Replace("\n", "<br />");
                SmtpClient sC = new SmtpClient("192.9.200.214", 25);
                sC.DeliveryMethod = SmtpDeliveryMethod.Network;
                sC.UseDefaultCredentials = false;
                sC.Credentials = new NetworkCredential("baansupport@isgec.co.in", "isgec");
                sC.EnableSsl = false;  // true
                sC.Timeout = 10000000;
                sC.Send(mM);

            }
            catch (Exception ex)
            {
            }
        }
    }
}
