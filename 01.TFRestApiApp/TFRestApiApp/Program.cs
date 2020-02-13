using Microsoft.TeamFoundation.Build.WebApi;
using Microsoft.TeamFoundation.Core.WebApi;
using Microsoft.TeamFoundation.SourceControl.WebApi;
using Microsoft.TeamFoundation.TestManagement.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;
using System;
using System.Data.SqlClient;
using System.Net;
using System.Configuration;
using System.Text;
using System.Data;

namespace TFRestApiApp
{
    class Program
    {
        static readonly string TFUrl = "http://tfs-srv:8080/tfs/DefaultCollection/";
        static readonly string UserAccount = "";
        static readonly string UserPassword = "";
        static readonly string UserPAT = "";

        static readonly string TN_NewRecordsInAIRT = "NewRecordsInAIRT";
        public static int reportType_1 = 1; //Where New Records not created in AIRT
        public static int reportType_2 = 2; //Where New Records created in AIRT

        static WorkItemTrackingHttpClient WitClient;
        static BuildHttpClient BuildClient;
        static ProjectHttpClient ProjectClient;
        static GitHttpClient GitClient;
        static TfvcHttpClient TfvsClient;
        static TestManagementHttpClient TestManagementClient;

        static void Main(string[] args)
        {
            string applicationType = ConfigurationManager.AppSettings["ApplicationType"];
            Console.WriteLine("<=========CSEO Accessbility Assessments ADO Features automater Creater Started=========>");
            System.Data.DataTable dtTotalAppNames = connectAndGetDataFromAIRTDB(TN_NewRecordsInAIRT, ConfigurationManager.AppSettings["TodayRecordsInAIRT"]);
            if (dtTotalAppNames.Rows.Count == 0)
            {
                Console.WriteLine("//====//====//====Today New records not created in AIRT====//====//====//");
                sendMail(reportType_1);
            } else
            {
                ConnectWithPAT(TFUrl, UserPAT);
            }
            
        }

        #region sendEmail
        private static void sendMail(int reporttype)
        {
            Console.WriteLine("//====//====//====Sending Mail for No new records are creating in AIRT====//====//====//");
            StringBuilder builder = new StringBuilder();
            if (reporttype == 1)
            {
                builder.Append("<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 3.2//EN\">");
                builder.Append("<html xmlns='http://www.w3.org/1999/xhtml'>");
                builder.Append("<head>");
                builder.Append("<title>");
                builder.Append(ConfigurationManager.AppSettings["TodayRecordsInAIRT_Type_1"].ToString() + new DateTime());
                builder.Append("</title>");
                builder.Append("<style type=\"text/css\">");
                builder.Append("table.MsoNormalTable{font-size:10.0pt;font-family:\"Calibri\",serif;}" +
                    "p.MsoNormal{margin-bottom:.0001pt;font-size:12.0pt;font-family:\"Calibri\",serif;margin-left: 0in;margin-right: 0in;margin-top: 0in;}" +
                    "h1{margin-right:0in;margin-left:0in;font-size:24.0pt;font-family:\"Calibri\",serif;font-weight:bold;}" +
                    "a:link{color:#0563C1;text-decoration:underline;text-underline:single;}p{margin-right:0in;margin-left:0in;font-size:12.0pt;font-family:\"Calibri\",serif;}");
                builder.Append("table.imagetable{font-family: verdana,Calibri,sans-serif;font-size:11px;" +
                    "color:#333333;border-width: 1px;border-color: #999999;border-collapse: collapse;}");
                builder.Append("table.imagetable th {background:#0070C0 url('cell-blue.jpg');border-width: 1px;padding: 8px;border-style: solid;border-color: #999999;}");
                builder.Append("table.imagetable td {background:#FFFFFF url('cell-grey.jpg');border-width: 1px;" +
                    "padding: 8px;border-style: solid;border-color: #999999;width:1px;white-space:nowrap;}</style>");
                builder.Append("</head>");
                builder.Append("<span style=\"font-size:11.0pt;mso-fareast-font-family:&quot;Times New Roman&quot;\">Hi All,</span><br><br/>");
                builder.Append("<span style=\"font-size:11.0pt;mso-fareast-font-family:&quot;Times New Roman&quot;\">There is no records are created in AIRT.</span><br><br>");
                builder.Append("<span style=\"mso-fareast-font-family:&quot;Times New Roman&quot;\"><u5:p></u5:p><o:p></o:p></span>");
                builder.Append("<u5:p></u5:p>");
                builder.Append("<body>");
                builder.Append("<span>Thanks</span>");
                builder.Append("<span>CSEO Accessbility Services</span>");
                builder.Append("<p style=\"color:blue;\">*** <i>This is an automatically generated email</i> ***</p>");
                builder.Append("</body>");
                builder.Append("</html>");
            }
            else
            {
                builder.Append("<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 3.2//EN\">");
                builder.Append("<html xmlns='http://www.w3.org/1999/xhtml'>");
                builder.Append("<head>");
                builder.Append("<title>");
                builder.Append(ConfigurationManager.AppSettings["TodayRecordsInAIRT_Type_1"].ToString() + new DateTime());
                builder.Append("</title>");
                builder.Append("<style type=\"text/css\">");
                builder.Append("table.MsoNormalTable{font-size:10.0pt;font-family:\"Calibri\",serif;}" +
                    "p.MsoNormal{margin-bottom:.0001pt;font-size:12.0pt;font-family:\"Calibri\",serif;margin-left: 0in;margin-right: 0in;margin-top: 0in;}" +
                    "h1{margin-right:0in;margin-left:0in;font-size:24.0pt;font-family:\"Calibri\",serif;font-weight:bold;}" +
                    "a:link{color:#0563C1;text-decoration:underline;text-underline:single;}p{margin-right:0in;margin-left:0in;font-size:12.0pt;font-family:\"Calibri\",serif;}");
                builder.Append("table.imagetable{font-family: verdana,Calibri,sans-serif;font-size:11px;" +
                    "color:#333333;border-width: 1px;border-color: #999999;border-collapse: collapse;}");
                builder.Append("table.imagetable th {background:#0070C0 url('cell-blue.jpg');border-width: 1px;padding: 8px;border-style: solid;border-color: #999999;}");
                builder.Append("table.imagetable td {background:#FFFFFF url('cell-grey.jpg');border-width: 1px;" +
                    "padding: 8px;border-style: solid;border-color: #999999;width:1px;white-space:nowrap;}</style>");
                builder.Append("</head>");
                builder.Append("<span style=\"font-size:11.0pt;mso-fareast-font-family:&quot;Times New Roman&quot;\">Hi All,</span><br><br/>");
                builder.Append("<span style=\"font-size:11.0pt;mso-fareast-font-family:&quot;Times New Roman&quot;\">Below are newly created records in AIRT</span><br><br>");
                builder.Append("<span style=\"mso-fareast-font-family:&quot;Times New Roman&quot;\"><u5:p></u5:p><o:p></o:p></span>");
                builder.Append("<u5:p></u5:p>");
                builder.Append("<body>");
                builder.Append("<table class=\"imagetable\">");
                builder.Append("<tr bgcolor=\"#0F2BDE\">");
                builder.Append("<th><font color=\"#FFFFFF\">Record ID</font></th>" +
                "<th><font color=\"#FFFFFF\">Digital Property</font></th>" +
                "<th><font color=\"#FFFFFF\">ComponentID</font></th>" +
                "<th><font color=\"#FFFFFF\">CreatedDate</font></th>" +
                "<th><font color=\"#FFFFFF\">Group</font></th>" +
                "<th><font color=\"#FFFFFF\">SubGroup</font></th>" +
                "<th><font color=\"#FFFFFF\">Priority</font></th>" +
                "<th><font color=\"#FFFFFF\">Grade</font></th>" +
                "<th><font color=\"#FFFFFF\">SubGroup Accessbility Lead</font></th>" +
                "<th><font color=\"#FFFFFF\">Eng Onwer</font></th>"+
                "<th><font color=\"#FFFFFF\">ADO Feature Link</font></th>" +
                "<th><font color=\"#FFFFFF\">Comments</font></th>"
                );
                
                //foreach(DataRow row in dtTotalAppNames.Rows)
                //{

                //}

                builder.Append("</table>");
                builder.Append("<span>Thanks</span>");
                builder.Append("<span>CSEO Accessbility Services</span>");
                builder.Append("<p style=\"color:blue;\">*** <i>This is an automatically generated email</i> ***</p>");
                builder.Append("</body>");
                builder.Append("</html>");
            }
            string HtmlFile = builder.ToString();
            Microsoft.Office.Interop.Outlook.Application oApp = new Microsoft.Office.Interop.Outlook.Application();
            Microsoft.Office.Interop.Outlook.MailItem oMsg = (Microsoft.Office.Interop.Outlook.MailItem)oApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
            /* Live */
            //oMsg.To = ConfigurationManager.AppSettings["DftSendMailTo"];
            //oMsg.CC = ConfigurationManager.AppSettings["DftSendMailCC"];
            /* QA*/
            oMsg.To = "v-chvak@microsoft.com";
            oMsg.Subject = ConfigurationManager.AppSettings["TodayRecordsInAIRT_Type_1"].ToString() + " : " + System.DateTime.Now.ToShortDateString() + " " + System.DateTime.Now.ToShortTimeString();
            oMsg.HTMLBody = HtmlFile;
            Console.WriteLine("Bugs report excel attached sending mail to CSEO group team");
            oMsg.Send();
        }
        #endregion

        #region AIRT DB Connection
        private static System.Data.DataTable connectAndGetDataFromAIRTDB(String tableName, string QueryName)
        {
            Console.WriteLine("Enter to Get Data From AIRT Live Db...........");
            System.Data.DataTable dt = new System.Data.DataTable();
            SqlCommand cmd = new SqlCommand();
            string dbConn = null;
            //dbConn = @"Data Source = xxxxxxx.database.windows.net; user id=xxxxxxxx; password=xxxxxxxx; Initial Catalog = xxxxxxx;";
            cmd.CommandText = QueryName;
            Console.WriteLine("Executes {0}", QueryName);
            SqlConnection sqlConnection1 = new SqlConnection(dbConn);
            cmd.Connection = sqlConnection1;
            sqlConnection1.Open();
            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            sda.Fill(dt);
            dt.TableName = tableName;
            sqlConnection1.Close();
            Console.WriteLine("In {0} table identified {1} rows and {2} colums.......",dt.TableName, dt.Rows.Count, dt.Columns.Count);
            return dt;
        }
            #endregion

        #region ADO client connection

            static void InitClients(VssConnection Connection)
        {
            WitClient = Connection.GetClient<WorkItemTrackingHttpClient>();
            BuildClient = Connection.GetClient<BuildHttpClient>();
            ProjectClient = Connection.GetClient<ProjectHttpClient>();
            GitClient = Connection.GetClient<GitHttpClient>();
            TfvsClient = Connection.GetClient<TfvcHttpClient>();
            TestManagementClient = Connection.GetClient<TestManagementHttpClient>();
        }

        //ADO connector through servie ADO URL
        static void ConnectWithDefaultCreds(string ServiceURL)
        {
            VssConnection connection = new VssConnection(new Uri(ServiceURL), new VssCredentials());
            InitClients(connection);
        }

        //ADO connector through Custom credentails
        static void ConnectWithCustomCreds(string ServiceURL, string User, string Password)
        {
            VssConnection connection = new VssConnection(new Uri(ServiceURL), new WindowsCredential(new NetworkCredential(User, Password)));
            InitClients(connection);
        }

        //ADO connector through PAT key token
        static void ConnectWithPAT(string ServiceURL, string PAT)
        {
            Console.WriteLine("//..//...//..//Connection {0} ADO Instance through PAT token {1}//..//...//..//", ServiceURL, PAT);
            VssConnection connection = new VssConnection(new Uri(ServiceURL), new VssBasicCredential(string.Empty, PAT));
            InitClients(connection);
        }
        #endregion
    }
}
