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
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using System.Linq;
using System.Collections.Generic;
using Microsoft.VisualStudio.Services.WebApi.Patch.Json;
using Microsoft.VisualStudio.Services.WebApi.Patch;

namespace TFRestApiApp
{
    class Program
    {
        static readonly string TFUrl = "https://smsgaccessibilityreviews.visualstudio.com/";
        static readonly string UserAccount = "";
        static readonly string UserPassword = "";
        static readonly string UserPAT = "";
        static readonly string teamProject = "CDSVSO";
        static readonly string workitemType = "Feature";

        static readonly string TN_NewRecordsInAIRT = "NewRecordsInAIRT";
        public static int reportType_1 = 1; //Where New Records not created in AIRT
        public static int reportType_2 = 2; //Where New Records created in AIRT

        static WorkItemTrackingHttpClient WitClient;
        static BuildHttpClient BuildClient;
        static ProjectHttpClient ProjectClient;
        static GitHttpClient GitClient;
        static TfvcHttpClient TfvsClient;
        static TestManagementHttpClient TestManagementClient;

        public static string WorkitemType_Feature_Title;
        public static string RecordID;
        public static string Group;
        public static string SubGroup;
        public static string ApplicationName;
        public static string Priority;
        public static string AreaPath = "CDSVSO"; //For QA area path is 'CDSVSO' for Dev Area path is 'OneITVSO\Shared Experiences\Studio\Accessibility\Accessibility PM'
        public static string shortcut_SubGroup;
        public static string shortcut_Group;
        public static string WorkitemType_Feature_Tag_NewAIRTRec = "NewAIRTRec";
        public static string WokitemType_Feature_Tag_Priority;
        public static string WorkitemType_Feature_AssignTo = "v-chvak@microsoft.com";

        public static string NewRecordAirt_URL;
        public static string AIRTPrtURL = "https://airt.azurewebsites.net/InventoryDetails/";
        public static string AIRTAppUrl = "https://aka.ms/A11yPriorities/";
        public static string WorkitemType_Feature_Tags;

        public static List<string> List_Group = new List<string>();
        public static List<string> List_SubGroup = new List<string>();
        public static List<string> List_AppPriority = new List<string>();

        public static DataTable dtTotalAppNames = new DataTable();

        static void Main(string[] args)
        {
            string applicationType = ConfigurationManager.AppSettings["ApplicationType"];
            Console.WriteLine("<=========CSEO Accessbility Assessments ADO Features automater Creater Started=========>");

            ConnectWithPAT(TFUrl, UserPAT);
            string recid_Pattern = "'RecID_";
            
            dtTotalAppNames = connectAndGetDataFromAIRTDB(TN_NewRecordsInAIRT, ConfigurationManager.AppSettings["TodayRecordsInAIRT"]);

            if (dtTotalAppNames.Rows.Count == 0)
            {
                Console.WriteLine("//====//====//====Today New records not created in AIRT====//====//====//");
                sendMail(reportType_1);
            } else
            {
                
                //get the count for newly created records in ALRT
                List<int> strSummRecID = new List<int>(dtTotalAppNames.Rows.Count);

                //Add the Newly added record ides in list
                foreach(DataRow row in dtTotalAppNames.Rows)
                {
                    strSummRecID.Add((int)row["RecId"]);
                }
                
                foreach(int newRecid in strSummRecID)
                {
                    //Get active ADO features based on record id
                    string queryWiqlList = @"SELECT [System.Id] FROM WorkItems WHERE [System.TeamProject] = @project" +
              " and [System.WorkItemType] = 'Feature' and [System.Title] Contains Words " +
              recid_Pattern + newRecid.ToString() + "'" +
              "and [System.State] <> 'Removed' and [System.State] <> 'Closed'";

                    WorkitemType_Feature_Title = "Recid="+newRecid.ToString();
                    GetQueryResult(queryWiqlList, teamProject);
                }
            }
        }

        /// <summary>
        /// Run query and show result
        /// </summary>
        /// <param name="wiqlStr">Wiql String</param>
        static void GetQueryResult(string wiqlStr, string teamProject)
        {
            System.Data.DataTable dt_FeatureFieldValues = new System.Data.DataTable();
            DataColumn dc_FeatureFieldValues = new DataColumn("ADO_Feature_ID");
            dt_FeatureFieldValues.Columns.Add(dc_FeatureFieldValues);
            WorkItemQueryResult result = RunQueryByWiql(wiqlStr, teamProject);

            if (result != null)
            {
                if (result.WorkItems != null) // this is Flat List 
                {
                    if (result.WorkItems.Count() != 0)
                    {
                        foreach (var wiRef in result.WorkItems)
                        {
                            var wi = GetWorkItem(wiRef.Id);
                            Console.WriteLine(String.Format("{0} - {1} - {2}", wi.Id, wi.Fields["System.Title"].ToString(), wi.Fields["System.State"].ToString()));
                        }
                    }
                    else //If workitem type feature not available based on Record id create the new feature
                    {
                        try{
                           //Create Workitem Type Feature and Get the ID
                            int bugId = CreateNewWorkitem(teamProject, workitemType);
                        }
                        catch(Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                            Console.WriteLine(ex.StackTrace);
                        }
                    }
                } else if (result.WorkItemRelations != null) // this is Tree of Work Items or Work Items and Direct Links
                {
                    foreach (var wiRel in result.WorkItemRelations)
                    {
                        if (wiRel.Source == null)
                        {
                            var wi = GetWorkItem(wiRel.Target.Id);
                            Console.WriteLine(String.Format("Top Level: {0} - {1}", wi.Id, wi.Fields["System.Title"].ToString()));
                        }
                        else
                        {
                            var wiParent = GetWorkItem(wiRel.Source.Id);
                            var wiChild = GetWorkItem(wiRel.Target.Id);
                            Console.WriteLine(String.Format("{0} --> {1} - {2}", wiParent.Id, wiChild.Id, wiChild.Fields["System.Title"].ToString()));
                        }
                    }
                }
                else Console.WriteLine("There is no query result");
            }
        }

        //Create New Workitem type
        static int CreateNewWorkitem(String ProjectName, string WorkitemType)
        {
            DataTable dt_AllNewRecordsInAIRT = connectAndGetDataFromAIRTDB(TN_NewRecordsInAIRT, ConfigurationManager.AppSettings["TodayRecordsInAIRT"]);

            DataRow[] rowsFiltered = dt_AllNewRecordsInAIRT.Select(WorkitemType_Feature_Title);

            for(int i=0; i< rowsFiltered.Length; i++)
            {
                //Get the Record id
                RecordID = rowsFiltered[i]["Recid"].ToString();

                //Get the Group
                Group = rowsFiltered[i]["Grp"].ToString();

                //Group shortcut form 
                shortcut_Group = string.Concat(Group.Where(c => c >= 'A' && c <= 'Z'));

                //Get the subgorup
                SubGroup = rowsFiltered[i]["SubGrp"].ToString();
                
                //Sub Group Shortcut form
                shortcut_SubGroup = string.Concat(SubGroup.Where(c => c >= 'A' && c <= 'Z'));

                //Get the application name
                ApplicationName = rowsFiltered[i]["NameDesc"].ToString();

                WorkitemType_Feature_Title = "[RecID_" + RecordID + "]" + "[" + shortcut_Group + "]" + "[" + shortcut_SubGroup + "]"+ ApplicationName;
                
                //Remove the empty [] expressions feature title that means when Grop or SubGroup or AppName not available 
                var charsToRemove = new string[] { "[]"};
                foreach(var c in charsToRemove)
                {
                    //Final Workitem type feature title
                    WorkitemType_Feature_Title = WorkitemType_Feature_Title.Replace(c, string.Empty);
                    Console.WriteLine("Workitem type Feature Generated for {0} record. the Title is {1}", RecordID, WorkitemType_Feature_Title);
                }

                //Get Application priority
                Priority = rowsFiltered[i]["Priority"].ToString();
                WokitemType_Feature_Tag_Priority = "P" + Priority;

                //Create the AIRT New record URL 
                NewRecordAirt_URL = AIRTPrtURL+ RecordID;

                WorkitemType_Feature_Tags = String.Concat(WokitemType_Feature_Tag_Priority,";", WorkitemType_Feature_Tag_NewAIRTRec,";", shortcut_Group);

            }
      
            Dictionary<string, object> fields = new Dictionary<string, object>();
            fields.Add("Title", WorkitemType_Feature_Title);
            fields.Add("Tags", WorkitemType_Feature_Tags);
            //fields.Add("Tags", WokitemType_Feature_Tag_Priority);
            //fields.Add("Tags", WorkitemType_Feature_Tag_NewAIRTRec);
            //fields.Add("Repro Steps", "<ol><li>Run app</li><li>Crash</li></ol>");
            fields.Add("Priority", Priority);
            fields.Add("Description", "<p>You are receiving this notification because you have a new "+
                "AIRT record that was created based on its related Service Tree record.</p>"+
                "<ul><li><a href=" + NewRecordAirt_URL + ">" + RecordID + "</a></li></ul>"+
                "<ul><li>For this new AIRT record, we need you to review and fill out the details."+
                "Noting that with a record in edit mode, all grayed out fields are locked to Service Tree."+
                "To update those fields, you need to update the ST record.</li></ul>"+
                "<ul><li>· If this project has no UI (example, is a service), then please go into the related ST record and " +
                "update the Has UI field from Yes to No. Within 8 hours of doing that, this AIRT record will be deleted.</li>" +
                "<li>You can find the related ST component link from within this AIRT record, within the Service Tree Alignment section</li>" +
                "<li>AIRT record fields to update:</li>" +
                "<ul><li>Digital Property section: Review and set the Priority field based " +
                "on <a href=" + AIRTAppUrl + ">https://aka.ms/A11yPriorities.</a> Fill in the other required fields," +
                "like Internet Facing (Internal/External), Category, etc.</li>" +
                "<li>Organization Details section: Review and fill in fields, replacing any fields with ST New - Verify</li>" +
                "<li>Dependency Information section: If the project engineering is purely done by Engineers within CSEO, then " +
                "leave this section blank. Otherwise, please review and update with 1st party and/or 3rd party dependencies.</li>" +
                "<ul><li>Any Microsoft projects/applications/components that your project has a dependency on, " +
                "please add that as a 1st party dependency</li>" +
                "<li>Any non-Microsoft projects/applications/components that your project has a dependency on, " +
                "please add that as a 3rd party dependency.</li></ul>" +
                "<li>Assessment section: This is about Accessibility testing. If you’ve done some Accessibility testing, " +
                "please add that, along with supporting documentation as an attachment to the assessment. Example would be " +
                "running a fast pass on your web site, with Accessibility Insights</li></ul>" +
                "</ul><br/><p><b>Thanks</br>CSEO Engineering System</b></p>");

            fields.Add("Assigned To", WorkitemType_Feature_AssignTo);
            fields.Add("Business Value", 999);
            var newBug = CreateWorkItem(ProjectName, WorkitemType, fields);
            //UpdateWorkItemLink(11242, newBug.Id.Value, RelConstants.ParrentRefStr);
            return newBug.Id.Value;
        }

        /// <summary>
        /// Create a work item
        /// </summary>
        /// <param name="ProjectName"></param>
        /// <param name="WorkItemTypeName"></param>
        /// <param name="Fields"></param>
        /// <returns></returns>
        static WorkItem CreateWorkItem(string ProjectName, string WorkItemTypeName, Dictionary<string, object> Fields)
        {
            JsonPatchDocument patchDocument = new JsonPatchDocument();

            foreach (var key in Fields.Keys)
                patchDocument.Add(new JsonPatchOperation()
                {
                    Operation = Operation.Add,
                    Path = "/fields/" + key,
                    Value = Fields[key]
                });

            return WitClient.CreateWorkItemAsync(patchDocument, ProjectName, WorkItemTypeName).Result;
        }

        /// <summary>
        /// Run Query with Wiql
        /// </summary>
        /// <param name="wiqlStr">Wiql String</param>
        /// <returns></returns>
        static WorkItemQueryResult RunQueryByWiql(string wiqlStr, string teamProject)
        {
            Wiql wiql = new Wiql();
            wiql.Query = wiqlStr;

            if (teamProject == "") return WitClient.QueryByWiqlAsync(wiql).Result;
            else return WitClient.QueryByWiqlAsync(wiql, teamProject).Result;
        }

        /// <summary>
        /// Get one work item
        /// </summary>
        /// <param name="Id"></param>
        /// <returns></returns>
        static WorkItem GetWorkItem(int Id)
        {
            return WitClient.GetWorkItemAsync(Id).Result;
        }

        /// <summary>
        /// Add Relations
        /// </summary>
        /// <param name="WIId"></param>
        /// <param name="References"></param>
        /// <returns></returns>
        static WorkItem AddWorkItemRelations(int WIId, List<object> References)
        {
            JsonPatchDocument patchDocument = new JsonPatchDocument();

            foreach (object rf in References)
                patchDocument.Add(new JsonPatchOperation()
                {
                    Operation = Operation.Add,
                    Path = "/relations/-",
                    Value = rf
                });

            return WitClient.UpdateWorkItemAsync(patchDocument, WIId).Result; // return updated work item
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
            dbConn = @"";
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

       

        class RelConstants
        {
            //https://docs.microsoft.com/en-us/azure/devops/boards/queries/link-type-reference?view=vsts

            public const string RelatedRefStr = "System.LinkTypes.Related";
            public const string ChildRefStr = "System.LinkTypes.Hierarchy-Forward";
            public const string ParrentRefStr = "System.LinkTypes.Hierarchy-Reverse";
            public const string DuplicateRefStr = "System.LinkTypes.Duplicate-Forward";
            public const string DuplicateOfRefStr = "System.LinkTypes.Duplicate-Reverse";
            public const string SuccessorRefStr = "System.LinkTypes.Dependency-Forward";
            public const string PredecessorRefStr = "System.LinkTypes.Dependency-Reverse";
            public const string TestedByRefStr = "Microsoft.VSTS.Common.TestedBy-Forward";
            public const string TestsRefStr = "Microsoft.VSTS.Common.TestedBy-Reverse";
            public const string TestCaseRefStr = "Microsoft.VSTS.TestCase.SharedStepReferencedBy-Forward";
            public const string SharedStepsRefStr = "Microsoft.VSTS.TestCase.SharedStepReferencedBy-Reverse";
            public const string AffectsRefStr = "Microsoft.VSTS.Common.Affects.Forward";
            public const string AffectedByRefStr = "Microsoft.VSTS.Common.Affects.Reverse";
            public const string AttachmentRefStr = "AttachedFile";
            public const string HyperLinkRefStr = "Hyperlink";
            public const string ArtifactLinkRefStr = "ArtifactLink";

            public const string LinkKeyForDict = "<NewLink>"; // key for dictionary to separate a link from fields            
        }
    }
}
