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
        static readonly string TFUrl = "https://microsoftit.visualstudio.com/"; //Prod:https://microsoftit.visualstudio.com/ ; QA:https://smsgaccessibilityreviews.visualstudio.com
        static readonly string UserAccount = "";
        static readonly string UserPassword = "";
        static readonly string UserPAT = "";
        static readonly string teamProject = "OneITVSO"; //Prod: OneITVSO ; QA:CDSVSO
        static readonly string workitemType = "Feature";
        static readonly string workitemType2 = "User Story";

        static readonly string TN_NewRecordsInAIRT = "NewRecordsInAIRT";
        public static int reportType_1 = 1; //Where New Records not created in AIRT
        public static int reportType_2 = 2; //Where New Records created in AIRT

        static WorkItemTrackingHttpClient WitClient;
        static BuildHttpClient BuildClient;
        static ProjectHttpClient ProjectClient;
        static GitHttpClient GitClient;
        static TfvcHttpClient TfvsClient;
        static TestManagementHttpClient TestManagementClient;

        public static string RecordID;
        public static string Grade;
        public static string Group;
        public static string SubGroup;
        public static string ApplicationName;
        public static string ComponentID;
        public static string Priority;
        public static string AreaPath = "OneITVSO\\Shared Experiences\\Studio\\Accessibility\\Accessibility PM"; //QA: CDSVSO Prod: OneITVSO\Shared Experiences\Studio\Accessibility\Accessibility PM
        public static string shortcut_SubGroup;
        public static string shortcut_Group;
        public static string WorkitemType_Feature_Tag_NewAIRTRec = "NewAIRTRec";
        public static string WokitemType_Feature_Tag_Priority;
        public static string WorkitemType_Feature_AssignTo = "v-chvak@microsoft.com";//"chaseco@microsoft.com;v-juew@microsoft.com;v-chvak@microsoft.com;v-chkov@microsoft.com;v-parama@microsoft.com";

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
            Console.WriteLine("CSEO Accessbility Assessments ADO Features automater Creater Started......!");
            ConnectWithPAT(TFUrl, UserPAT);
            string recid_Pattern = "'RecID_";

            dtTotalAppNames = connectAndGetDataFromAIRTDB(TN_NewRecordsInAIRT, ConfigurationManager.AppSettings["TodayRecordsInAIRT"]);

            if (dtTotalAppNames.Rows.Count == 0)
            {
                Console.WriteLine("Today New records not created in AIRT.....!");
                sendMail(reportType_1, null);
            }
            else
            {

                //get the count for newly created records in ALRT
                List<string> dataList = new List<string>();
                //Add the Newly added record ides in list
                foreach (DataRow row in dtTotalAppNames.Rows)
                {
                    string finalString = "";
                    string nameDesc = "";
                    string componentID = "";
                    string createdDt = "";
                    string grp = "";
                    string subGrp = "";
                    string priority = "";
                    string grade = "";
                    int newRecid = (int)row["RecId"];
                    if (!DBNull.Value.Equals(row["NameDesc"]))
                        nameDesc = (string)row["NameDesc"];
                    if (!DBNull.Value.Equals(row["CreatedDt"]))
                        createdDt = (string)row["CreatedDt"].ToString();
                    if (!DBNull.Value.Equals(row["ComponentID"]))
                        componentID = (string)row["ComponentID"].ToString();
                    if (!DBNull.Value.Equals(row["Grp"]))
                        grp = (string)row["Grp"];
                    if (!DBNull.Value.Equals(row["SubGrp"]))
                        subGrp = (string)row["SubGrp"];
                    if (!DBNull.Value.Equals(row["Priority"]))
                        priority = (string)row["Priority"].ToString();
                    if (!DBNull.Value.Equals(row["Grade"]))
                        grade = (string)row["Grade"];
                    finalString = newRecid.ToString() + "," + nameDesc + "," + componentID + "," + createdDt + "," + grp + "," + subGrp + "," + priority + "," + grade + ",";
                    //Get active ADO features based on record id
                    string queryWiqlList = @"SELECT [System.Id] FROM WorkItems WHERE [System.TeamProject] = @project" +
                    " and [System.WorkItemType] = 'Feature' and [System.Title] Contains Words " +
                    recid_Pattern + newRecid.ToString() + "'" +
                    "and [System.State] <> 'Removed' and [System.State] <> 'Closed'";

                    RecordID = newRecid.ToString();
                    Grade = grade;
                    Group = grp;
                    SubGroup = subGrp;
                    ApplicationName = nameDesc;
                    ComponentID = componentID;
                    Priority = priority;
                    shortcut_Group = string.Concat(Group.Where(c => c >= 'A' && c <= 'Z'));
                    shortcut_SubGroup = string.Concat(SubGroup.Where(c => c >= 'A' && c <= 'Z'));
                    string appendString = GetQueryResult(queryWiqlList, teamProject);
                    finalString = finalString + appendString;
                    dataList.Add(finalString);
                }
                sendMail(reportType_2, dataList);
            }

        }

        /// <summary>
        /// Get one work item with information about linked work items
        /// </summary>
        /// <param name="Id"></param>
        /// <returns></returns>
        static WorkItem GetWorkItemWithRelations(int Id)
        {
            return WitClient.GetWorkItemAsync(Id, expand: WorkItemExpand.Relations).Result;
        }

        static int ExtractWiIdFromUrl(string Url)
        {
            int id = -1;

            string splitStr = "_apis/wit/workItems/";

            if (Url.Contains(splitStr))
            {
                string[] strarr = Url.Split(new string[] { splitStr }, StringSplitOptions.RemoveEmptyEntries);

                if (strarr.Length == 2 && int.TryParse(strarr[1], out id))
                    return id;
            }

            return id;
        }

        /// <summary>
        /// Run query and show result
        /// </summary>
        /// <param name="wiqlStr">Wiql String</param>
        static string GetQueryResult(string wiqlStr, string teamProject)
        {
            string appendString = "";
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
                            appendString = wi.Id.ToString() + ",Feature Already Created";
                            var wii = GetWorkItemWithRelations(wi.Id.Value);
                            List<string> userStoriesList = new List<string>();
                            foreach (var wiLink in wii.Relations)
                            {
                                if (wiLink.Rel.Equals("System.LinkTypes.Hierarchy-Forward"))
                                {
                                    int existId = ExtractWiIdFromUrl(wiLink.Url);
                                    if (existId != -1)
                                    {
                                        var wiUserStory = GetWorkItem(existId);
                                        if (wiUserStory.Fields["System.WorkItemType"].Equals("User Story"))
                                        {
                                            userStoriesList.Add(existId.ToString());
                                        }
                                    }
                                }
                            }
                            if (userStoriesList.Count == 0)
                            {
                                appendString = appendString + CreateNewWorkitem(teamProject, workitemType, wi.Id.Value);
                            }
                            else
                            {
                                appendString = appendString + "," + string.Join(":", userStoriesList) + ",User Stories Already Created";
                            }
                            Console.WriteLine(String.Format("{0} - {1} - {2}", wi.Id, wi.Fields["System.Title"].ToString(), wi.Fields["System.State"].ToString()));
                        }

                        //if (result.WorkItemRelations != null)
                        //{
                        //    foreach (var wiRel in result.WorkItemRelations)
                        //    {
                        //        if (wiRel.Source == null)
                        //        {
                        //            var wi = GetWorkItem(wiRel.Target.Id);
                        //            Console.WriteLine(String.Format("Top Level: {0} - {1}", wi.Id, wi.Fields["System.Title"].ToString()));
                        //        }
                        //        else
                        //        {
                        //            var wiParent = GetWorkItem(wiRel.Source.Id);
                        //            var wiChild = GetWorkItem(wiRel.Target.Id);
                        //            Console.WriteLine(String.Format("{0} --> {1} - {2}", wiParent.Id, wiChild.Id, wiChild.Fields["System.Title"].ToString()));
                        //        }
                        //    }
                        //}
                    }
                    else //If workitem type feature not available based on Record id create the new feature
                    {
                        try
                        {
                            //Create Workitem Type Feature and Get the ID
                            appendString = CreateNewWorkitem(teamProject, workitemType, -1);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                            Console.WriteLine(ex.StackTrace);
                        }
                    }
                }
                else
                {
                    appendString = " , , , ,";
                    Console.WriteLine("There is no query result");
                }
            }
            return appendString;
        }

        //Create New Workitem type
        static string CreateNewWorkitem(String ProjectName, string WorkitemType, int id)
        {
            string finalTagString = getFinalTagString(RecordID);
            string dataString = "";
            int parentId;
            int userStoryParentId;

            //Verify the Parent Scenario ID Mappign Based on Group and Priority
            if (shortcut_Group.Equals("CSEO"))
            {
                if (Priority.Equals("3"))
                {
                    parentId = Convert.ToInt32(ConfigurationManager.AppSettings["Grade_Review_Parent_ScenarioId"]);
                }
                else
                {
                    parentId = Convert.ToInt32(ConfigurationManager.AppSettings["CSEO_Group_Parent_ScenarioId"]);
                }
            }
            else
            {
                parentId = Convert.ToInt32(ConfigurationManager.AppSettings["NON_CSEO_Group_Parent_ScenarioId"]);
            }

            if (id == -1)
            {
                string WorkitemType_Feature_Title = "[RecID_" + RecordID + "]" + "[" + shortcut_Group + "]" + "[" + shortcut_SubGroup + "]" + ApplicationName;
                //Remove the empty [] expressions feature title that means when Grop or SubGroup or AppName not available 
                var charsToRemove = new string[] { "[]" };
                foreach (var c in charsToRemove)
                {
                    //Final Workitem type feature title
                    WorkitemType_Feature_Title = WorkitemType_Feature_Title.Replace(c, string.Empty);
                    Console.WriteLine("Workitem type Feature Generated for {0} record. the Title is {1}", RecordID, WorkitemType_Feature_Title);
                }

                WokitemType_Feature_Tag_Priority = "P" + Priority;


                //Create the AIRT New record URL 
                NewRecordAirt_URL = AIRTPrtURL + RecordID;
                if (Priority != "" && Priority != null)
                    WorkitemType_Feature_Tags = String.Concat(WokitemType_Feature_Tag_Priority, ";", WorkitemType_Feature_Tag_NewAIRTRec, ";", shortcut_Group, ";", shortcut_SubGroup);
                else
                    WorkitemType_Feature_Tags = String.Concat(WorkitemType_Feature_Tag_NewAIRTRec, ";", shortcut_Group, ";", shortcut_SubGroup);

                Dictionary<string, object> fields = new Dictionary<string, object>();
                fields.Add("Title", WorkitemType_Feature_Title);
                fields.Add("Tags", WorkitemType_Feature_Tags);
                //fields.Add("Tags", WokitemType_Feature_Tag_Priority);
                //fields.Add("Tags", WorkitemType_Feature_Tag_NewAIRTRec);
                //fields.Add("Repro Steps", "<ol><li>Run app</li><li>Crash</li></ol>");
                if (Priority != "" && Priority != null)
                    fields.Add("Priority", Priority);
                fields.Add("Description", "<p>You are receiving this notification because you have a new " +
                    "AIRT record that was created based on its related Service Tree record.</p>" +
                    "<ul><li><a href=" + NewRecordAirt_URL + ">" + RecordID + "</a></li></ul>" +
                    "<ul><li>For this new AIRT record, we need you to review and fill out the details." +
                    "Noting that with a record in edit mode, all grayed out fields are locked to Service Tree." +
                    "To update those fields, you need to update the ST record.</li></ul>" +
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
                fields.Add("Area Path", AreaPath);
                var newBug = CreateWorkItem(ProjectName, WorkitemType, fields);
                dataString = newBug.Id.Value.ToString() + ",Feature Newly Created";
                if (parentId > 0)
                {
                    AddParentLink(newBug.Id.Value, parentId);
                }
                userStoryParentId = newBug.Id.Value;

            }
            else
            {
                userStoryParentId = id;
            }
            string usrStoryCreateTag = ConfigurationManager.AppSettings["userStoryCreationTag"];
            if (usrStoryCreateTag.Equals("yes"))
            {
                Dictionary<string, object> usrStoryFields = new Dictionary<string, object>();
                string WorkitemType_Feature_Title2 = "[RecID_" + RecordID + "]" + finalTagString + "[" + shortcut_Group + "]" + "[" + shortcut_SubGroup + "]" + ApplicationName;
                var charsToRemove = new string[] { "[]" };
                foreach (var c in charsToRemove)
                {
                    //Final Workitem type feature title
                    WorkitemType_Feature_Title2 = WorkitemType_Feature_Title2.Replace(c, string.Empty);
                    Console.WriteLine("Workitem type 2 Feature Generated for {0} record. the Title is {1}", RecordID, WorkitemType_Feature_Title2);
                }
                usrStoryFields.Add("Title", "[Assessment]" + WorkitemType_Feature_Title2);
                if (Priority != "" && Priority != null)
                    usrStoryFields.Add("Priority", Priority);
                if (Priority != "" && Priority != null)
                    WorkitemType_Feature_Tags = String.Concat(WokitemType_Feature_Tag_Priority, ";", shortcut_Group, ";", shortcut_SubGroup);
                else
                    WorkitemType_Feature_Tags = String.Concat(shortcut_Group, ";", shortcut_SubGroup);

                if (Grade.Equals("C"))
                {
                    usrStoryFields.Add("Tags", WorkitemType_Feature_Tags + ";StayHealthy");
                }
                else
                {
                    usrStoryFields.Add("Tags", WorkitemType_Feature_Tags + ";GetHealthy");
                }
                usrStoryFields.Add("Area Path", AreaPath);

                var newUsrStory = CreateWorkItem(ProjectName, workitemType2, usrStoryFields);
                dataString = dataString + "," + newUsrStory.Id.Value + ",User Story Newly Created";
                if (newUsrStory.Id.Value > 0)
                {
                    AddParentLink(newUsrStory.Id.Value, userStoryParentId);
                }
            }
            else
            {
                dataString = dataString + ",,User Story Not Created";
            }

            return dataString;
        }

        static String getFinalTagString(String RecordId)
        {
            DataTable dt_AllNewRecordsInAIRT = connectAndGetDataFromAIRTDB(TN_NewRecordsInAIRT, "SELECT STUFF((SELECT ',' + CONVERT(varchar(10),TagId) FROM [dbo].[InvTagMapping] invTagMap  where invTagMap.isDeleted = 0 and invTagMap.RecId = '" + RecordID + "' FOR XML PATH('')) ,1,1,'') TagIds");
            string tagMapString = "";
            string finalTagString = "";
            Object tagObj = dt_AllNewRecordsInAIRT.Rows[0]["TagIds"];
            if (tagObj != null && !DBNull.Value.Equals(tagObj))
            {
                tagMapString = tagObj.ToString();
            }

            if (!tagMapString.Equals(""))
            {
                if (tagMapString.Contains(","))
                {
                    String[] tagArray = tagMapString.Split(',');
                    foreach (string tagString in tagArray)
                    {
                        finalTagString = finalTagString + "[" + ConfigurationManager.AppSettings[tagString] + "]";
                    }
                }
                else
                {
                    finalTagString = "[" + ConfigurationManager.AppSettings[tagMapString] + "]";
                }
            }
            return finalTagString;
        }
        /// <summary>
        /// Add new parent link to existing work item
        /// </summary>
        /// <param name="WiId"></param>
        /// <param name="ParentWiId"></param>
        /// <returns></returns>
        static WorkItem AddParentLink(int WiId, int ParentWiId)
        {
            WorkItem wi = WitClient.GetWorkItemAsync(WiId, expand: WorkItemExpand.Relations).Result;
            bool parentExists = false;

            // check existing parent link
            if (wi.Relations != null)
                if (wi.Relations.Where(x => x.Rel == RelConstants.ParentRefStr).FirstOrDefault() != null)
                    parentExists = true;

            if (!parentExists)
            {
                WorkItem parentWi = WitClient.GetWorkItemAsync(ParentWiId).Result; // get parent to retrieve its url

                Dictionary<string, object> fields = new Dictionary<string, object>();

                fields.Add(RelConstants.LinkKeyForDict + RelConstants.ParentRefStr + parentWi.Id, // to use as unique key
                CreateNewLinkObject(RelConstants.ParentRefStr, parentWi.Url, "Parent " + parentWi.Id));

                return SubmitWorkItem(fields, WiId);
            }

            Console.WriteLine("Work Item " + WiId + " contains a parent link");

            return null;
        }

        /// <summary>
        /// Create or update a work item
        /// </summary>
        /// <param name="WIId"></param>
        /// <param name="References"></param>
        /// <returns></returns>
        static WorkItem SubmitWorkItem(Dictionary<string, object> Fields, int WIId = 0, string TeamProjectName = "", string WorkItemTypeName = "")
        {
            JsonPatchDocument patchDocument = new JsonPatchDocument();

            foreach (var key in Fields.Keys)
                patchDocument.Add(new JsonPatchOperation()
                {
                    Operation = Operation.Add,
                    Path = (key.StartsWith(RelConstants.LinkKeyForDict)) ? "/relations/-" : "/fields/" + key,
                    Value = Fields[key]
                });

            if (WIId == 0) return WitClient.CreateWorkItemAsync(patchDocument, TeamProjectName, WorkItemTypeName).Result; // create new work item

            return WitClient.UpdateWorkItemAsync(patchDocument, WIId).Result; // return updated work item
        }

        /// <summary>
        /// Create a link object
        /// </summary>
        /// <param name="RelName"></param>
        /// <param name="RelUrl"></param>
        /// <param name="Comment"></param>
        /// <param name="IsLocked"></param>
        /// <returns></returns>
        static object CreateNewLinkObject(string RelName, string RelUrl, string Comment = null, bool IsLocked = false)
        {
            return new
            {
                rel = RelName,
                url = RelUrl,
                attributes = new
                {
                    comment = Comment,
                    isLocked = IsLocked // you must be an administrator to lock a link
                }
            };
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
        private static void sendMail(int reporttype, List<string> dataList)
        {
            Console.WriteLine("//====//====//====Sending Mail for No new records are creating in AIRT====//====//====//");
            StringBuilder builder = new StringBuilder();
            Microsoft.Office.Interop.Outlook.Application oApp = new Microsoft.Office.Interop.Outlook.Application();
            Microsoft.Office.Interop.Outlook.MailItem oMsg = (Microsoft.Office.Interop.Outlook.MailItem)oApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
            if (reporttype == 1)
            {
                builder.Append("<!DOCTYPE html>");
                builder.Append("<html lang='en'>");
                builder.Append("<head>");
                builder.Append("<meta charset='utf-8'>");
                builder.Append("<meta http-equiv='X-UA-Compatible' content='IE=edge'>");
                builder.Append("<meta name='viewport' content='width=device-width, initial-scale=1'>");
                builder.Append("<title>Tables</title>");
                builder.Append("<style type='text/css'> * { -webkit-box-sizing: border-box; -moz-box-sizing: border-box; box-sizing: border-box; } *, ::after, ::before { box-sizing: border-box; } body { margin: 0; font-family: -apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,'Helvetica Neue',Arial,'Noto Sans',sans-serif,'Apple Color Emoji','Segoe UI Emoji','Segoe UI Symbol','Noto Color Emoji'; line-height: 1.5; } .container { width: 100%; padding-right: 15px; padding-left: 15px;  margin: auto;  max-width: 725px; } .register { margin-top: 1%;  border-radius: 10px; background: #fff; }  p,table, caption, td, tr, th { margin:0; padding:0; font-weight:normal; }  /* ---- Paragraphs ---- */  p { margin-bottom:15px; }  /* ---- Table ---- */  h1.headreview{ font-size:40px; margin-bottom:10px; } .separator{ background: #b4beda;; margin-left: -15px; margin-right: -15px; height: 100px;  }  table { border-collapse:collapse; margin-bottom:15px;  }  caption {  font-size:15px; padding-bottom:10px; }  table td, table th {  border-width:0 1px 1px 0; border: 1px solid black; border-collapse: collapse; } th, td { padding: 1px; text-align: left; } } thead th[colspan], thead th[rowspan] { background:#66a9bd; }  tbody th, tfoot th { border-bottom:1px solid #000000; text-align: center; font-size: 14px; }  tbody td, tfoot td { border-bottom: 1px solid #000000; font-size: 12px; }  tfoot th { }  tfoot td { font-weight:bold; }  tbody tr.odd td {  background:#bcd9e1; }   .img { vertical-align: middle; border-style: none; height: 30px; width: 30px; margin-right:7px; } .header-up{ display: inline-block; width: 100%; padding-top: 10px; } .header-down{ display: inline-block; width: 100%; font-weight: bold; } .left-side{  float: left; width: 24%; display: flex; font-weight: bold; } .right-side{ float: left; border-left:2px solid #000000; padding: 0px 0px 0px 13px; font-weight: bold; }  .left-logo {  float:left;  width:50%;  }  .right-logo {   float:right; width:50%;   }   .right-logo span{ float: right; font-weight: bold;  } .eve-header{ font-size: 25px; padding: 0px; margin: 0px; } .footer{ height:50px; background: #000000; margin-right: -15px; margin-left: -15px; color: #fff; } .foot-content { position: absolute; color: #fff; margin-top: 10px; padding-left: 10px; font-size: 13px;  }  </style>");
                builder.Append("</head>");
                builder.Append("<body>");
                builder.Append("<div class='container register'>");
                builder.Append("<div class='row'>");
                builder.Append("<div class='header-up'>");
                builder.Append("<div class='left-logo'>");
                builder.Append("<div class='img'><svg aria-hidden='true' focusable='false' data-prefix='fab' data-icon='microsoft' role='img' xmlns='http://www.w3.org/2000/svg' viewBox='0 0 448 512' class='svg-inline--fa fa-microsoft fa-w-14 fa-3x'><path fill='currentColor' d='M0 32h214.6v214.6H0V32zm233.4 0H448v214.6H233.4V32zM0 265.4h214.6V480H0V265.4zm233.4 0H448V480H233.4V265.4z' class=''></path></svg></div>");
                builder.Append("</div>");
                builder.Append("<div class='right-logo'>");
                builder.Append(" <span>CSEO Assessment Services : No Features Created in ADO</span>");
                builder.Append("</div></div>");
                builder.Append("<h2>Contact us</h2>");
                builder.Append("<span><a href='mailto:CSEOA11yteam@microsoft.com?Subject=CSEO%20Accessibiltiy%20Services'>CSEO Accessibility Team</a></span>");
                builder.Append("<span>CSEO Studio Accessibility</span> ");
                builder.Append("</div>");
                builder.Append("<div class='footer' style='background:#A49F9E;'><span class='foot-content'>Visit our <a href='https://microsoft.sharepoint.com/teams/meconnection/SitePages/Accessibility.aspx'>Accessibility Hub website</a> to learn more.</span></div>");
                builder.Append(" </div>");
                builder.Append(" </div>");
                builder.Append("</body>");
                builder.Append("</html>");
                oMsg.Subject = "CSEO Accessbility Services - No Features Created in ADO : " + System.DateTime.Now.ToShortDateString() + " " + System.DateTime.Now.ToShortTimeString();
            }
            else
            {
                builder.Append("<!DOCTYPE html>");
                builder.Append("<html lang='en'>");
                builder.Append("<head>");
                builder.Append("<meta charset='utf-8'>");
                builder.Append("<meta http-equiv='X-UA-Compatible' content='IE=edge'>");
                builder.Append("<meta name='viewport' content='width=device-width, initial-scale=1'>");
                builder.Append("<title>Tables</title>");
                builder.Append("<style type='text/css'> * { -webkit-box-sizing: border-box; -moz-box-sizing: border-box; box-sizing: border-box; } *, ::after, ::before { box-sizing: border-box; } body { margin: 0; font-family: -apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,'Helvetica Neue',Arial,'Noto Sans',sans-serif,'Apple Color Emoji','Segoe UI Emoji','Segoe UI Symbol','Noto Color Emoji'; line-height: 1.5; } .container { width: 100%; padding-right: 15px; padding-left: 15px;  margin: auto;  max-width: 725px; } .register { margin-top: 1%;  border-radius: 10px; background: #fff; }  p,table, caption, td, tr, th { margin:0; padding:0; font-weight:normal; }  /* ---- Paragraphs ---- */  p { margin-bottom:15px; }  /* ---- Table ---- */  h1.headreview{ font-size:40px; margin-bottom:10px; } .separator{ background: #b4beda;; margin-left: -15px; margin-right: -15px; height: 100px;  }  table { border-collapse:collapse; margin-bottom:15px;  }  caption {  font-size:15px; padding-bottom:10px; }  table td, table th {  border-width:0 1px 1px 0; border: 1px solid black; border-collapse: collapse; } th, td { padding: 1px; text-align: left; } } thead th[colspan], thead th[rowspan] { background:#66a9bd; }  tbody th, tfoot th { border-bottom:1px solid #000000; text-align: center; font-size: 14px; }  tbody td, tfoot td { border-bottom: 1px solid #000000; font-size: 12px; }  tfoot th { }  tfoot td { font-weight:bold; }  tbody tr.odd td {  background:#bcd9e1; }   .img { vertical-align: middle; border-style: none; height: 30px; width: 30px; margin-right:7px; } .header-up{ display: inline-block; width: 100%; padding-top: 10px; } .header-down{ display: inline-block; width: 100%; font-weight: bold; } .left-side{  float: left; width: 24%; display: flex; font-weight: bold; } .right-side{ float: left; border-left:2px solid #000000; padding: 0px 0px 0px 13px; font-weight: bold; }  .left-logo {  float:left;  width:50%;  }  .right-logo {   float:right; width:50%;   }   .right-logo span{ float: right; font-weight: bold;  } .eve-header{ font-size: 25px; padding: 0px; margin: 0px; } .footer{ height:50px; background: #000000; margin-right: -15px; margin-left: -15px; color: #fff; } .foot-content { position: absolute; color: #fff; margin-top: 10px; padding-left: 10px; font-size: 13px;  }  </style>");
                builder.Append("</head>");
                builder.Append("<body>");
                builder.Append("<div class='container register'>");
                builder.Append("<div class='row'>");
                builder.Append("<div class='header-up'>");
                builder.Append("<div class='left-logo'>");
                builder.Append("<div class='img'><svg aria-hidden='true' focusable='false' data-prefix='fab' data-icon='microsoft' role='img' xmlns='http://www.w3.org/2000/svg' viewBox='0 0 448 512' class='svg-inline--fa fa-microsoft fa-w-14 fa-3x'><path fill='currentColor' d='M0 32h214.6v214.6H0V32zm233.4 0H448v214.6H233.4V32zM0 265.4h214.6V480H0V265.4zm233.4 0H448V480H233.4V265.4z' class=''></path></svg></div>");
                builder.Append("</div>");
                builder.Append("<div class='right-logo'>");
                //builder.Append(" <span>CSEO Assessment Services</span><br>");
                builder.Append("</div></div>");
                builder.Append("<br><br>");
                builder.Append(" <div class='header-down'>");
                builder.Append("<div class='left-side'>");
                builder.Append("<div class='img'>");
                builder.Append("<svg xmlns='http://www.w3.org/2000/svg' id='Layer_3' enable-background='new 0 0 64 64'viewBox='0 0 64 64'><g><path d='m20.042 29.373c5.944-5.944 9.409-13.701 9.893-22.023l26.716 26.716c-8.323.484-16.08 3.95-22.023 9.894l-2.628 2.626-14.586-14.586zm8.958 20.213-14.586-14.586 1.586-1.586 14.586 14.586zm-12 11-13.586-13.586 1.586-1.586 13.586 13.586zm-8.586-14.586 1.586-1.586 9.586 9.586-1.586 1.586zm16.758 7.414c-.372.373-.888.586-1.414.586h-.516c-.526 0-1.042-.213-1.414-.586l-11.242-11.242c-.378-.379-.586-.881-.586-1.415v-.515c0-.534.208-1.036.586-1.415l2.414-2.413 14.586 14.586zm13.535 3.293 3.293-3.293 1.586 1.586-5.586 5.586-8.586-8.586 3.586-3.586 5.586 5.586-1.293 1.293zm-1.707-7.121-2.586-2.586.856-.856 3.878 1.293zm22-16-28.586-28.586 1.586-1.586 28.586 28.586z'/><path d='m51.169 13h11.662v2h-11.662z' transform='matrix(.858 -.515 .515 .858 .92 31.321)'/><path d='m44.169 6h11.662v2h-11.662z' transform='matrix(.515 -.857 .857 .515 18.266 46.268)'/></g></svg>");
                builder.Append("</div>");
                builder.Append("</div>");
                builder.Append("</div>");
                builder.Append("Hi All,<br/><br/>");
                builder.Append("<p>The <b>CSEO ADO Workitems creater automated tool</b>, creates a new features for the newly created AIRT ID (record) if it doesn’t exist already. The feature will be a child to the below Scenarios : " +
                    "<ul><li><p><a href='https://microsoftit.visualstudio.com/OneITVSO/_workitems/edit/4889949'>SCENARIO 4889949</a>[Assessment] Plan and execute Assessments for CSEO applications. It will as well add a child link UserStory to the that feature.</p></li>" +
                    "<li><p><a href='https://microsoftit.visualstudio.com/OneITVSO/_workitems/edit/4889951'>SCENARIO 4889951</a>[Assessment] Assist Non-CSEO organizations with Assessment tasks. It will as well add a child link UserStory to the that feature.</p></li>" +
                    "<li><p><a href='https://microsoftit.visualstudio.com/OneITVSO/_workitems/edit/5324882'>SCENARIO 5324882</a>[Grade Review] Plan and Execute CSEO P3 Applications Grade Reviews. It will as well add a child link UserStory to the that feature</p></li>" +
                    "</ul>" +
                    "");
                builder.Append("<p>P.S – you need to manually define the iteration for the Feature/UserStory created above. The tag GetHealthy/StayHealthy will be added automatically by the tool.</p>");
                builder.Append("<h2>Newly Created ADO Features</h2>");
                builder.Append("<table style='width: 100%'>");
                builder.Append("<tr>");
                builder.Append("<th><b>RecID</b></th>");
                builder.Append("<th><b>Application Name</b></th>");
                builder.Append("<th><b>Component ID</b></th>");
                builder.Append("<th><b>Created Date</b></th>");
                builder.Append("<th><b>Group</b></th>");
                builder.Append("<th><b>Sub Group</b></th>");
                builder.Append("<th><b>Priority</b></th>");
                builder.Append("<th><b>Grade</b></th>");
                builder.Append("<th><b>ADO Feature Id</b></th>");
                builder.Append("<th><b>Comments</b></th>");
                builder.Append("<th><b>ADO UserStory Id</b></th>");
                builder.Append("<th><b>Comments</b></th>");
                builder.Append("</tr>");

                if (dataList != null)
                {
                    foreach (string dataString in dataList)
                    {
                        string[] dataArray = dataString.Split(',');
                        builder.Append("<tr>");
                        builder.Append("<td><a href='https://airt.azurewebsites.net/InventoryDetails/" + dataArray[0] + "'>" + dataArray[0] + "</a> </td>");
                        builder.Append("<td>" + dataArray[1] + "</td>");
                        builder.Append("<td>" + dataArray[2] + "</td>");
                        builder.Append("<td>" + dataArray[3] + "</td>");
                        builder.Append("<td>" + dataArray[4] + "</td>");
                        builder.Append("<td>" + dataArray[5] + "</td>");
                        builder.Append("<td>" + dataArray[6] + "</td>");
                        builder.Append("<td>" + dataArray[7] + "</td>");
                        builder.Append("<td><center><a href='https://microsoftit.visualstudio.com/OneITVSO/_workitems/edit/" + dataArray[8] + "'>" + dataArray[8] + "</a></center></td>");
                        builder.Append("<td>" + dataArray[9] + "</td>");
                        if (dataArray[10].Contains(":"))
                        {
                            builder.Append("<td><center>");
                            String[] tagArray = dataArray[10].Split(':');
                            foreach (string tagString in tagArray)
                            {
                                builder.Append("<a href='https://microsoftit.visualstudio.com/OneITVSO/_workitems/edit/" + tagString + "'>" + tagString + "</a><br>");
                            }
                            builder.Append("</center></td>");
                        }
                        else
                        {
                            builder.Append("<td><center><a href='https://microsoftit.visualstudio.com/OneITVSO/_workitems/edit/" + dataArray[10] + "'>" + dataArray[10] + "</a></center></td>");
                        }

                        builder.Append("<td>" + dataArray[11] + "</td>");
                        builder.Append("</tr>");
                    }
                }
                builder.Append("</table>");
                builder.Append("<br><p>For the above UserStory, you need to create two tasks.(E.g. sample tasks below). One for onboarding and other for assessment.</p>");
                //builder.Append("<h2>Next steps and checklist</h2>");
                builder.Append("<ul>");
                builder.Append("<li>Task1 => [Activity:Onboarding][2002][RecID_1491][CFE] Succession Planning</li>");
                builder.Append("<li>Task2 => [Activity:Assessment][2002][RecID_1491][CFE] Succession Planning</li>");
                builder.Append("</ul>");
                builder.Append("<p>Please add the tags as well define the iterations for the tasks.</p>");
                builder.Append("<div style='display:grid;background:#cac4c4;margin-left: -15px;margin-right:-15px;padding-left: 15px;padding-right: 15px;'>");
                builder.Append("<h2>Contact us</h2>");
                builder.Append("<span><a href='#'>CSEO Accessibility Team</a></span>");
                //builder.Append("<span>CSEO Studio Accessibility</span> ");
                builder.Append("</div>");
                builder.Append("<div class='footer' style='background:#A49F9E;'><span class='foot-content'>Visit our <a href='#'>Accessibility Hub website</a> to learn more.</span></div>");
                builder.Append(" </div>");
                builder.Append(" </div>");
                builder.Append("</body>");
                builder.Append("</html>");
                oMsg.Subject = "CSEO Accessbility Services - Newly Created Features in ADO : " + System.DateTime.Now.ToShortDateString() + " " + System.DateTime.Now.ToShortTimeString();
            }
            string HtmlFile = builder.ToString();

            /* Live */
            oMsg.To = ConfigurationManager.AppSettings["DftSendMailTo"];
            oMsg.CC = ConfigurationManager.AppSettings["DftSendMailCC"];

            oMsg.HTMLBody = HtmlFile;
            Console.WriteLine("ADO Features are created and sent mail to respective folks");
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
            dbConn = @"Data Source = ";
            cmd.CommandText = QueryName;
            Console.WriteLine("Executes {0}", QueryName);
            SqlConnection sqlConnection1 = new SqlConnection(dbConn);
            cmd.Connection = sqlConnection1;
            sqlConnection1.Open();
            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            sda.Fill(dt);
            dt.TableName = tableName;
            sqlConnection1.Close();
            Console.WriteLine("In {0} table identified {1} rows and {2} colums.......", dt.TableName, dt.Rows.Count, dt.Columns.Count);
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
            public const string ParentRefStr = "System.LinkTypes.Hierarchy-Reverse";
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
