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


        /*Prod*/
        //static readonly string UserPAT = "";
        //static readonly string TFUrl = "https://microsoftit.visualstudio.com/";
        //static readonly string teamProject = "OneITVSO";
        //public static string AreaPath = "OneITVSO\\Shared Experiences\\Studio\\Accessibility\\Accessibility PM";

        /*QA*/
        static readonly string UserPAT = "";
        static readonly string TFUrl = "https://smsgaccessibilityreviews.visualstudio.com";
        static readonly string teamProject = "CDSVSO";
        public static string AreaPath = "CDSVSO";

        static readonly string UserAccount = "";
        static readonly string UserPassword = "";

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
            //sendMail(reportType_2, null);
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
                    int tag = 0;
                    int newRecid = (int)row["RecId"];
                    if (!DBNull.Value.Equals(row["TagId"]))
                        tag = (int)row["TagId"];
                    if (tag != 2)
                    {
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
                }
                if (dataList.Count == 0)
                {
                    sendMail(reportType_1, null);
                }
                else
                {
                    sendMail(reportType_2, dataList);
                }

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

        /// <summary>
        /// Simple operations woth query
        /// </summary>
        /// <param name="project">Team Project Name</param>
        /// <param name="queryRootPath">Root Folder for Query</param>
        /// <param name="queryName">Query Name</param>
        static void OperateWithQuery(string project, string queryRootPath, string queryName)
        {
            //Get new and active tasks
            string customWiql = @"SELECT [System.Id], [System.Title], [System.State] FROM WorkItems WHERE [System.TeamProject] = @project" +
                @" and [System.WorkItemType] = 'Task' and [System.State] <> 'Closed'";

            string customWiql_Features = @"SELECT
    [System.Id],
    [System.WorkItemType],
    [System.Title],
    [System.AssignedTo],
    [System.State],
    [System.Tags]
FROM workitems
WHERE
    [System.TeamProject] = @project
    AND [System.WorkItemType] = 'Feature'
    AND (
        [System.Title] CONTAINS WORDS 'RecID_830'
        OR [System.Title] CONTAINS WORDS 'RecID_916'
        OR [System.Title] CONTAINS WORDS 'RecID_4035'
    )
    AND [System.State] IN ('New', 'Active')";

            //Get new tasks only
            string updatedWiql = @"SELECT [System.Id], [System.Title] FROM WorkItems WHERE [System.TeamProject] = @project" +
                @" and [System.WorkItemType] = 'Task' and [System.State] == 'New'";

            Console.WriteLine("Create New Query");
            AddQuery(project, queryRootPath, queryName, customWiql);
            RunStoredQuery(project, queryRootPath + "/" + queryName);

            Console.WriteLine("\nUpdate Query");
            EditQuery(project, queryRootPath + "/" + queryName, updatedWiql);
            RunStoredQuery(project, queryRootPath + "/" + queryName);

            Console.WriteLine("\nDelete Query");
            RemoveQuery(project, queryRootPath + "/" + queryName);
        }

        /// <summary>
        /// Run stored query on tfs/vsts
        /// </summary>
        /// <param name="project">Team Project Name</param>
        /// <param name="queryPath">Path to Query</param>
        static void RunStoredQuery(string project, string queryPath)
        {
            QueryHierarchyItem query = WitClient.GetQueryAsync(project, queryPath, QueryExpand.Wiql).Result;

            string wiqlStr = query.Wiql;

            GetQueryResults(wiqlStr, project);
        }

        /// <summary>
        /// Run query and show result
        /// </summary>
        /// <param name="wiqlStr">Wiql String</param>
        static void GetQueryResults(string wiqlStr, string teamProject)
        {
            WorkItemQueryResult result = RunQueryByWiql(wiqlStr, teamProject);

            if (result != null)
            {
                if (result.WorkItems != null) // this is Flat List 
                    foreach (var wiRef in result.WorkItems)
                    {
                        var wi = GetWorkItem(wiRef.Id);
                        Console.WriteLine(String.Format("{0} - {1}", wi.Id, wi.Fields["System.Title"].ToString()));
                    }
                else if (result.WorkItemRelations != null) // this is Tree of Work Items or Work Items and Direct Links
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

            /// <summary>
            /// Update Existing Query
            /// </summary>
            /// <param name="project"></param>
            /// <param name="queryPath"></param>
            /// <param name="newWiqlStr"></param>
            static void EditQuery(string project, string queryPath, string newWiqlStr)
        {
            QueryHierarchyItem query = WitClient.GetQueryAsync(project, queryPath).Result;
            query.Wiql = newWiqlStr;

            query = WitClient.UpdateQueryAsync(query, project, queryPath).Result;
        }

        /// <summary>
        /// Remove Existing Query or Folder
        /// </summary>
        /// <param name="project"></param>
        /// <param name="queryPath"></param>
        static void RemoveQuery(string project, string queryPath)
        {
            WitClient.DeleteQueryAsync(project, queryPath);
        }

        /// <summary>
        /// Create new query
        /// </summary>
        /// <param name="project"></param>
        /// <param name="queryPath"></param>
        /// <param name="QueryName"></param>
        /// <param name="wiqlStr"></param>
        static void AddQuery(string project, string queryPath, string QueryName, string wiqlStr)
        {
            QueryHierarchyItem query = new QueryHierarchyItem();
            query.QueryType = QueryType.Flat;
            query.Name = QueryName;
            query.Wiql = wiqlStr;

            query = WitClient.CreateQueryAsync(query, project, queryPath).Result;
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
                builder.Append("<html xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:w=\"urn:schemas-microsoft-com:office:word\" xmlns:m=\"http://schemas.microsoft.com/office/2004/12/omml\" xmlns=\"http://www.w3.org/TR/REC-html40\"><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=Windows-1252\"><meta name=\"Generator\" content=\"Microsoft Word 15 (filtered medium)\">");
                builder.Append("<!--[if !mso]><style>v:*{behavior:url(#default#VML)}o:*{behavior:url(#default#VML)}w:*{behavior:url(#default#VML)}.shape{behavior:url(#default#VML)}</style><![endif]-->");
                builder.Append("<style><!--");
                builder.Append("@font-face{font-family:\"Cambria Math\";panose-1:2 4 5 3 5 4 6 3 2 4}@font-face{font-family:Calibri;panose-1:2 15 5 2 2 2 4 3 2 4}@font-face{font-family:\"Calibri Light\";panose-1:2 15 3 2 2 2 4 3 2 4}@font-face{font-family:\"Segoe UI\";panose-1:2 11 5 2 4 2 4 2 2 3}@font-face{font-family:\"Segoe UI Semibold\";panose-1:2 11 7 2 4 2 4 2 2 3}div.MsoNormal,li.MsoNormal,p.MsoNormal{margin:0;margin-bottom:.0001pt;font-size:11pt;font-family:Calibri,sans-serif}h1{mso-style-priority:9;mso-style-link:\"Heading 1 Char\";margin-top:12pt;margin-right:0;margin-bottom:0;margin-left:0;margin-bottom:.0001pt;page-break-after:avoid;font-size:16pt;font-family:\"Calibri Light\",sans-serif;color:#2f5496;font-weight:400}h2{mso-style-priority:9;mso-style-link:\"Heading 2 Char\";margin-top:2pt;margin-right:0;margin-bottom:0;margin-left:0;margin-bottom:.0001pt;page-break-after:avoid;font-size:13pt;font-family:\"Calibri Light\",sans-serif;color:#2f5496;font-weight:400}div.MsoTitle,li.MsoTitle,p.MsoTitle{mso-style-name:\"Title\\,Heading 1 Bold\";mso-style-priority:10;mso-style-link:\"Title Char\\,Heading 1 Bold Char\";margin-top:0;margin-right:0;margin-bottom:6pt;margin-left:0;line-height:90%;font-size:11pt;font-family:\"Segoe UI Semibold\",sans-serif;color:#282828;letter-spacing:-1pt}a:link,span.MsoHyperlink{mso-style-priority:99;color:#0563c1;text-decoration:underline}span.MsoIntenseEmphasis{mso-style-priority:21;font-family:\"Times New Roman\",serif;color:#5b9bd5;font-style:italic}span.Heading1Char{mso-style-name:\"Heading 1 Char\";mso-style-priority:9;mso-style-link:\"Heading 1\";font-family:\"Calibri Light\",sans-serif;color:#2f5496}span.Heading2Char{mso-style-name:\"Heading 2 Char\";mso-style-priority:9;mso-style-link:\"Heading 2\";font-family:\"Calibri Light\",sans-serif;color:#2f5496}span.TitleChar{mso-style-name:\"Title Char\\,Heading 1 Bold Char\";mso-style-priority:10;mso-style-link:\"Title\\,Heading 1 Bold\";font-family:\"Segoe UI Semibold\",sans-serif;color:#282828;letter-spacing:-1pt}span.LogoChar{mso-style-name:\"Logo Char\";mso-style-link:Logo;font-family:\"Segoe UI Semibold\",sans-serif;color:#282828;letter-spacing:-.3pt}div.Logo,li.Logo,p.Logo{mso-style-name:Logo;mso-style-link:\"Logo Char\";margin-top:0;margin-right:0;margin-bottom:28pt;margin-left:0;font-size:11pt;font-family:\"Segoe UI Semibold\",sans-serif;color:#282828;letter-spacing:-.3pt}.MsoChpDefault{mso-style-type:export-only;font-size:10pt}@page WordSection1{size:8.5in 11in;margin:1in 1in 1in 1in}div.WordSection1{page:WordSection1}");
                builder.Append("-->");
                builder.Append(".TFtable{width:100%;border-collapse:collapse}.TFtable td,th{padding:7px;border:#4e95f4 1px solid}.th{background-color:#c1e4ec}.TFtable tr{background:#b8d1f3}.TFtable tr:nth-child(odd){background:#b8d1f3}.TFtable tr:nth-child(even){background:#dae5f4}");
                builder.Append("</style>");
                builder.Append("<!--[if gte mso 9]> <xml> <o:shapedefaults v:ext=\"edit\" spidmax=\"1027\" /> </xml> <![endif]--> <!--[if gte mso 9]> <xml> <o:shapelayout v:ext=\"edit\"> <o:idmap v:ext=\"edit\" data=\"1\" /> </o:shapelayout> </xml> <![endif]-->");
                builder.Append("</head><body lang=\"EN-US\" link=\"#0563C1\" vlink=\"#954F72\"><div class=\"WordSection1\"><p class=\"MsoNormal\"> <o:p>&nbsp;</o:p></p><div align=\"center\"><table class=\"MsoNormalTable\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\" width=\"837\" style=\"width:627.6pt;border-collapse:collapse\"><tr style=\"height:83.7pt\"><td width=\"837\" nowrap=\"\" valign=\"bottom\" style=\"width:627.6pt;background:#c1e4ec;padding:0in 0in 0in .3in;height:83.7pt\"><table class=\"MsoNormalTable\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\" style=\"border-collapse:collapse\"><tr><td width=\"499\" valign=\"bottom\" style=\"width:374.05pt;padding:0in 0in .15in 0in\"><p class=\"MsoNormal\"> <a href=\"https://microsoftit.visualstudio.com/DefaultCollection/OneITVSO/\"> <img border=\"0\" style=\"width:-1.6875in;height:.4791in\" width=\"93\" height=\"45\" id=\"_x0000_i1026\" src=\"https://www.freepnglogos.com/uploads/microsoft-logo-png-transparent-20.png\"> <o:p></o:p> </a></p></td><td width=\"499\" valign=\"bottom\" style=\"width:374.05pt;padding:0in 0in .15in 0in\">");
                builder.Append("<p class=\"MsoNormal\"><b>CSEO Assessment Services: " + DateTime.Now.ToString("dddd , MMM dd yyyy,hh:mm:ss"));
                builder.Append("</b></p>");
                builder.Append("</td></tr><tr></tr></table></td></tr><tr style=\"height:.1in\"><td width=\"837\" valign=\"top\" style=\"width:627.6pt;padding:0in 0in 0in 0in;height:.1in\"><p class=\"MsoNormal\"> <span style=\"font-size:1.0pt;font-family:&quot;Segoe UI&quot;,sans-serif;color:#262626;letter-spacing:-.2pt\"> <o:p>&nbsp;</o:p> </span></p></td></tr><tr style=\"height:1.8pt\"><td width=\"837\" valign=\"top\" style=\"width:627.6pt;background:#F2F2F2;padding:.15in .3in 0in .3in;height:1.8pt\"></td></tr><tr style=\"height:64.35pt\"><td width=\"837\" valign=\"top\" style=\"width:627.6pt;background:#F2F2F2;padding:0in .3in 0in .3in;height:64.35pt\"><h1>");
                //builder.Append("<!--[if gte vml 1]> <v:shapetype id=\"_x0000_t75\" coordsize=\"21600,21600\" o:spt=\"75\" o:preferrelative=\"t\" path=\"m@4@5l@4@11@9@11@9@5xe\" filled=\"f\" stroked=\"f\"> <v:stroke joinstyle=\"miter\" /> <v:formulas> <v:f eqn=\"if lineDrawn pixelLineWidth 0\" /> <v:f eqn=\"sum @0 1 0\" /> <v:f eqn=\"sum 0 0 @1\" /> <v:f eqn=\"prod @2 1 2\" /> <v:f eqn=\"prod @3 21600 pixelWidth\" /> <v:f eqn=\"prod @3 21600 pixelHeight\" /> <v:f eqn=\"sum @0 0 1\" /> <v:f eqn=\"prod @6 1 2\" /> <v:f eqn=\"prod @7 21600 pixelWidth\" /> <v:f eqn=\"sum @8 21600 0\" /> <v:f eqn=\"prod @7 21600 pixelHeight\" /> <v:f eqn=\"sum @10 21600 0\" /> </v:formulas> <v:path o:extrusionok=\"f\" gradientshapeok=\"t\" o:connecttype=\"rect\" /> <o:lock v:ext=\"edit\" aspectratio=\"t\" /> </v:shapetype> <v:shape id=\"Picture_x0020_1\" o:spid=\"_x0000_s1026\" type=\"#_x0000_t75\" style='position:absolute;margin-left:-30.6pt;margin-top:0;width:125.25pt;height:102.7pt;z-index:251659264;visibility:visible;mso-wrap-style:square;mso-width-percent:0;mso-height-percent:0;mso-wrap-distance-left:9pt;mso-wrap-distance-top:0;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:0;mso-position-horizontal:absolute;mso-position-horizontal-relative:text;mso-position-vertical:absolute;mso-position-vertical-relative:text;mso-width-percent:0;mso-height-percent:0;mso-width-relative:page;mso-height-relative:page'> <v:imagedata src=\"cid:image002.jpg@01D60205.B6519420\" o:title=\"\" /> <w:wrap type=\"square\"/> </v:shape> <![endif]-->");
                //builder.Append("<![if !vml]> <!img style=\"width:1.7395in;height:1.427in;float: left;\" src=\"https://pngimage.net/wp-content/uploads/2018/05/alert-logo-png-1.png\" align=\"left\" hspace=\"12\" v:shapes=\"Picture_x0020_1\">");
                builder.Append("<p class=\"MsoNormal\" style=\"display: flex !important; \">");
                builder.Append("<img border=\"0\" width=\"38\" height=\"34\" id=\"_x0000_i1026\" src=\"https://cdn.pixabay.com/photo/2013/07/13/09/36/megaphone-155780_960_720.png\"> <o:p></o:p><h2 style=\"color:#17067a;\" ><b>Alert (No Action Required: CSEO Accessability Services- No Features and UserStories are Created</b> <o:p></o:p></h2>");
                builder.Append("</p> <br/><p class=\"MsoNormal\" style=\"vertical-align:middle\"> <span style=\"color:black\">The CSEO ADO Workitems creater tool, Creates a new features for the newly created AIRT ID (record) if it doesn’t exist already. The feature will be a child to the below Scenarios:</span></p> <o:p></o:p><p class=\"MsoNormal\" style=\"vertical-align:middle\"> <span style=\"color:black\">1. <a href=\"https://microsoftit.visualstudio.com/OneITVSO/_workitems/edit/4889949\">SCENARIO 4889949: </a>[Assessment] Plan and Assessments for CSEO applications. it will as well add a child link UserStory to that feature.</span>");
                //builder.Append("<br/><span style=\"color:black\">2. <a href=\"https://microsoftit.visualstudio.com/OneITVSO/_workitems/edit/4889951\">SCENARIO 4889951: </a>[Assessment] Assist Non-CSEO organizations with Assessment tasks. It will as well add a child link UserStory to the that feature.</span>");
                builder.Append("<br/><span style=\"color:black\">2. <a href=\"https://microsoftit.visualstudio.com/OneITVSO/_workitems/edit/5324882\">SCENARIO 5324882: </a>[Grade Review] Plan and Execute CSEO P3 Applications Grade Reviews. It will as well add a child link UserStory to the that feature</span>" +
                    "</p> <br/>" +
                    "<p class=\"MsoNormal\" style=\"color:black;\"> P.S- You need to manually define the iteration for the feature/UserStory created above. The tag GetHealthy/StayHealthy will be added automatically by the tool. <o:p>&nbsp;</o:p></p>" +
                    "<p class=\"MsoNormal\"> <o:p>&nbsp;</o:p></p>");
                builder.Append("<p class=\"MsoNormal\">" +
                    "<span class=\"Heading2Char\">" +
                    "<span style=\"font-size:13.0pt\"><b>Newly Created ADO Features and UserStories:0</b></span> </span> <span style=\"font-size:10.0pt;font-family:&quot;Segoe UI Semibold&quot;,sans-serif\"> <o:p></o:p> </span></p><p class=\"MsoNormal\"> <span style=\"color:black\">");
                builder.Append("<br/><p class=\"MsoNormal\"> <span class=\"Heading2Char\"> <span style=\"font-size:13.0pt\"><b>Overview:</b></span> </span> <span style=\"font-size:10.0pt;font-family:&quot;Segoe UI Semibold&quot;,sans-serif\"> <o:p></o:p> </span></p>");
                builder.Append("<p class=\"MsoNormal\"> " +
                    "<span style=\"color:black\">" +
                    "<table class=\"TFtable\">" +
                    "<colgroup span=\"2\"></colgroup>" +
                    "<colgroup span=\"2\"></colgroup>" +
                    "<colgroup span=\"2\"></colgroup>" +
                    "<colgroup span=\"2\"></colgroup>" +
                    "<tr>" +
                    "<th colspan=\"2\" scope=\"colgroup\">Group/SubGroup</th>" +
                    "<th colspan=\"3\" scope=\"colgroup\">Priority In AIRT</th>" +
                    "<th colspan=\"3\" scope=\"colgroup\">Active / New Features in ADO</th>" +
                    "<th colspan=\"3\" scope=\"colgroup\">Active / New User Story in ADO</th>" +
                    "</tr>" +
                    "<tr>" +
                    "<th>Group</th>" +
                    "<th>Sub Group</th>" +
                    "<th>P1</th>" +
                    "<th>P2</th>" +
                    "<th>P3</th>" +
                    "<th>P1</th>" +
                    "<th>P2</th>" +
                    "<th>P3</th>" +
                    "<th>P1</th>" +
                    "<th>P2</th>" +
                    "<th>P3</th>" +
                    "</tr>");
                builder.Append("<tr>");
                builder.Append("<td>Text</td>");
                builder.Append("<td>Text</td>");
                builder.Append("<td>0</td>");
                builder.Append("<td>0</td>");
                builder.Append("<td>0</td>");
                builder.Append("<td>0</td>");
                builder.Append("<td>0</td>");
                builder.Append("<td>0</td>");
                builder.Append("<td>0</td>");
                builder.Append("<td>0</td>");
                builder.Append("<td>0</td>");
                builder.Append("</tr>");
                //builder.Append("<tr style=\"height:64.35pt\"><td width=\"837\" valign=\"top\" style=\"width:627.6pt;background:#F2F2F2;padding:0in .3in 0in .3in;height:64.35pt\"> <br/><p class=\"MsoNormal\"> <span class=\"MsoIntenseEmphasis\"><b>Note: </b>CSEO Accessability Services ADO features and User story creator tool will create features and user storys only CSEO applications P1, P2 and P3 applications where TAGs doesn’t contains Non-Inventory.</span> <span class=\"MsoIntenseEmphasis\"> <span style=\"font-family:&quot;Calibri&quot;,sans-serif\"> <o:p></o:p> </span> </span></p></td></tr><tr style=\"height:64.35pt\"><td width=\"837\" valign=\"top\" style=\"width:627.6pt;background:#F2F2F2;padding:0in .3in 0in .3in;height:64.35pt\"><p class=\"MsoNormal\"> <o:p>&nbsp;</o:p></p></td></tr><tr style=\"height:41.85pt\"><td width=\"837\" valign=\"top\" style=\"width:627.6pt;background:#c1e4ec;padding:0in 0in 0in .3in;height:41.85pt\"><table class=\"MsoNormalTable\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\" style=\"border-collapse:collapse\"><tr style=\"height:27.9pt\"><td width=\"499\" valign=\"top\" style=\"width:374.25pt;padding:.15in .2in .15in 0in;height:27.9pt\"><p class=\"MsoNormal\"> <span style=\"font-size:8.0pt;color:#262626\"><b>Connect <a href=\"mailto:CSEOA11yteam@microsoft.com\">CSEO Accessibility Team</a></b></span></p><p class=\"MsoNormal\"> <span style=\"font-size:8.0pt;color:#262626\">Please visit our </span> <a href=\"https://microsoft.sharepoint.com/teams/meconnection/SitePages/Accessibility-Tags.aspx\"> <span style=\"font-size:8.0pt\">Accessbility Hub Website</span> </a> <span style=\"font-size:8.0pt\">to learn more. Microsoft <span style=\"color:#262626\">C</span>onfidential. </span> <span style=\"font-size:8.0pt;color:#282828;mso-fareast-language:IT\"> <o:p></o:p> </span></p></td><td width=\"144\" nowrap=\"\" style=\"width:107.75pt;border:solid windowtext 1.0pt;border-left:none;background:#282828;padding:0in 0in 0in 0in;height:27.9pt\"><p class=\"MsoNormal\" align=\"center\" style=\"text-align:center\"> <span style=\"color:white\"> <img border=\"0\" width=\"96\" height=\"42\" style=\"width:1.0in;height:.4375in\" id=\"_x0000_i1025\" src=\"https://www.freepnglogos.com/uploads/microsoft-logo-png-transparent-20.png\"> </span> <span style=\"font-size:9.0pt;color:white\"> <o:p></o:p> </span></p></td></tr></table></td></tr></table></div><p class=\"MsoNormal\"> <o:p>&nbsp;</o:p></p></div></body></html>");
                builder.Append("</table> </span></p></td></tr><tr style=\"height:64.35pt\"><td width=\"837\" valign=\"top\" style=\"width:627.6pt;background:#F2F2F2;padding:0in .3in 0in .3in;height:64.35pt\"> <br/>" +
                    "<p class=\"MsoNormal\"> <span class=\"MsoIntenseEmphasis\" style=\"color: black;\"><b>Note: </b>CSEO Accessability Services ADO features and User story creator tool will create features and user storys only CSEO applications P1, P2 and P3 applications where TAGs doesn’t contains Non-Inventory.</span> <span class=\"MsoIntenseEmphasis\"> <span style=\"font-family:&quot;Calibri&quot;,sans-serif\"> <o:p></o:p> </span> </span></p></td></tr><tr style=\"height:64.35pt\"><td width=\"837\" valign=\"top\" style=\"width:627.6pt;background:#F2F2F2;padding:0in .3in 0in .3in;height:64.35pt\"><p class=\"MsoNormal\"> <o:p>&nbsp;</o:p></p></td></tr><tr style=\"height:41.85pt\"><td width=\"837\" valign=\"top\" style=\"width:627.6pt;background:#c1e4ec;padding:0in 0in 0in .3in;height:41.85pt\"><table class=\"MsoNormalTable\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\" style=\"border-collapse:collapse\"><tr style=\"height:27.9pt\"><td width=\"499\" valign=\"top\" style=\"width:374.25pt;padding:.15in .2in .15in 0in;height:27.9pt\"><p class=\"MsoNormal\"> <span style=\"font-size:8.0pt;color:#262626\"><b>Connect <a href=\"mailto:CSEOA11yteam@microsoft.com\">CSEO Accessibility Team</a></b></span></p><p class=\"MsoNormal\"> <span style=\"font-size:8.0pt;color:#262626\">Please visit our </span> <a href=\"https://microsoft.sharepoint.com/teams/meconnection/SitePages/Accessibility-Tags.aspx\"> <span style=\"font-size:8.0pt\">Accessbility Hub Website</span> </a> <span style=\"font-size:8.0pt\">to learn more. Microsoft <span style=\"color:#262626\">C</span>onfidential. </span> <span style=\"font-size:8.0pt;color:#282828;mso-fareast-language:IT\"> <o:p></o:p> </span></p></td><td width=\"144\" nowrap=\"\" style=\"width:107.75pt;border:solid windowtext 1.0pt;border-left:none;background:#282828;padding:0in 0in 0in 0in;height:27.9pt\"><p class=\"MsoNormal\" align=\"center\" style=\"text-align:center\"> <span style=\"color:white\"> <img border=\"0\" width=\"96\" height=\"42\" style=\"width:1.0in;height:.4375in\" id=\"_x0000_i1025\" src=\"https://www.freepnglogos.com/uploads/microsoft-logo-png-transparent-20.png\"> </span> <span style=\"font-size:9.0pt;color:white\"> <o:p></o:p> </span></p></td></tr></table></td></tr></table></div><p class=\"MsoNormal\"> <o:p>&nbsp;</o:p></p></div></body></html>");
                oMsg.Subject = "CSEO Accessbility Services - No Features Created in ADO : " + System.DateTime.Now.ToShortDateString() + " " + System.DateTime.Now.ToShortTimeString();
            }
            else
            {
                builder.Append("<!DOCTYPE html>");
                builder.Append("<html xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:w=\"urn:schemas-microsoft-com:office:word\" xmlns:m=\"http://schemas.microsoft.com/office/2004/12/omml\" xmlns=\"http://www.w3.org/TR/REC-html40\"><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=Windows-1252\"><meta name=\"Generator\" content=\"Microsoft Word 15 (filtered medium)\">");
                builder.Append("<!--[if !mso]><style>v:*{behavior:url(#default#VML)}o:*{behavior:url(#default#VML)}w:*{behavior:url(#default#VML)}.shape{behavior:url(#default#VML)}</style><![endif]-->");
                builder.Append("<style><!--");
                builder.Append("@font-face{font-family:\"Cambria Math\";panose-1:2 4 5 3 5 4 6 3 2 4}@font-face{font-family:Calibri;panose-1:2 15 5 2 2 2 4 3 2 4}@font-face{font-family:\"Calibri Light\";panose-1:2 15 3 2 2 2 4 3 2 4}@font-face{font-family:\"Segoe UI\";panose-1:2 11 5 2 4 2 4 2 2 3}@font-face{font-family:\"Segoe UI Semibold\";panose-1:2 11 7 2 4 2 4 2 2 3}div.MsoNormal,li.MsoNormal,p.MsoNormal{margin:0;margin-bottom:.0001pt;font-size:11pt;font-family:Calibri,sans-serif}h1{mso-style-priority:9;mso-style-link:\"Heading 1 Char\";margin-top:12pt;margin-right:0;margin-bottom:0;margin-left:0;margin-bottom:.0001pt;page-break-after:avoid;font-size:16pt;font-family:\"Calibri Light\",sans-serif;color:#2f5496;font-weight:400}h2{mso-style-priority:9;mso-style-link:\"Heading 2 Char\";margin-top:2pt;margin-right:0;margin-bottom:0;margin-left:0;margin-bottom:.0001pt;page-break-after:avoid;font-size:13pt;font-family:\"Calibri Light\",sans-serif;color:#2f5496;font-weight:400}div.MsoTitle,li.MsoTitle,p.MsoTitle{mso-style-name:\"Title\\,Heading 1 Bold\";mso-style-priority:10;mso-style-link:\"Title Char\\,Heading 1 Bold Char\";margin-top:0;margin-right:0;margin-bottom:6pt;margin-left:0;line-height:90%;font-size:11pt;font-family:\"Segoe UI Semibold\",sans-serif;color:#282828;letter-spacing:-1pt}a:link,span.MsoHyperlink{mso-style-priority:99;color:#0563c1;text-decoration:underline}span.MsoIntenseEmphasis{mso-style-priority:21;font-family:\"Times New Roman\",serif;color:#5b9bd5;font-style:italic}span.Heading1Char{mso-style-name:\"Heading 1 Char\";mso-style-priority:9;mso-style-link:\"Heading 1\";font-family:\"Calibri Light\",sans-serif;color:#2f5496}span.Heading2Char{mso-style-name:\"Heading 2 Char\";mso-style-priority:9;mso-style-link:\"Heading 2\";font-family:\"Calibri Light\",sans-serif;color:#2f5496}span.TitleChar{mso-style-name:\"Title Char\\,Heading 1 Bold Char\";mso-style-priority:10;mso-style-link:\"Title\\,Heading 1 Bold\";font-family:\"Segoe UI Semibold\",sans-serif;color:#282828;letter-spacing:-1pt}span.LogoChar{mso-style-name:\"Logo Char\";mso-style-link:Logo;font-family:\"Segoe UI Semibold\",sans-serif;color:#282828;letter-spacing:-.3pt}div.Logo,li.Logo,p.Logo{mso-style-name:Logo;mso-style-link:\"Logo Char\";margin-top:0;margin-right:0;margin-bottom:28pt;margin-left:0;font-size:11pt;font-family:\"Segoe UI Semibold\",sans-serif;color:#282828;letter-spacing:-.3pt}.MsoChpDefault{mso-style-type:export-only;font-size:10pt}@page WordSection1{size:8.5in 11in;margin:1in 1in 1in 1in}div.WordSection1{page:WordSection1}");
                builder.Append("-->");
                builder.Append(".TFtable{width:100%;border-collapse:collapse}.TFtable td,th{padding:7px;border:#4e95f4 1px solid}.th{background-color:#c1e4ec}.TFtable tr{background:#b8d1f3}.TFtable tr:nth-child(odd){background:#b8d1f3}.TFtable tr:nth-child(even){background:#dae5f4}");
                builder.Append("</style>");
                builder.Append("<!--[if gte mso 9]> <xml> <o:shapedefaults v:ext=\"edit\" spidmax=\"1027\" /> </xml> <![endif]--> <!--[if gte mso 9]> <xml> <o:shapelayout v:ext=\"edit\"> <o:idmap v:ext=\"edit\" data=\"1\" /> </o:shapelayout> </xml> <![endif]-->");
                builder.Append("</head><body lang=\"EN-US\" link=\"#0563C1\" vlink=\"#954F72\"><div class=\"WordSection1\"><p class=\"MsoNormal\"> <o:p>&nbsp;</o:p></p><div align=\"center\"><table class=\"MsoNormalTable\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\" width=\"837\" style=\"width:627.6pt;border-collapse:collapse\"><tr style=\"height:83.7pt\"><td width=\"837\" nowrap=\"\" valign=\"bottom\" style=\"width:627.6pt;background:#c1e4ec;padding:0in 0in 0in .3in;height:83.7pt\"><table class=\"MsoNormalTable\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\" style=\"border-collapse:collapse\"><tr><td width=\"499\" valign=\"bottom\" style=\"width:374.05pt;padding:0in 0in .15in 0in\"><p class=\"MsoNormal\"> <a href=\"https://microsoftit.visualstudio.com/DefaultCollection/OneITVSO/\"> <img border=\"0\" style=\"width:-1.6875in;height:.4791in\" width=\"93\" height=\"45\" id=\"_x0000_i1026\" src=\"https://www.freepnglogos.com/uploads/microsoft-logo-png-transparent-20.png\"> <o:p></o:p> </a></p></td><td width=\"499\" valign=\"bottom\" style=\"width:374.05pt;padding:0in 0in .15in 0in\">");
                builder.Append("<p class=\"MsoNormal\"><b>CSEO Assessment Services: "+ DateTime.Now.ToString("dddd , MMM dd yyyy,hh:mm:ss"));
                builder.Append("</b></p>");
                builder.Append("</td></tr><tr></tr></table></td></tr><tr style=\"height:.1in\"><td width=\"837\" valign=\"top\" style=\"width:627.6pt;padding:0in 0in 0in 0in;height:.1in\"><p class=\"MsoNormal\"> <span style=\"font-size:1.0pt;font-family:&quot;Segoe UI&quot;,sans-serif;color:#262626;letter-spacing:-.2pt\"> <o:p>&nbsp;</o:p> </span></p></td></tr><tr style=\"height:1.8pt\"><td width=\"837\" valign=\"top\" style=\"width:627.6pt;background:#F2F2F2;padding:.15in .3in 0in .3in;height:1.8pt\"></td></tr><tr style=\"height:64.35pt\"><td width=\"837\" valign=\"top\" style=\"width:627.6pt;background:#F2F2F2;padding:0in .3in 0in .3in;height:64.35pt\"><h1>");
                //builder.Append("<!--[if gte vml 1]> <v:shapetype id=\"_x0000_t75\" coordsize=\"21600,21600\" o:spt=\"75\" o:preferrelative=\"t\" path=\"m@4@5l@4@11@9@11@9@5xe\" filled=\"f\" stroked=\"f\"> <v:stroke joinstyle=\"miter\" /> <v:formulas> <v:f eqn=\"if lineDrawn pixelLineWidth 0\" /> <v:f eqn=\"sum @0 1 0\" /> <v:f eqn=\"sum 0 0 @1\" /> <v:f eqn=\"prod @2 1 2\" /> <v:f eqn=\"prod @3 21600 pixelWidth\" /> <v:f eqn=\"prod @3 21600 pixelHeight\" /> <v:f eqn=\"sum @0 0 1\" /> <v:f eqn=\"prod @6 1 2\" /> <v:f eqn=\"prod @7 21600 pixelWidth\" /> <v:f eqn=\"sum @8 21600 0\" /> <v:f eqn=\"prod @7 21600 pixelHeight\" /> <v:f eqn=\"sum @10 21600 0\" /> </v:formulas> <v:path o:extrusionok=\"f\" gradientshapeok=\"t\" o:connecttype=\"rect\" /> <o:lock v:ext=\"edit\" aspectratio=\"t\" /> </v:shapetype> <v:shape id=\"Picture_x0020_1\" o:spid=\"_x0000_s1026\" type=\"#_x0000_t75\" style='position:absolute;margin-left:-30.6pt;margin-top:0;width:125.25pt;height:102.7pt;z-index:251659264;visibility:visible;mso-wrap-style:square;mso-width-percent:0;mso-height-percent:0;mso-wrap-distance-left:9pt;mso-wrap-distance-top:0;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:0;mso-position-horizontal:absolute;mso-position-horizontal-relative:text;mso-position-vertical:absolute;mso-position-vertical-relative:text;mso-width-percent:0;mso-height-percent:0;mso-width-relative:page;mso-height-relative:page'> <v:imagedata src=\"cid:image002.jpg@01D60205.B6519420\" o:title=\"\" /> <w:wrap type=\"square\"/> </v:shape> <![endif]-->");
                //builder.Append("<![if !vml]> <!img style=\"width:1.7395in;height:1.427in;float: left;\" src=\"https://pngimage.net/wp-content/uploads/2018/05/alert-logo-png-1.png\" align=\"left\" hspace=\"12\" v:shapes=\"Picture_x0020_1\">");
                builder.Append("<p class=\"MsoNormal\" style=\"display: flex !important; \">");
                builder.Append("<img border=\"0\" width=\"38\" height=\"34\" id=\"_x0000_i1026\" src=\"https://cdn.pixabay.com/photo/2013/07/13/09/36/megaphone-155780_960_720.png\"> <o:p></o:p><h2 style=\"color:#f02d0b;\" ><b>Alert (Action Required: CSEO Accessability Services- Newly Features and UserStories are Created</b> <o:p></o:p></h2>");
                builder.Append("</p> <br/><p class=\"MsoNormal\" style=\"vertical-align:middle\"> <span style=\"color:black\">The CSEO ADO Workitems creater tool, Creates a new features for the newly created AIRT ID (record) if it doesn’t exist already. The feature will be a child to the below Scenarios:</span></p> <o:p></o:p><p class=\"MsoNormal\" style=\"vertical-align:middle\"> <span style=\"color:black\">1. <a href=\"https://microsoftit.visualstudio.com/OneITVSO/_workitems/edit/4889949\">SCENARIO 4889949: </a>[Assessment] Plan and Assessments for CSEO applications. it will as well add a child link UserStory to that feature.</span>");
                //builder.Append("<br/><span style=\"color:black\">2. <a href=\"https://microsoftit.visualstudio.com/OneITVSO/_workitems/edit/4889951\">SCENARIO 4889951: </a>[Assessment] Assist Non-CSEO organizations with Assessment tasks. It will as well add a child link UserStory to the that feature.</span>");
                builder.Append("<br/><span style=\"color:black\">2. <a href=\"https://microsoftit.visualstudio.com/OneITVSO/_workitems/edit/5324882\">SCENARIO 5324882: </a>[Grade Review] Plan and Execute CSEO P3 Applications Grade Reviews. It will as well add a child link UserStory to the that feature</span>" +
                    "</p> <br/>" +
                    "<p class=\"MsoNormal\" style=\"color:black;\"> P.S- You need to manually define the iteration for the feature/UserStory created above. The tag GetHealthy/StayHealthy will be added automatically by the tool. <o:p>&nbsp;</o:p></p>" +
                    "<p class=\"MsoNormal\"> <o:p>&nbsp;</o:p></p>");
                builder.Append("<p class=\"MsoNormal\">" +
                    "<span class=\"Heading2Char\">" +
                    "<span style=\"font-size:13.0pt\"><b>Newly Created ADO Features and UserStories:</b></span> </span> <span style=\"font-size:10.0pt;font-family:&quot;Segoe UI Semibold&quot;,sans-serif\"> <o:p></o:p> </span></p><p class=\"MsoNormal\"> <span style=\"color:black\">");
                builder.Append("<table class=\"TFtable\">");
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
                builder.Append("</span></p><p class=\"MsoNormal\"> <span class=\"Heading2Char\"> <br/><span style=\"font-size:13.0pt\"><b>Next Steps:</b></span> </span> <span style=\"font-size:10.0pt;font-family:&quot;Segoe UI Semibold&quot;,sans-serif\"> <o:p></o:p> </span></p>" +
                    "<p class=\"MsoNormal\" style=\"color: black;\">For the above UserStory you needto create t6wo tasks (E.g. Sample tasks below)</p>" +
                    "<p class=\"MsoNormal\" style=\"color: black;\">1. Task 1 => [Activity.Onboarding][2002][RecID_1491][CFE]Succession Planning</p>" +
                    "<p class=\"MsoNormal\" style=\"color: black;\">2. Task 2 => [Activity.Assessment][2002][RecID_1491][CFE]Succession Planning</p> " +
                    "<br/><p class=\"MsoNormal\" style=\"color: black;\">Please add the tags as well define the iterations for the tasks.</p>");
                builder.Append("<br/><p class=\"MsoNormal\"> <span class=\"Heading2Char\"> <span style=\"font-size:13.0pt\"><b>Overview:</b></span> </span> <span style=\"font-size:10.0pt;font-family:&quot;Segoe UI Semibold&quot;,sans-serif\"> <o:p></o:p> </span></p>");
                builder.Append("<p class=\"MsoNormal\"> <span style=\"color:black\"><table class=\"TFtable\"><colgroup span=\"2\"></colgroup><colgroup span=\"2\"></colgroup><colgroup span=\"2\"></colgroup><colgroup span=\"2\"></colgroup><tr><th colspan=\"2\" scope=\"colgroup\">Group/SubGroup</th><th colspan=\"3\" scope=\"colgroup\">Priority In AIRT</th><th colspan=\"3\" scope=\"colgroup\">Active / New Features in ADO</th><th colspan=\"3\" scope=\"colgroup\">Active / New User Story in ADO</th></tr><tr><th>Group</th><th>Sub Group</th><th>P1</th><th>P2</th><th>P3</th><th>P1</th><th>P2</th><th>P3</th><th>P1</th><th>P2</th><th>P3</th></tr>");
                builder.Append("<tr>");
                builder.Append("<td>Text</td>");
                builder.Append("<td>Text</td>");
                builder.Append("<td>0</td>");
                builder.Append("<td>0</td>");
                builder.Append("<td>0</td>");
                builder.Append("<td>0</td>");
                builder.Append("<td>0</td>");
                builder.Append("<td>0</td>");
                builder.Append("<td>0</td>");
                builder.Append("<td>0</td>");
                builder.Append("<td>0</td>");
                builder.Append("</tr>");
                //builder.Append("<tr style=\"height:64.35pt\"><td width=\"837\" valign=\"top\" style=\"width:627.6pt;background:#F2F2F2;padding:0in .3in 0in .3in;height:64.35pt\"> <br/><p class=\"MsoNormal\"> <span class=\"MsoIntenseEmphasis\"><b>Note: </b>CSEO Accessability Services ADO features and User story creator tool will create features and user storys only CSEO applications P1, P2 and P3 applications where TAGs doesn’t contains Non-Inventory.</span> <span class=\"MsoIntenseEmphasis\"> <span style=\"font-family:&quot;Calibri&quot;,sans-serif\"> <o:p></o:p> </span> </span></p></td></tr><tr style=\"height:64.35pt\"><td width=\"837\" valign=\"top\" style=\"width:627.6pt;background:#F2F2F2;padding:0in .3in 0in .3in;height:64.35pt\"><p class=\"MsoNormal\"> <o:p>&nbsp;</o:p></p></td></tr><tr style=\"height:41.85pt\"><td width=\"837\" valign=\"top\" style=\"width:627.6pt;background:#c1e4ec;padding:0in 0in 0in .3in;height:41.85pt\"><table class=\"MsoNormalTable\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\" style=\"border-collapse:collapse\"><tr style=\"height:27.9pt\"><td width=\"499\" valign=\"top\" style=\"width:374.25pt;padding:.15in .2in .15in 0in;height:27.9pt\"><p class=\"MsoNormal\"> <span style=\"font-size:8.0pt;color:#262626\"><b>Connect <a href=\"mailto:CSEOA11yteam@microsoft.com\">CSEO Accessibility Team</a></b></span></p><p class=\"MsoNormal\"> <span style=\"font-size:8.0pt;color:#262626\">Please visit our </span> <a href=\"https://microsoft.sharepoint.com/teams/meconnection/SitePages/Accessibility-Tags.aspx\"> <span style=\"font-size:8.0pt\">Accessbility Hub Website</span> </a> <span style=\"font-size:8.0pt\">to learn more. Microsoft <span style=\"color:#262626\">C</span>onfidential. </span> <span style=\"font-size:8.0pt;color:#282828;mso-fareast-language:IT\"> <o:p></o:p> </span></p></td><td width=\"144\" nowrap=\"\" style=\"width:107.75pt;border:solid windowtext 1.0pt;border-left:none;background:#282828;padding:0in 0in 0in 0in;height:27.9pt\"><p class=\"MsoNormal\" align=\"center\" style=\"text-align:center\"> <span style=\"color:white\"> <img border=\"0\" width=\"96\" height=\"42\" style=\"width:1.0in;height:.4375in\" id=\"_x0000_i1025\" src=\"https://www.freepnglogos.com/uploads/microsoft-logo-png-transparent-20.png\"> </span> <span style=\"font-size:9.0pt;color:white\"> <o:p></o:p> </span></p></td></tr></table></td></tr></table></div><p class=\"MsoNormal\"> <o:p>&nbsp;</o:p></p></div></body></html>");
                builder.Append("</table> </span></p></td></tr><tr style=\"height:64.35pt\"><td width=\"837\" valign=\"top\" style=\"width:627.6pt;background:#F2F2F2;padding:0in .3in 0in .3in;height:64.35pt\"> <br/>" +
                    "<p class=\"MsoNormal\"> <span class=\"MsoIntenseEmphasis\" style=\"color: black;\"><b>Note: </b>CSEO Accessability Services ADO features and User story creator tool will create features and user storys only CSEO applications P1, P2 and P3 applications where TAGs doesn’t contains Non-Inventory.</span> <span class=\"MsoIntenseEmphasis\"> <span style=\"font-family:&quot;Calibri&quot;,sans-serif\"> <o:p></o:p> </span> </span></p></td></tr><tr style=\"height:64.35pt\"><td width=\"837\" valign=\"top\" style=\"width:627.6pt;background:#F2F2F2;padding:0in .3in 0in .3in;height:64.35pt\"><p class=\"MsoNormal\"> <o:p>&nbsp;</o:p></p></td></tr><tr style=\"height:41.85pt\"><td width=\"837\" valign=\"top\" style=\"width:627.6pt;background:#c1e4ec;padding:0in 0in 0in .3in;height:41.85pt\"><table class=\"MsoNormalTable\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\" style=\"border-collapse:collapse\"><tr style=\"height:27.9pt\"><td width=\"499\" valign=\"top\" style=\"width:374.25pt;padding:.15in .2in .15in 0in;height:27.9pt\"><p class=\"MsoNormal\"> <span style=\"font-size:8.0pt;color:#262626\"><b>Connect <a href=\"mailto:CSEOA11yteam@microsoft.com\">CSEO Accessibility Team</a></b></span></p><p class=\"MsoNormal\"> <span style=\"font-size:8.0pt;color:#262626\">Please visit our </span> <a href=\"https://microsoft.sharepoint.com/teams/meconnection/SitePages/Accessibility-Tags.aspx\"> <span style=\"font-size:8.0pt\">Accessbility Hub Website</span> </a> <span style=\"font-size:8.0pt\">to learn more. Microsoft <span style=\"color:#262626\">C</span>onfidential. </span> <span style=\"font-size:8.0pt;color:#282828;mso-fareast-language:IT\"> <o:p></o:p> </span></p></td><td width=\"144\" nowrap=\"\" style=\"width:107.75pt;border:solid windowtext 1.0pt;border-left:none;background:#282828;padding:0in 0in 0in 0in;height:27.9pt\"><p class=\"MsoNormal\" align=\"center\" style=\"text-align:center\"> <span style=\"color:white\"> <img border=\"0\" width=\"96\" height=\"42\" style=\"width:1.0in;height:.4375in\" id=\"_x0000_i1025\" src=\"https://www.freepnglogos.com/uploads/microsoft-logo-png-transparent-20.png\"> </span> <span style=\"font-size:9.0pt;color:white\"> <o:p></o:p> </span></p></td></tr></table></td></tr></table></div><p class=\"MsoNormal\"> <o:p>&nbsp;</o:p></p></div></body></html>");
                oMsg.Subject = "CSEO Accessbility Services - Newly Created Features in ADO : " + System.DateTime.Now.ToShortDateString() + " " + System.DateTime.Now.ToShortTimeString();
            }
            string HtmlFile = builder.ToString();

            /* Live */
            //oMsg.To = ConfigurationManager.AppSettings["DftSendMailTo"];
            //oMsg.CC = ConfigurationManager.AppSettings["DftSendMailCC"];

            /* QA */
            oMsg.To = ConfigurationManager.AppSettings["QA_DftSendMailTo"];
            oMsg.CC = ConfigurationManager.AppSettings["QA_DftSendMailTo"];

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