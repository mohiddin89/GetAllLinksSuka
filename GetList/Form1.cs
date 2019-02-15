using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Entities;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using Microsoft.SharePoint.Client.WebParts;
using System.Xml;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using System.Xml.Linq;

namespace GetList
{
    public partial class Form1 : System.Windows.Forms.Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        int FoldCount = 0;
        int FileCount = 0;
        int ListFoldCount = 0;
        int ListFileCount = 0;
        public static Web web;

        public static string siteTitle;

        Dictionary<string, int> BuiltinGroups = new Dictionary<string, int>();
        Dictionary<string, int> ADGroups = new Dictionary<string, int>();
        List<string> lstADGroupsColl = new List<string>();
        //List<string> ADGroups = new List<string>();

        private void button1_Click(object sender, EventArgs e)
        {

            #region AD Groups CSV Reading

            //lstADGroupsColl.Clear();

            //if (!string.IsNullOrEmpty(textBox3.Text))
            //{
            //    StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox3.Text));

            //    while (!sr.EndOfStream)
            //    {
            //        try
            //        {
            //            lstADGroupsColl.Add(sr.ReadLine().Trim().ToLower());
            //        }
            //        catch
            //        {
            //            continue;
            //        }
            //    }
            //}
            //else
            //{
            //    MessageBox.Show("Please browse the path for ADGroups.csv");
            //}

            #endregion

            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();

            //if (!string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrEmpty(textBox2.Text))
            {
                StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

                while (!sr.EndOfStream)
                {
                    try
                    {
                        lstSiteColl.Add(sr.ReadLine().Trim());
                    }
                    catch
                    {
                        continue;
                    }
                }
            }

            //else
            //{
            //    MessageBox.Show("Please browse the path for SiteColl.csv / Reports folder");
            //}

            #endregion

            //StreamWriter excelWriterScoringMatrixNew = null;

            //excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "ScoringMatrix" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");

            //excelWriterScoringMatrixNew.WriteLine("Filename" + "," + "URL" + "," + "Owners" + "," + "Built-in-Groups" + "," + "AD Groups" + "," + "Start Time" + "," + "End Date" + "," + "Remarks");
            //excelWriterScoringMatrixNew.Flush();

            List<string> ListNames = new List<string>();
            ListNames.Add("Site Assets");
            ListNames.Add("2_Documents and Pages");
            ListNames.Add("1_Uploaded Files");
            ListNames.Add("Discussions");

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();
                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString().Trim(), "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        clientcontext.Load(clientcontext.Web);
                        clientcontext.ExecuteQuery();

                        ListCollection _Lists = clientcontext.Web.Lists;
                        clientcontext.Load(_Lists);
                        clientcontext.ExecuteQuery();

                        foreach (string listName in ListNames)
                        {
                            try
                            {
                                bool _dListExist = clientcontext.Web.Lists.Cast<List>().Any(list => string.Equals(list.Title, listName));

                                if (_dListExist)
                                {
                                    if (listName == "Discussions")
                                    {
                                        // try
                                        {
                                            List Pagelist = clientcontext.Web.Lists.GetByTitle(listName);
                                            clientcontext.Load(Pagelist);
                                            clientcontext.ExecuteQuery();

                                            ViewCollection ViewColl = Pagelist.Views;
                                            clientcontext.Load(ViewColl);
                                            clientcontext.ExecuteQuery();

                                            Microsoft.SharePoint.Client.View v = ViewColl.GetByTitle("Featured Discussions");
                                            clientcontext.Load(v);
                                            clientcontext.ExecuteQuery();

                                            v.DeleteObject();
                                            clientcontext.ExecuteQuery();
                                        }
                                    }
                                    else if (listName == "2_Documents and Pages")
                                    {
                                        List Pagelist = clientcontext.Web.Lists.GetByTitle(listName);
                                        clientcontext.Load(Pagelist);
                                        clientcontext.ExecuteQuery();

                                        ViewCollection ViewColl = Pagelist.Views;
                                        clientcontext.Load(ViewColl);
                                        clientcontext.ExecuteQuery();

                                        Microsoft.SharePoint.Client.View v = ViewColl[0];
                                        clientcontext.Load(v);
                                        clientcontext.ExecuteQuery();

                                        v.ViewFields.RemoveAll();
                                        v.Update();
                                        clientcontext.ExecuteQuery();

                                        v.ViewFields.Add("DocIcon");
                                        v.ViewFields.Add("Title");
                                        v.ViewFields.Add("LinkFilename");
                                        v.ViewFields.Add("Created");
                                        v.ViewFields.Add("Created By");
                                        v.ViewFields.Add("Modified");
                                        v.ViewFields.Add("Modified By");
                                        v.ViewFields.Add("Checked Out To");
                                        v.ViewFields.Add("Tags");
                                        v.ViewFields.Add("Categorization");
                                        v.ViewFields.Add("Approval Status");
                                        v.Update();
                                        clientcontext.ExecuteQuery();
                                    }
                                    else
                                    {
                                        List Pagelist = clientcontext.Web.Lists.GetByTitle(listName);
                                        clientcontext.Load(Pagelist);
                                        clientcontext.ExecuteQuery();

                                        ViewCollection ViewColl = Pagelist.Views;
                                        clientcontext.Load(ViewColl);
                                        clientcontext.ExecuteQuery();

                                        Microsoft.SharePoint.Client.View v = ViewColl[0];
                                        clientcontext.Load(v);
                                        clientcontext.ExecuteQuery();

                                        v.ViewFields.RemoveAll();
                                        v.Update();
                                        clientcontext.ExecuteQuery();

                                        v.ViewFields.Add("DocIcon");
                                        v.ViewFields.Add("Title");
                                        v.ViewFields.Add("LinkFilename");
                                        v.ViewFields.Add("Created");
                                        v.ViewFields.Add("Created By");
                                        v.ViewFields.Add("Modified");
                                        v.ViewFields.Add("Modified By");
                                        v.ViewFields.Add("Tags");
                                        v.ViewFields.Add("Categorization");
                                        v.Update();
                                        clientcontext.ExecuteQuery();
                                    }

                                    //Pagelist.ContentTypesEnabled = true;
                                    //Pagelist.Update();
                                    //clientcontext.ExecuteQuery();
                                }
                            }
                            catch (Exception ex)
                            {
                                continue;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
                #region OLD
                //    this.Text = lstSiteColl[j] + "  Processing...";

                //    string startingTime = DateTime.Now.ToString();

                //    try
                //    {
                //        siteTitle = string.Empty;
                //        AuthenticationManager authManager = new AuthenticationManager();

                //        List<string> SPoExist = new List<string>();

                //        SPoExist.Add("spo.admin.verinon@agilent.onmicrosoft.com");
                //        SPoExist.Add("spo.admin2.verinon@agilent.onmicrosoft.com");
                //        SPoExist.Add("spo.admin3.verinon@agilent.onmicrosoft.com");
                //        SPoExist.Add("spo.admin4.verinon@agilent.onmicrosoft.com");
                //        SPoExist.Add("spo.admin5.verinon@agilent.onmicrosoft.com");

                //        string actualSPO = string.Empty;

                //        foreach (string sp in SPoExist)
                //        {
                //            try
                //            {
                //                using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString().Trim(), sp, "Lot62215"))
                //                {
                //                    clientcontext.Load(clientcontext.Web);
                //                    clientcontext.ExecuteQuery();

                //                    actualSPO = sp;

                //                    break;
                //                }
                //            }
                //            catch (Exception ex)
                //            {
                //                continue;
                //            }
                //        }

                //        //using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString().Trim(), "adam.a@VerinonTechnology.onmicrosoft.com", "Lot62215##"))
                //        using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString().Trim(), actualSPO, "Lot62215"))
                //        {


                //            ListCollection _Lists = clientcontext.Web.Lists;
                //            clientcontext.Load(_Lists);
                //            clientcontext.ExecuteQuery();

                //            bool _dListExist = clientcontext.Web.Lists.Cast<List>().Any(list => string.Equals(list.Title, "Site Assets"));

                //            if (_dListExist)
                //            {
                //                List Pagelist = clientcontext.Web.Lists.GetByTitle("Site Assets");
                //                clientcontext.Load(Pagelist);
                //                clientcontext.Load(Pagelist.RootFolder);
                //                clientcontext.ExecuteQuery();

                //                Pagelist.ContentTypesEnabled = true;
                //                Pagelist.Update();
                //                clientcontext.ExecuteQuery();
                //            }

                //            string admins = string.Empty;
                //            List<UserEntity> adminsColl = clientcontext.Site.RootWeb.GetAdministrators();

                //            foreach (UserEntity admin in adminsColl)
                //            {//SPO Admin 

                //                //User adUser = clientcontext.Site.RootWeb.SiteUsers.GetByLoginName(admin.LoginName);
                //                //adUser.is

                //                if (admin.Title != "FUN-SPO-SITECOLL-ADMINS" && (!admin.Title.ToLower().Contains("spo admin")) && (!admin.Email.ToLower().Contains("spo.admin@agilent.onmicrosoft.com")))
                //                {
                //                    if (!string.IsNullOrEmpty(admin.Email))
                //                    {
                //                        admins += admin.Email + ";";
                //                    }
                //                    else
                //                    {
                //                        admins += admin.Title + ";";
                //                    }
                //                }
                //            }

                //            Web oWebcurr = clientcontext.Site.RootWeb;
                //            clientcontext.Load(oWebcurr);
                //            clientcontext.ExecuteQuery();

                //            BuiltinGroups.Clear();
                //            ADGroups.Clear();

                //            siteTitle = oWebcurr.Title;

                //            string siteCollName = siteTitle.Replace(" ", "_");

                //            siteCollName = siteCollName.Replace("//", "_");

                //            string siteCollNameFileName = string.Empty;

                //            StreamWriter excelWriterScoringNew = null;

                //            if (!string.IsNullOrEmpty(siteTitle))
                //            {
                //                siteCollNameFileName = siteCollName;
                //                excelWriterScoringNew = System.IO.File.CreateText(textBox2.Text + "\\" + siteCollName + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
                //            }
                //            else
                //            {
                //                string[] siteCollNameFileNameXX = lstSiteColl[j].ToString().Trim().Split(new char[] { '/' });

                //                string actName = siteCollNameFileNameXX[siteCollNameFileNameXX.Length - 1];

                //                actName = actName.Replace(" ", "_");
                //                actName = actName.Replace("\\", "_");

                //                siteCollNameFileName = actName;
                //                excelWriterScoringNew = System.IO.File.CreateText(textBox2.Text + "\\" + actName + "_Report_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
                //            }

                //            excelWriterScoringNew.WriteLine("Site Coll Owners" + "," + admins + "," + "" + "," + "" + "," + "" + "," + "" + "," + "" + "," + "" + "," + "" + "," + "" + "," + "" + "," + "" + "," + "" + "," + "" + "," + "");
                //            excelWriterScoringNew.Flush();


                //            excelWriterScoringNew.WriteLine("Object Type" + "," + "URL" + "," + "Group" + "," + "Given though" + "," + "Folders" + "," + "Files" + "," + "Design" + "," + "Contribute" + "," + "Read" + "," + "Full Control" + "," + "Edit" + "," + "View Only" + "," + "Approve" + "," + "Contribute Limited" + "," + "OtherPermissions");
                //            excelWriterScoringNew.Flush();

                //            //////excelWriterScoringNew.WriteLine("Object Type" + "," + "URL" + "," + "SiteCollection Owners" + "," + "AD Group/Everyone granted directly" + "," + "Granted directly/added inside SP-Group" + "," + "Total number of Folders" + "," + "Total number of Files" + "," + "Design" + "," + "Contribute" + "," + "Design" + "," + "Design" + "," + "Design" + "," + "Design" + "," + "Design" + "," + "Design" + "," + "Design");
                //            //excelWriterScoringNew.WriteLine("Object Type" + "," + "URL" + "," + "SiteCollection Owners" + "," + "AD Group/Everyone granted directly" + "," + "Granted directly/added inside SP-Group" + "," + "Total number of Folders" + "," + "Total number of Files" + "," + "Design" + "," + "Contribute" + "," + "Read" + "," + "Full Control" + "," + "Edit" + "," + "View Only" + "," + "Approve" + "," + "Contribute Limited" + "," + "OtherPermissions");
                //            //excelWriterScoringNew.Flush();

                //            #region Site Coll

                //            RoleAssignmentCollection webRoleAssignments = null;
                //            GroupCollection webGroups = null;

                //            try
                //            {
                //                webRoleAssignments = clientcontext.Web.RoleAssignments;
                //                clientcontext.Load(webRoleAssignments);
                //                clientcontext.ExecuteQuery();

                //                clientcontext.Load(clientcontext.Web);
                //                clientcontext.ExecuteQuery();

                //                webGroups = clientcontext.Web.SiteGroups;
                //                clientcontext.Load(webGroups);
                //                clientcontext.ExecuteQuery();

                //                bool foundatSiteLevel = false;

                //                string AdGroupsinGroup = string.Empty;
                //                string AdGroupsatSite = string.Empty;

                //                foreach (RoleAssignment member1 in webRoleAssignments)
                //                { //c:0u.c|tenant|                             

                //                    try
                //                    {
                //                        //if (!foundatSiteLevel)
                //                        //{
                //                        clientcontext.Load(member1.Member);
                //                        clientcontext.ExecuteQuery();

                //                        if (member1.Member.Title.Contains("c:0u.c|tenant|"))
                //                        {
                //                            continue;
                //                        }

                //                        #region Role Definations

                //                        RoleDefinitionBindingCollection rdefColl = member1.RoleDefinitionBindings;
                //                        clientcontext.Load(rdefColl);
                //                        clientcontext.ExecuteQuery();

                //                        string Design = string.Empty;
                //                        string Contribute = string.Empty;
                //                        string Read = string.Empty;
                //                        string FullControl = string.Empty;
                //                        string Edit = string.Empty;
                //                        string ViewOnly = string.Empty;
                //                        string Approve = string.Empty;
                //                        string ContributeLimited = string.Empty;
                //                        string OtherPermissions = string.Empty;

                //                        foreach (RoleDefinition rdef in rdefColl)
                //                        {
                //                            clientcontext.Load(rdef);
                //                            clientcontext.ExecuteQuery();

                //                            switch (rdef.Name)
                //                            {
                //                                case "Design":
                //                                    Design = "Yes";
                //                                    break;

                //                                case "Contribute":
                //                                    Contribute = "Yes";
                //                                    break;

                //                                case "Read":
                //                                    Read = "Yes";
                //                                    break;

                //                                case "Full Control":
                //                                    FullControl = "Yes";
                //                                    break;

                //                                case "Edit":
                //                                    Edit = "Yes";
                //                                    break;

                //                                case "View Only":
                //                                    ViewOnly = "Yes";
                //                                    break;

                //                                case "Contribute Limited":
                //                                    ContributeLimited = "Yes";
                //                                    break;

                //                                case "Approve":
                //                                    Approve = "Yes";
                //                                    break;

                //                                default:
                //                                    OtherPermissions = rdef.Name;
                //                                    break;
                //                            }
                //                        }

                //                        #endregion

                //                        if (member1.Member.PrincipalType == PrincipalType.SharePointGroup)
                //                        {
                //                            Group ouserGroup = (Group)member1.Member.TypedObject;
                //                            clientcontext.Load(ouserGroup);
                //                            clientcontext.ExecuteQuery();

                //                            UserCollection userColl = ouserGroup.Users;
                //                            clientcontext.Load(userColl);
                //                            clientcontext.ExecuteQuery();

                //                            foreach (User xUser in userColl)
                //                            {
                //                                if (xUser.PrincipalType == PrincipalType.SecurityGroup)
                //                                {
                //                                    //excelWriterScoringNew.WriteLine("\"" + "SiteCollection" + "\"" + "," + "\"" + clientcontext.Web.Url.ToString() + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + "--" + "\"" + "," + "\"" + ouserGroup.Title + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"");
                //                                    //AdGroupsinGroup += ouserGroup.Title + "; ";
                //                                    //break;
                //                                    if (lstADGroupsColl.Contains(xUser.Title.ToString().Trim().ToLower()))
                //                                    {
                //                                        if (((!string.IsNullOrEmpty(Design)) || (!string.IsNullOrEmpty(Contribute)) || (!string.IsNullOrEmpty(Read)) || (!string.IsNullOrEmpty(FullControl)) || (!string.IsNullOrEmpty(Edit)) || (!string.IsNullOrEmpty(ViewOnly)) || (!string.IsNullOrEmpty(Approve)) || (!string.IsNullOrEmpty(ContributeLimited))) || ((!string.IsNullOrEmpty(OtherPermissions)) && (OtherPermissions != "Limited Access")))
                //                                        {
                //                                            if (xUser.Title.ToString().ToLower().Contains("everyone"))
                //                                            {
                //                                                //if (!BuiltinGroups.Contains(xUser.Title))
                //                                                //{
                //                                                //    BuiltinGroups.Add(xUser.Title);
                //                                                //}

                //                                                if (BuiltinGroups.ContainsKey(xUser.Title))
                //                                                {
                //                                                    BuiltinGroups[xUser.Title]++;
                //                                                }
                //                                                else
                //                                                {
                //                                                    BuiltinGroups.Add(xUser.Title, 1);
                //                                                }
                //                                            }
                //                                            else
                //                                            {
                //                                                //if (!ADGroups.Contains(xUser.Title))
                //                                                //{
                //                                                //    ADGroups.Add(xUser.Title);
                //                                                //}

                //                                                if (ADGroups.ContainsKey(xUser.Title))
                //                                                {
                //                                                    ADGroups[xUser.Title]++;
                //                                                }
                //                                                else
                //                                                {
                //                                                    ADGroups.Add(xUser.Title, 1);
                //                                                }
                //                                            }

                //                                            excelWriterScoringNew.WriteLine("\"" + "SiteCollection" + "\"" + "," + "\"" + clientcontext.Web.Url.ToString() + "\"" + "," + "\"" + xUser.Title + "\"" + "," + "\"" + ouserGroup.Title + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + Design + "\"" + "," + "\"" + Contribute + "\"" + "," + "\"" + Read + "\"" + "," + "\"" + FullControl + "\"" + "," + "\"" + Edit + "\"" + "," + "\"" + ViewOnly + "\"" + "," + "\"" + Approve + "\"" + "," + "\"" + ContributeLimited + "\"" + "," + "\"" + OtherPermissions + "\"");
                //                                            excelWriterScoringNew.Flush();
                //                                        }
                //                                    }
                //                                    //foundatSiteLevel = true;
                //                                    //break;
                //                                }

                //                                //if (xUser.Title == "Everyone except external users")
                //                                //{
                //                                //    excelWriterScoringNew.WriteLine("\"" + "SiteCollection" + "\"" + "," + "\"" + clientcontext.Web.Url.ToString() + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + "Yes" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"");

                //                                //    foundatSiteLevel = true;
                //                                //    break;
                //                                //}
                //                            }
                //                        }
                //                        if (member1.Member.PrincipalType == PrincipalType.SecurityGroup)
                //                        {
                //                            //if (member1.Member.Title == "Everyone except external users")
                //                            //{
                //                            //excelWriterScoringNew.WriteLine("\"" + "SiteCollection" + "\"" + "," + "\"" + clientcontext.Web.Url.ToString() + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + member1.Member.Title + "\"" + "," + "\"" + "--" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"");
                //                            //AdGroupsatSite += member1.Member.Title + "; ";

                //                            if (lstADGroupsColl.Contains(member1.Member.Title.ToString().Trim().ToLower()))
                //                            {
                //                                if (((!string.IsNullOrEmpty(Design)) || (!string.IsNullOrEmpty(Contribute)) || (!string.IsNullOrEmpty(Read)) || (!string.IsNullOrEmpty(FullControl)) || (!string.IsNullOrEmpty(Edit)) || (!string.IsNullOrEmpty(ViewOnly)) || (!string.IsNullOrEmpty(Approve)) || (!string.IsNullOrEmpty(ContributeLimited))) || ((!string.IsNullOrEmpty(OtherPermissions)) && (OtherPermissions != "Limited Access")))
                //                                {
                //                                    if (member1.Member.Title.ToString().ToLower().Contains("everyone"))
                //                                    {
                //                                        //if (!BuiltinGroups.Contains(member1.Member.Title))
                //                                        //{
                //                                        //    BuiltinGroups.Add(member1.Member.Title);
                //                                        //}

                //                                        if (BuiltinGroups.ContainsKey(member1.Member.Title))
                //                                        {
                //                                            BuiltinGroups[member1.Member.Title]++;
                //                                        }
                //                                        else
                //                                        {
                //                                            BuiltinGroups.Add(member1.Member.Title, 1);
                //                                        }
                //                                    }
                //                                    else
                //                                    {
                //                                        //if (!ADGroups.Contains(member1.Member.Title))
                //                                        //{
                //                                        //    ADGroups.Add(member1.Member.Title);
                //                                        //}
                //                                        if (ADGroups.ContainsKey(member1.Member.Title))
                //                                        {
                //                                            ADGroups[member1.Member.Title]++;
                //                                        }
                //                                        else
                //                                        {
                //                                            ADGroups.Add(member1.Member.Title, 1);
                //                                        }
                //                                    }

                //                                    excelWriterScoringNew.WriteLine("\"" + "SiteCollection" + "\"" + "," + "\"" + clientcontext.Web.Url.ToString() + "\"" + "," + "\"" + member1.Member.Title + "\"" + "," + "\"" + "Directly Assigned" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + Design + "\"" + "," + "\"" + Contribute + "\"" + "," + "\"" + Read + "\"" + "," + "\"" + FullControl + "\"" + "," + "\"" + Edit + "\"" + "," + "\"" + ViewOnly + "\"" + "," + "\"" + Approve + "\"" + "," + "\"" + ContributeLimited + "\"" + "," + "\"" + OtherPermissions + "\"");
                //                                    excelWriterScoringNew.Flush();
                //                                }
                //                            }
                //                            //foundatSiteLevel = true;
                //                            //break;
                //                            //}
                //                        }

                //                        #region Commented Is Uesr

                //                        //if (member1.Member.PrincipalType == PrincipalType.User)
                //                        //{
                //                        //    if (member1.Member.Title == "Everyone except external users")
                //                        //    {
                //                        //        excelWriterScoringNew.WriteLine("\"" + "Site" + "\"" + "," + "\"" + clientcontext.Web.Url.ToString() + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + "Yes" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"");
                //                        //        foundatSiteLevel = true;
                //                        //        break;
                //                        //    }
                //                        //} 

                //                        #endregion
                //                        //}
                //                        //else
                //                        //{
                //                        //    break;
                //                        //}
                //                    }
                //                    catch (Exception ex)
                //                    {
                //                        continue;
                //                    }
                //                }
                //                //excelWriterScoringNew.WriteLine("\"" + "SiteCollection" + "\"" + "," + "\"" + clientcontext.Web.Url.ToString() + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + AdGroupsatSite + "\"" + "," + "\"" + AdGroupsinGroup + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"");
                //                //excelWriterScoringNew.Flush();
                //            }
                //            catch (Exception ex)
                //            {
                //            }

                //            #endregion

                //            #region Lists

                //            ListCollection olistColl = clientcontext.Web.Lists;
                //            clientcontext.Load(olistColl);
                //            clientcontext.ExecuteQuery();

                //            foreach (List oList in olistColl)
                //            {
                //                bool foundatListLevel = false;

                //                clientcontext.Load(oList);
                //                clientcontext.Load(oList, li => li.HasUniqueRoleAssignments);
                //                clientcontext.ExecuteQuery();

                //                if (oList.BaseType == BaseType.DocumentLibrary)
                //                {
                //                    if (oList.Title == "Documents")
                //                    {
                //                        bool foXXXXundatListLevel = false;
                //                    }

                //                    if ((oList.Title != "Form Templates" && oList.Title != "Site Assets" && oList.Title != "SitePages" && oList.Title != "Style Library" && oList.Hidden == false && oList.IsCatalog == false && oList.BaseTemplate == 101) || oList.BaseTemplate == 700)
                //                    {
                //                        string UniqueRoles = string.Empty;

                //                        #region Commented Test

                //                        //if (oList.Title == "Documents")
                //                        //{
                //                        //    clientcontext.Load(oList.RootFolder);
                //                        //    clientcontext.ExecuteQuery();

                //                        //    clientcontext.Load(clientcontext.Web);
                //                        //    clientcontext.ExecuteQuery();

                //                        //    GetCounts(oList.RootFolder, clientcontext);
                //                        //}

                //                        #endregion

                //                        //if (oList.HasUniqueRoleAssignments)
                //                        //{
                //                        //    UniqueRoles = "Unique Permissions";
                //                        //}
                //                        //else
                //                        //{
                //                        //    UniqueRoles = "Inherit from Parent";
                //                        //}

                //                        if (oList.HasUniqueRoleAssignments)
                //                        {
                //                            ListFoldCount = 0;
                //                            ListFileCount = 0;

                //                            GetCountsatListLevel(oList.RootFolder, clientcontext);

                //                            string AdGroupsinSileCollListGroup = string.Empty;
                //                            string AdGroupsatSileCollListSite = string.Empty;

                //                            #region SiteColl Lists Permission check

                //                            RoleAssignmentCollection roles = oList.RoleAssignments;
                //                            clientcontext.Load(roles);
                //                            clientcontext.ExecuteQuery();

                //                            Web oWebx = clientcontext.Web;
                //                            clientcontext.Load(oWebx);
                //                            clientcontext.ExecuteQuery();

                //                            foreach (RoleAssignment rAssignment in roles)
                //                            {


                //                                #region Role Definations

                //                                RoleDefinitionBindingCollection rdefColl = rAssignment.RoleDefinitionBindings;
                //                                clientcontext.Load(rdefColl);
                //                                clientcontext.ExecuteQuery();

                //                                string Design = string.Empty;
                //                                string Contribute = string.Empty;
                //                                string Read = string.Empty;
                //                                string FullControl = string.Empty;
                //                                string Edit = string.Empty;
                //                                string ViewOnly = string.Empty;
                //                                string Approve = string.Empty;
                //                                string ContributeLimited = string.Empty;
                //                                string OtherPermissions = string.Empty;

                //                                foreach (RoleDefinition rdef in rdefColl)
                //                                {
                //                                    clientcontext.Load(rdef);
                //                                    clientcontext.ExecuteQuery();

                //                                    switch (rdef.Name)
                //                                    {
                //                                        case "Design":
                //                                            Design = "Yes";
                //                                            break;

                //                                        case "Contribute":
                //                                            Contribute = "Yes";
                //                                            break;

                //                                        case "Read":
                //                                            Read = "Yes";
                //                                            break;

                //                                        case "Full Control":
                //                                            FullControl = "Yes";
                //                                            break;

                //                                        case "Edit":
                //                                            Edit = "Yes";
                //                                            break;

                //                                        case "View Only":
                //                                            ViewOnly = "Yes";
                //                                            break;

                //                                        case "Contribute Limited":
                //                                            ContributeLimited = "Yes";
                //                                            break;

                //                                        case "Approve":
                //                                            Approve = "Yes";
                //                                            break;

                //                                        default:
                //                                            OtherPermissions = rdef.Name;
                //                                            break;
                //                                    }
                //                                }

                //                                #endregion

                //                                try
                //                                {
                //                                    //if (!foundatListLevel)
                //                                    //{
                //                                    clientcontext.Load(rAssignment.Member);
                //                                    clientcontext.ExecuteQuery();

                //                                    if (rAssignment.Member.Title.Contains("c:0u.c|tenant|"))
                //                                    {
                //                                        continue;
                //                                    }

                //                                    if (rAssignment.Member.PrincipalType == PrincipalType.SharePointGroup)
                //                                    {
                //                                        Group ouserGroup = (Group)rAssignment.Member.TypedObject;
                //                                        clientcontext.Load(ouserGroup);
                //                                        clientcontext.ExecuteQuery();

                //                                        UserCollection userColl = ouserGroup.Users;
                //                                        clientcontext.Load(userColl);
                //                                        clientcontext.ExecuteQuery();

                //                                        foreach (User xUser in userColl)
                //                                        {
                //                                            if (xUser.PrincipalType == PrincipalType.SecurityGroup)
                //                                            {

                //                                                //if (xUser.Title == "Everyone except external users")
                //                                                //{
                //                                                //clientcontext.Load(oList.RootFolder);
                //                                                //clientcontext.ExecuteQuery();   

                //                                                //AdGroupsinSileCollListGroup += ouserGroup.Title + ";";
                //                                                //foundatListLevel = true;
                //                                                //break;

                //                                                if (lstADGroupsColl.Contains(xUser.Title.ToString().Trim().ToLower()))
                //                                                {
                //                                                    if (((!string.IsNullOrEmpty(Design)) || (!string.IsNullOrEmpty(Contribute)) || (!string.IsNullOrEmpty(Read)) || (!string.IsNullOrEmpty(FullControl)) || (!string.IsNullOrEmpty(Edit)) || (!string.IsNullOrEmpty(ViewOnly)) || (!string.IsNullOrEmpty(Approve)) || (!string.IsNullOrEmpty(ContributeLimited))) || ((!string.IsNullOrEmpty(OtherPermissions)) && (OtherPermissions != "Limited Access")))
                //                                                    {
                //                                                        if (xUser.Title.ToString().ToLower().Contains("everyone"))
                //                                                        {
                //                                                            //if (!BuiltinGroups.Contains(xUser.Title))
                //                                                            //{
                //                                                            //    BuiltinGroups.Add(xUser.Title);
                //                                                            //}

                //                                                            if (BuiltinGroups.ContainsKey(xUser.Title))
                //                                                            {
                //                                                                BuiltinGroups[xUser.Title]++;
                //                                                            }
                //                                                            else
                //                                                            {
                //                                                                BuiltinGroups.Add(xUser.Title, 1);
                //                                                            }
                //                                                        }
                //                                                        else
                //                                                        {
                //                                                            //if (!ADGroups.Contains(xUser.Title))
                //                                                            //{
                //                                                            //    ADGroups.Add(xUser.Title);
                //                                                            //}
                //                                                            if (ADGroups.ContainsKey(xUser.Title))
                //                                                            {
                //                                                                ADGroups[xUser.Title]++;
                //                                                            }
                //                                                            else
                //                                                            {
                //                                                                ADGroups.Add(xUser.Title, 1);
                //                                                            }
                //                                                        }

                //                                                        excelWriterScoringNew.WriteLine("\"" + "Doc Library" + "\"" + "," + "\"" + oWebx.Url.Replace(oWebx.ServerRelativeUrl, "") + oList.RootFolder.ServerRelativeUrl + "\"" + "," + "\"" + xUser.Title + "\"" + "," + "\"" + ouserGroup.Title + "\"" + "," + "\"" + ListFoldCount.ToString() + "\"" + "," + "\"" + ListFileCount.ToString() + "\"" + "," + "\"" + Design + "\"" + "," + "\"" + Contribute + "\"" + "," + "\"" + Read + "\"" + "," + "\"" + FullControl + "\"" + "," + "\"" + Edit + "\"" + "," + "\"" + ViewOnly + "\"" + "," + "\"" + Approve + "\"" + "," + "\"" + ContributeLimited + "\"" + "," + "\"" + OtherPermissions + "\"");
                //                                                        excelWriterScoringNew.Flush();
                //                                                        //excelWriterScoringNew.WriteLine("\"" + "Doc Library" + "\"" + "," + "\"" + oWebx.Url.Replace(oWebx.ServerRelativeUrl, "") + oList.RootFolder.ServerRelativeUrl + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + "--" + "\"" + "\"" + "," + "\"" + ouserGroup.Title + "\"" + "," + "\"" + UniqueRoles + "\"" + "," + "\"" + FoldCount.ToString() + "\"" + "," + "\"" + FileCount.ToString() + "\"");
                //                                                    }
                //                                                }

                //                                                //break;
                //                                            }
                //                                        }
                //                                    }
                //                                    if (rAssignment.Member.PrincipalType == PrincipalType.SecurityGroup)
                //                                    {
                //                                        //if (rAssignment.Member.Title == "Everyone except external users")
                //                                        //{
                //                                        //clientcontext.Load(oList.RootFolder);
                //                                        //clientcontext.ExecuteQuery();                                                  

                //                                        //AdGroupsatSileCollListSite += rAssignment.Member.Title + ";";
                //                                        //foundatListLevel = true;
                //                                        if (lstADGroupsColl.Contains(rAssignment.Member.Title.ToString().Trim().ToLower()))
                //                                        {


                //                                            if (((!string.IsNullOrEmpty(Design)) || (!string.IsNullOrEmpty(Contribute)) || (!string.IsNullOrEmpty(Read)) || (!string.IsNullOrEmpty(FullControl)) || (!string.IsNullOrEmpty(Edit)) || (!string.IsNullOrEmpty(ViewOnly)) || (!string.IsNullOrEmpty(Approve)) || (!string.IsNullOrEmpty(ContributeLimited))) || ((!string.IsNullOrEmpty(OtherPermissions)) && (OtherPermissions != "Limited Access")))
                //                                            {
                //                                                if (rAssignment.Member.Title.ToString().ToLower().Contains("everyone"))
                //                                                {
                //                                                    //if (!BuiltinGroups.Contains(rAssignment.Member.Title))
                //                                                    //{
                //                                                    //    BuiltinGroups.Add(rAssignment.Member.Title);
                //                                                    //}

                //                                                    if (BuiltinGroups.ContainsKey(rAssignment.Member.Title))
                //                                                    {
                //                                                        BuiltinGroups[rAssignment.Member.Title]++;
                //                                                    }
                //                                                    else
                //                                                    {
                //                                                        BuiltinGroups.Add(rAssignment.Member.Title, 1);
                //                                                    }
                //                                                }
                //                                                else
                //                                                {
                //                                                    //if (!ADGroups.Contains(rAssignment.Member.Title))
                //                                                    //{
                //                                                    //    ADGroups.Add(rAssignment.Member.Title);
                //                                                    //}

                //                                                    if (ADGroups.ContainsKey(rAssignment.Member.Title))
                //                                                    {
                //                                                        ADGroups[rAssignment.Member.Title]++;
                //                                                    }
                //                                                    else
                //                                                    {
                //                                                        ADGroups.Add(rAssignment.Member.Title, 1);
                //                                                    }
                //                                                }

                //                                                excelWriterScoringNew.WriteLine("\"" + "Doc Library" + "\"" + "," + "\"" + oWebx.Url.Replace(oWebx.ServerRelativeUrl, "") + oList.RootFolder.ServerRelativeUrl + "\"" + "," + "\"" + rAssignment.Member.Title + "\"" + "," + "\"" + "Directly Assigned" + "\"" + "," + "\"" + ListFoldCount.ToString() + "\"" + "," + "\"" + ListFileCount.ToString() + "\"" + "," + "\"" + Design + "\"" + "," + "\"" + Contribute + "\"" + "," + "\"" + Read + "\"" + "," + "\"" + FullControl + "\"" + "," + "\"" + Edit + "\"" + "," + "\"" + ViewOnly + "\"" + "," + "\"" + Approve + "\"" + "," + "\"" + ContributeLimited + "\"" + "," + "\"" + OtherPermissions + "\"");
                //                                                excelWriterScoringNew.Flush();
                //                                            }
                //                                        }
                //                                        //excelWriterScoringNew.WriteLine("\"" + "Doc Library" + "\"" + "," + "\"" + oWebx.Url.Replace(oWebx.ServerRelativeUrl, "") + oList.RootFolder.ServerRelativeUrl + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + "Yes" + "\"" + "," + "\"" + "--" + "\"" + "\"" + "," + "\"" + UniqueRoles + "\"" + "," + "\"" + FoldCount.ToString() + "\"" + "," + "\"" + FileCount.ToString() + "\"");


                //                                        //break;
                //                                        //}
                //                                    }
                //                                    //}
                //                                    //else
                //                                    //{
                //                                    //    break;
                //                                    //}
                //                                }
                //                                catch (Exception ex)
                //                                {
                //                                    continue;
                //                                }
                //                            }

                //                            //if (foundatListLevel)
                //                            //{
                //                            //    ListFoldCount = 0;
                //                            //    ListFileCount = 0;

                //                            //    GetCountsatListLevel(oList.RootFolder, clientcontext);

                //                            //    excelWriterScoringNew.WriteLine("\"" + "Doc Library" + "\"" + "," + "\"" + oWebx.Url.Replace(oWebx.ServerRelativeUrl, "") + oList.RootFolder.ServerRelativeUrl + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + AdGroupsatSileCollListSite + "\"" + "," + "\"" + AdGroupsinSileCollListGroup + "\"" + "," + "\"" + ListFoldCount.ToString() + "\"" + "," + "\"" + ListFileCount.ToString() + "\"");
                //                            //    excelWriterScoringNew.Flush();
                //                            //}

                //                            #endregion
                //                        }

                //                        clientcontext.Load(oList.RootFolder.Folders);
                //                        clientcontext.ExecuteQuery();

                //                        foreach (Folder sFolder in oList.RootFolder.Folders)
                //                        {
                //                            clientcontext.Load(sFolder);
                //                            clientcontext.ExecuteQuery();

                //                            if (sFolder.Name != "Forms")
                //                            {
                //                                GetCounts(sFolder, clientcontext, excelWriterScoringNew);
                //                            }
                //                        }
                //                    }
                //                }
                //            }

                //            #endregion

                //            #region SubSites

                //            WebCollection oWebs = clientcontext.Web.Webs;
                //            clientcontext.Load(oWebs);
                //            clientcontext.ExecuteQuery();

                //            foreach (Web oWeb in oWebs)
                //            {
                //                try
                //                {
                //                    clientcontext.Load(oWeb);
                //                    clientcontext.ExecuteQuery();
                //                    this.Text = oWeb.Url + "  Processing...";
                //                    getWeb(oWeb.Url, excelWriterScoringNew);
                //                }
                //                catch (Exception ex)
                //                {
                //                    continue;
                //                }
                //            }

                //            #endregion


                //            excelWriterScoringNew.Flush();
                //            excelWriterScoringNew.Close();

                //            string bGroups = string.Empty;
                //            string AdsGroups = string.Empty;

                //            foreach (KeyValuePair<string, int> kp in BuiltinGroups)
                //            {
                //                if (kp.Key != "FUN-SPO-SITECOLL-ADMINS" && (!kp.Key.ToLower().Contains("spo admin")))
                //                {
                //                    bGroups += kp.Key.ToString().Trim() + "(" + kp.Value.ToString() + ")" + "; ";
                //                }
                //            }

                //            foreach (KeyValuePair<string, int> kp in ADGroups)
                //            {
                //                if (kp.Key != "FUN-SPO-SITECOLL-ADMINS" && (!kp.Key.ToLower().Contains("spo admin")))
                //                {
                //                    AdsGroups += kp.Key.ToString().Trim() + "(" + kp.Value.ToString() + ")" + "; ";
                //                }
                //            }

                //            //foreach (string gp in BuiltinGroups)
                //            //{
                //            //    bGroups += gp + "; ";
                //            //}

                //            //foreach (string ap in ADGroups)
                //            //{
                //            //    AdsGroups += ap + "; ";
                //            //}

                //            excelWriterScoringMatrixNew.WriteLine("\"" + siteCollNameFileName + ".xlsx" + "\"" + "," + "\"" + clientcontext.Web.Url.ToString() + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + bGroups + "\"" + "," + "\"" + AdsGroups + "\"" + "," + "\"" + startingTime + "\"" + "," + "\"" + DateTime.Now.ToString() + "\"" + "," + "\"" + "" + "\"");
                //            excelWriterScoringMatrixNew.Flush();

                //            //excelWriterScoringMatrixNew.WriteLine(siteCollNameFileName +".xlsx" + "," + clientcontext.Web.Url.ToString() + "," + admins + "," + bGroups + "," + AdsGroups + "," + DateTime.Now.ToString());
                //            //excelWriterScoringMatrixNew.Flush();
                //        }
                //    }
                //    catch (Exception ex)
                //    {
                //        excelWriterScoringMatrixNew.WriteLine("\"" + "--" + "\"" + "," + "\"" + lstSiteColl[j] + "\"" + "," + "\"" + "" + "\"" + "," + "\"" + "" + "\"" + "," + "\"" + "" + "\"" + "," + "\"" + startingTime + "\"" + "," + "\"" + DateTime.Now.ToString() + "\"" + "," + "\"" + ex.Message + "\"");
                //        excelWriterScoringMatrixNew.Flush();

                //        continue;
                //    }
                //}

                //excelWriterScoringMatrixNew.Flush();
                //excelWriterScoringMatrixNew.Close();

                //this.Text = "Process completed successfully.";
                //MessageBox.Show("Process Completed"); 
                #endregion
            }
            this.Text = "Completed..";
        }
        public void GetCounts(Folder Fld, ClientContext clientcontext, StreamWriter excelWriterScoringNew)
        {

            try
            {
                clientcontext.Load(Fld);
                clientcontext.Load(Fld, li => li.ListItemAllFields.HasUniqueRoleAssignments);
                clientcontext.ExecuteQuery();

                this.Text = "Folder : " + Fld.Name + " is Processing...";

                if (Fld.Name.Contains("drophere"))
                {
                    int j = 0;
                }

                if (Fld.ListItemAllFields.HasUniqueRoleAssignments)
                {
                    FoldCount = 0;
                    FileCount = 0;

                    clientcontext.Load(Fld.Files);
                    clientcontext.ExecuteQuery();

                    FileCount += Fld.Files.Count;

                    clientcontext.Load(Fld.Folders);
                    clientcontext.ExecuteQuery();

                    foreach (Folder folder in Fld.Folders)
                    {
                        clientcontext.Load(folder);
                        clientcontext.ExecuteQuery();

                        if (folder.Name != "Forms")
                        {
                            FoldCount++;
                        }
                    }

                    #region IF Folder has Unique permissions

                    bool folderhasSecurityGroup = false;

                    string AdGroupsinFolderGroup = string.Empty;
                    string AdGroupsatFolder = string.Empty;

                    RoleAssignmentCollection roles = Fld.ListItemAllFields.RoleAssignments;
                    clientcontext.Load(roles);
                    clientcontext.ExecuteQuery();

                    foreach (RoleAssignment rAssignment in roles)
                    {
                        clientcontext.Load(rAssignment.Member);
                        clientcontext.ExecuteQuery();

                        if (rAssignment.Member.Title.Contains("c:0u.c|tenant|"))
                        {
                            continue;
                        }

                        #region Role Definations

                        RoleDefinitionBindingCollection rdefColl = rAssignment.RoleDefinitionBindings;
                        clientcontext.Load(rdefColl);
                        clientcontext.ExecuteQuery();

                        string Design = string.Empty;
                        string Contribute = string.Empty;
                        string Read = string.Empty;
                        string FullControl = string.Empty;
                        string Edit = string.Empty;
                        string ViewOnly = string.Empty;
                        string Approve = string.Empty;
                        string ContributeLimited = string.Empty;
                        string OtherPermissions = string.Empty;

                        foreach (RoleDefinition rdef in rdefColl)
                        {
                            clientcontext.Load(rdef);
                            clientcontext.ExecuteQuery();

                            switch (rdef.Name)
                            {
                                case "Design":
                                    Design = "Yes";
                                    break;

                                case "Contribute":
                                    Contribute = "Yes";
                                    break;

                                case "Read":
                                    Read = "Yes";
                                    break;

                                case "Full Control":
                                    FullControl = "Yes";
                                    break;

                                case "Edit":
                                    Edit = "Yes";
                                    break;

                                case "View Only":
                                    ViewOnly = "Yes";
                                    break;

                                case "Contribute Limited":
                                    ContributeLimited = "Yes";
                                    break;

                                case "Approve":
                                    Approve = "Yes";
                                    break;

                                default:
                                    OtherPermissions = rdef.Name;
                                    break;
                            }
                        }

                        #endregion

                        try
                        {

                            if (rAssignment.Member.PrincipalType == PrincipalType.SharePointGroup)
                            {
                                Microsoft.SharePoint.Client.Group ouserGroup = (Microsoft.SharePoint.Client.Group)rAssignment.Member.TypedObject;
                                clientcontext.Load(ouserGroup);
                                clientcontext.ExecuteQuery();

                                UserCollection userColl = ouserGroup.Users;
                                clientcontext.Load(userColl);
                                clientcontext.ExecuteQuery();

                                foreach (User xUser in userColl)
                                {
                                    if (xUser.PrincipalType == PrincipalType.SecurityGroup)
                                    {
                                        //AdGroupsinFolderGroup += ouserGroup.Title + ";";
                                        //folderhasSecurityGroup = true;
                                        //break;
                                        if (lstADGroupsColl.Contains(xUser.Title.ToString().Trim().ToLower()))
                                        {
                                            if (((!string.IsNullOrEmpty(Design)) || (!string.IsNullOrEmpty(Contribute)) || (!string.IsNullOrEmpty(Read)) || (!string.IsNullOrEmpty(FullControl)) || (!string.IsNullOrEmpty(Edit)) || (!string.IsNullOrEmpty(ViewOnly)) || (!string.IsNullOrEmpty(Approve)) || (!string.IsNullOrEmpty(ContributeLimited))) || ((!string.IsNullOrEmpty(OtherPermissions)) && (OtherPermissions != "Limited Access")))
                                            {
                                                if (xUser.Title.ToString().ToLower().Contains("everyone"))
                                                {
                                                    //if (!BuiltinGroups.Contains(xUser.Title))
                                                    //{
                                                    //    BuiltinGroups.Add(xUser.Title);
                                                    //}

                                                    if (BuiltinGroups.ContainsKey(xUser.Title))
                                                    {
                                                        BuiltinGroups[xUser.Title]++;
                                                    }
                                                    else
                                                    {
                                                        BuiltinGroups.Add(xUser.Title, 1);
                                                    }
                                                }
                                                else
                                                {
                                                    //if (!ADGroups.Contains(xUser.Title))
                                                    //{
                                                    //    ADGroups.Add(xUser.Title);
                                                    //}

                                                    if (ADGroups.ContainsKey(xUser.Title))
                                                    {
                                                        ADGroups[xUser.Title]++;
                                                    }
                                                    else
                                                    {
                                                        ADGroups.Add(xUser.Title, 1);
                                                    }
                                                }

                                                excelWriterScoringNew.WriteLine("\"" + "Folder" + "\"" + "," + "\"" + clientcontext.Web.Url.Replace(clientcontext.Web.ServerRelativeUrl, "") + Fld.ServerRelativeUrl + "\"" + "," + "\"" + xUser.Title + "\"" + "," + "\"" + ouserGroup.Title + "\"" + "," + "\"" + FoldCount.ToString() + "\"" + "," + "\"" + FileCount.ToString() + "\"" + "," + "\"" + Design + "\"" + "," + "\"" + Contribute + "\"" + "," + "\"" + Read + "\"" + "," + "\"" + FullControl + "\"" + "," + "\"" + Edit + "\"" + "," + "\"" + ViewOnly + "\"" + "," + "\"" + Approve + "\"" + "," + "\"" + ContributeLimited + "\"" + "," + "\"" + OtherPermissions + "\"");
                                                excelWriterScoringNew.Flush();
                                            }
                                        }
                                    }
                                }
                            }
                            if (rAssignment.Member.PrincipalType == PrincipalType.SecurityGroup)
                            {
                                //AdGroupsatFolder += rAssignment.Member.Title + ";";
                                //folderhasSecurityGroup = true;
                                if (lstADGroupsColl.Contains(rAssignment.Member.Title.ToString().Trim().ToLower()))
                                {
                                    if (((!string.IsNullOrEmpty(Design)) || (!string.IsNullOrEmpty(Contribute)) || (!string.IsNullOrEmpty(Read)) || (!string.IsNullOrEmpty(FullControl)) || (!string.IsNullOrEmpty(Edit)) || (!string.IsNullOrEmpty(ViewOnly)) || (!string.IsNullOrEmpty(Approve)) || (!string.IsNullOrEmpty(ContributeLimited))) || ((!string.IsNullOrEmpty(OtherPermissions)) && (OtherPermissions != "Limited Access")))
                                    {
                                        if (rAssignment.Member.Title.ToString().ToLower().Contains("everyone"))
                                        {
                                            //if (!BuiltinGroups.Contains(rAssignment.Member.Title))
                                            //{
                                            //    BuiltinGroups.Add(rAssignment.Member.Title);
                                            //}

                                            if (BuiltinGroups.ContainsKey(rAssignment.Member.Title))
                                            {
                                                BuiltinGroups[rAssignment.Member.Title]++;
                                            }
                                            else
                                            {
                                                BuiltinGroups.Add(rAssignment.Member.Title, 1);
                                            }
                                        }
                                        else
                                        {
                                            //if (!ADGroups.Contains(rAssignment.Member.Title))
                                            //{
                                            //    ADGroups.Add(rAssignment.Member.Title);
                                            //}

                                            if (ADGroups.ContainsKey(rAssignment.Member.Title))
                                            {
                                                ADGroups[rAssignment.Member.Title]++;
                                            }
                                            else
                                            {
                                                ADGroups.Add(rAssignment.Member.Title, 1);
                                            }
                                        }

                                        excelWriterScoringNew.WriteLine("\"" + "Folder" + "\"" + "," + "\"" + clientcontext.Web.Url.Replace(clientcontext.Web.ServerRelativeUrl, "") + Fld.ServerRelativeUrl + "\"" + "," + "\"" + rAssignment.Member.Title + "\"" + "," + "\"" + "Directly Assigned" + "\"" + "," + "\"" + FoldCount.ToString() + "\"" + "," + "\"" + FileCount.ToString() + "\"" + "," + "\"" + Design + "\"" + "," + "\"" + Contribute + "\"" + "," + "\"" + Read + "\"" + "," + "\"" + FullControl + "\"" + "," + "\"" + Edit + "\"" + "," + "\"" + ViewOnly + "\"" + "," + "\"" + Approve + "\"" + "," + "\"" + ContributeLimited + "\"" + "," + "\"" + OtherPermissions + "\"");
                                        excelWriterScoringNew.Flush();
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            continue;
                        }
                    }

                    //if (folderhasSecurityGroup)
                    //{
                    //    FoldCount = 0;
                    //    FileCount = 0;

                    //    clientcontext.Load(Fld.Files);
                    //    clientcontext.ExecuteQuery();

                    //    FileCount += Fld.Files.Count;

                    //    clientcontext.Load(Fld.Folders);
                    //    clientcontext.ExecuteQuery();

                    //    foreach (Folder folder in Fld.Folders)
                    //    {
                    //        clientcontext.Load(folder);
                    //        clientcontext.ExecuteQuery();

                    //        if (folder.Name != "Forms")
                    //        {
                    //            FoldCount++;
                    //        }
                    //    }

                    //    excelWriterScoringNew.WriteLine("\"" + "Folder" + "\"" + "," + "\"" + clientcontext.Web.Url.Replace(clientcontext.Web.ServerRelativeUrl, "") + Fld.ServerRelativeUrl + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + AdGroupsatFolder + "\"" + "," + "\"" + AdGroupsinFolderGroup + "\"" + "," + "\"" + FoldCount.ToString() + "\"" + "," + "\"" + FileCount.ToString() + "\"");
                    //    excelWriterScoringNew.Flush();
                    //}


                    #endregion
                }

                clientcontext.Load(Fld.Files);
                clientcontext.ExecuteQuery();

                foreach (Microsoft.SharePoint.Client.File sFile in Fld.Files)
                {
                    clientcontext.Load(sFile);
                    clientcontext.Load(sFile, li => li.ListItemAllFields.HasUniqueRoleAssignments);
                    clientcontext.ExecuteQuery();

                    if (sFile.ListItemAllFields.HasUniqueRoleAssignments)
                    {
                        bool filehasSecurityGroup = false;

                        string AdGroupsinFileGroup = string.Empty;
                        string AdGroupsatFile = string.Empty;

                        #region File Level Permission check

                        RoleAssignmentCollection roles = sFile.ListItemAllFields.RoleAssignments;
                        clientcontext.Load(roles);
                        clientcontext.ExecuteQuery();

                        foreach (RoleAssignment rAssignment in roles)
                        {

                            #region Role Definations

                            RoleDefinitionBindingCollection rdefColl = rAssignment.RoleDefinitionBindings;
                            clientcontext.Load(rdefColl);
                            clientcontext.ExecuteQuery();

                            string Design = string.Empty;
                            string Contribute = string.Empty;
                            string Read = string.Empty;
                            string FullControl = string.Empty;
                            string Edit = string.Empty;
                            string ViewOnly = string.Empty;
                            string Approve = string.Empty;
                            string ContributeLimited = string.Empty;
                            string OtherPermissions = string.Empty;

                            foreach (RoleDefinition rdef in rdefColl)
                            {
                                clientcontext.Load(rdef);
                                clientcontext.ExecuteQuery();

                                switch (rdef.Name)
                                {
                                    case "Design":
                                        Design = "Yes";
                                        break;

                                    case "Contribute":
                                        Contribute = "Yes";
                                        break;

                                    case "Read":
                                        Read = "Yes";
                                        break;

                                    case "Full Control":
                                        FullControl = "Yes";
                                        break;

                                    case "Edit":
                                        Edit = "Yes";
                                        break;

                                    case "View Only":
                                        ViewOnly = "Yes";
                                        break;

                                    case "Contribute Limited":
                                        ContributeLimited = "Yes";
                                        break;

                                    case "Approve":
                                        Approve = "Yes";
                                        break;

                                    default:
                                        OtherPermissions = rdef.Name;
                                        break;
                                }
                            }

                            #endregion

                            try
                            {
                                clientcontext.Load(rAssignment.Member);
                                clientcontext.ExecuteQuery();

                                if (rAssignment.Member.Title.Contains("c:0u.c|tenant|"))
                                {
                                    continue;
                                }

                                if (rAssignment.Member.PrincipalType == PrincipalType.SharePointGroup)
                                {
                                    Microsoft.SharePoint.Client.Group ouserGroup = (Microsoft.SharePoint.Client.Group)rAssignment.Member.TypedObject;
                                    clientcontext.Load(ouserGroup);
                                    clientcontext.ExecuteQuery();

                                    UserCollection userColl = ouserGroup.Users;
                                    clientcontext.Load(userColl);
                                    clientcontext.ExecuteQuery();

                                    foreach (User xUser in userColl)
                                    {
                                        if (xUser.PrincipalType == PrincipalType.SecurityGroup)
                                        {
                                            //AdGroupsinFileGroup += ouserGroup.Title + ";";
                                            //filehasSecurityGroup = true;
                                            if (lstADGroupsColl.Contains(xUser.Title.ToString().Trim().ToLower()))
                                            {
                                                if (((!string.IsNullOrEmpty(Design)) || (!string.IsNullOrEmpty(Contribute)) || (!string.IsNullOrEmpty(Read)) || (!string.IsNullOrEmpty(FullControl)) || (!string.IsNullOrEmpty(Edit)) || (!string.IsNullOrEmpty(ViewOnly)) || (!string.IsNullOrEmpty(Approve)) || (!string.IsNullOrEmpty(ContributeLimited))) || ((!string.IsNullOrEmpty(OtherPermissions)) && (OtherPermissions != "Limited Access")))
                                                {
                                                    if (xUser.Title.ToString().ToLower().Contains("everyone"))
                                                    {
                                                        //if (!BuiltinGroups.Contains(xUser.Title))
                                                        //{
                                                        //    BuiltinGroups.Add(xUser.Title);
                                                        //}

                                                        if (BuiltinGroups.ContainsKey(xUser.Title))
                                                        {
                                                            BuiltinGroups[xUser.Title]++;
                                                        }
                                                        else
                                                        {
                                                            BuiltinGroups.Add(xUser.Title, 1);
                                                        }
                                                    }
                                                    else
                                                    {
                                                        //if (!ADGroups.Contains(xUser.Title))
                                                        //{
                                                        //    ADGroups.Add(xUser.Title);
                                                        //}

                                                        if (ADGroups.ContainsKey(xUser.Title))
                                                        {
                                                            ADGroups[xUser.Title]++;
                                                        }
                                                        else
                                                        {
                                                            ADGroups.Add(xUser.Title, 1);
                                                        }
                                                    }

                                                    excelWriterScoringNew.WriteLine("\"" + "File" + "\"" + "," + "\"" + clientcontext.Web.Url.Replace(clientcontext.Web.ServerRelativeUrl, "") + sFile.ServerRelativeUrl + "\"" + "," + "\"" + xUser.Title + "\"" + "," + "\"" + ouserGroup.Title + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + Design + "\"" + "," + "\"" + Contribute + "\"" + "," + "\"" + Read + "\"" + "," + "\"" + FullControl + "\"" + "," + "\"" + Edit + "\"" + "," + "\"" + ViewOnly + "\"" + "," + "\"" + Approve + "\"" + "," + "\"" + ContributeLimited + "\"" + "," + "\"" + OtherPermissions + "\"");
                                                    excelWriterScoringNew.Flush();
                                                }
                                            }
                                        }
                                    }
                                }
                                if (rAssignment.Member.PrincipalType == PrincipalType.SecurityGroup)
                                {
                                    //AdGroupsatFile += rAssignment.Member.Title + ";";
                                    //filehasSecurityGroup = true;
                                    if (lstADGroupsColl.Contains(rAssignment.Member.Title.ToString().Trim().ToLower()))
                                    {
                                        if (((!string.IsNullOrEmpty(Design)) || (!string.IsNullOrEmpty(Contribute)) || (!string.IsNullOrEmpty(Read)) || (!string.IsNullOrEmpty(FullControl)) || (!string.IsNullOrEmpty(Edit)) || (!string.IsNullOrEmpty(ViewOnly)) || (!string.IsNullOrEmpty(Approve)) || (!string.IsNullOrEmpty(ContributeLimited))) || ((!string.IsNullOrEmpty(OtherPermissions)) && (OtherPermissions != "Limited Access")))
                                        {
                                            if (rAssignment.Member.Title.ToString().ToLower().Contains("everyone"))
                                            {
                                                //if (!BuiltinGroups.Contains(rAssignment.Member.Title))
                                                //{
                                                //    BuiltinGroups.Add(rAssignment.Member.Title);
                                                //}

                                                if (BuiltinGroups.ContainsKey(rAssignment.Member.Title))
                                                {
                                                    BuiltinGroups[rAssignment.Member.Title]++;
                                                }
                                                else
                                                {
                                                    BuiltinGroups.Add(rAssignment.Member.Title, 1);
                                                }
                                            }
                                            else
                                            {
                                                //if (!ADGroups.Contains(rAssignment.Member.Title))
                                                //{
                                                //    ADGroups.Add(rAssignment.Member.Title);
                                                //}

                                                if (ADGroups.ContainsKey(rAssignment.Member.Title))
                                                {
                                                    ADGroups[rAssignment.Member.Title]++;
                                                }
                                                else
                                                {
                                                    ADGroups.Add(rAssignment.Member.Title, 1);
                                                }
                                            }

                                            excelWriterScoringNew.WriteLine("\"" + "File" + "\"" + "," + "\"" + clientcontext.Web.Url.Replace(clientcontext.Web.ServerRelativeUrl, "") + sFile.ServerRelativeUrl + "\"" + "," + "\"" + rAssignment.Member.Title + "\"" + "," + "\"" + "Directly Assigned" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + Design + "\"" + "," + "\"" + Contribute + "\"" + "," + "\"" + Read + "\"" + "," + "\"" + FullControl + "\"" + "," + "\"" + Edit + "\"" + "," + "\"" + ViewOnly + "\"" + "," + "\"" + Approve + "\"" + "," + "\"" + ContributeLimited + "\"" + "," + "\"" + OtherPermissions + "\"");
                                            excelWriterScoringNew.Flush();
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                continue;
                            }
                        }

                        //if (filehasSecurityGroup)
                        //{
                        //    excelWriterScoringNew.WriteLine("\"" + "File" + "\"" + "," + "\"" + clientcontext.Web.Url.Replace(clientcontext.Web.ServerRelativeUrl, "") + sFile.ServerRelativeUrl + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + AdGroupsatFile + "\"" + "," + "\"" + AdGroupsinFileGroup + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"");
                        //    excelWriterScoringNew.Flush();
                        //}

                        #endregion

                    }
                }

                clientcontext.Load(Fld.Folders);
                clientcontext.ExecuteQuery();

                foreach (Folder folder in Fld.Folders)
                {
                    clientcontext.Load(folder);
                    clientcontext.ExecuteQuery();

                    if (folder.Name != "Forms")
                    {
                        GetCounts(folder, clientcontext, excelWriterScoringNew);
                    }
                }
            }
            catch (Exception ex)
            {

            }

            #region Extra Recurssive

            //foreach (Folder folder in Fld.Folders)
            //{
            //    clientcontext.Load(folder.Files);
            //    clientcontext.ExecuteQuery();

            //    FileCount += folder.Files.Count;

            //    clientcontext.Load(folder.Folders);
            //    clientcontext.ExecuteQuery();

            //    foreach (Folder subsubfolder in folder.Folders)
            //    {
            //        clientcontext.Load(subsubfolder);
            //        clientcontext.ExecuteQuery();

            //        if (subsubfolder.Name != "Forms")
            //        {
            //            FoldCount++;
            //        }
            //    }

            //    if (folder.Folders.Count > 0)
            //    {
            //        foreach (Folder subFolder in folder.Folders)
            //        {
            //            GetCounts(subFolder, clientcontext);
            //        }
            //    }
            //} 

            #endregion
        }
        public void GetCountsatListLevel(Folder Fld, ClientContext clientcontext)
        {
            clientcontext.Load(Fld);
            clientcontext.ExecuteQuery();

            clientcontext.Load(Fld.Files);
            clientcontext.ExecuteQuery();

            ListFileCount += Fld.Files.Count;

            clientcontext.Load(Fld.Folders);
            clientcontext.ExecuteQuery();

            foreach (Folder folder in Fld.Folders)
            {
                clientcontext.Load(folder);
                clientcontext.ExecuteQuery();

                if (folder.Name != "Forms")
                {
                    ListFoldCount++;
                }
            }

            clientcontext.Load(Fld.Folders);
            clientcontext.ExecuteQuery();

            foreach (Folder folder in Fld.Folders)
            {
                clientcontext.Load(folder);
                clientcontext.ExecuteQuery();

                if (folder.Name != "Forms")
                {
                    GetCountsatListLevel(folder, clientcontext);
                }
            }
        }
        public void getWeb(string siteURL, StreamWriter excelWriterScoringNew)
        {
            try
            {
                AuthenticationManager authManager = new AuthenticationManager();

                //using (var clientcontextSub = authManager.GetSharePointOnlineAuthenticatedContextTenant(siteURL, "adam.a@VerinonTechnology.onmicrosoft.com", "Lot62215##"))
                using (var clientcontextSub = authManager.GetSharePointOnlineAuthenticatedContextTenant(siteURL, "spo.admin.verinon@agilent.onmicrosoft.com", "Lot62215"))
                {
                    #region SubSite                 

                    try
                    {
                        RoleAssignmentCollection webRoleAssignments = null;
                        string webUniqueRoles = string.Empty;

                        Web oWeb = clientcontextSub.Web;
                        clientcontextSub.Load(oWeb);
                        clientcontextSub.Load(oWeb, li => li.HasUniqueRoleAssignments);
                        clientcontextSub.ExecuteQuery();

                        //if (oWeb.HasUniqueRoleAssignments)
                        //{
                        //    webUniqueRoles = "Unique Permissions";
                        //}
                        //else
                        //{
                        //    webUniqueRoles = "Inherit from Parent";
                        //}

                        if (oWeb.HasUniqueRoleAssignments)
                        {
                            bool foundatSiteLevel = false;

                            string AdGroupsinGroupWeb = string.Empty;
                            string AdGroupsatSiteWeb = string.Empty;

                            #region Subsite Permission check

                            webRoleAssignments = clientcontextSub.Web.RoleAssignments;
                            clientcontextSub.Load(webRoleAssignments);
                            clientcontextSub.ExecuteQuery();

                            foreach (RoleAssignment member1 in webRoleAssignments)
                            {

                                #region Role Definations

                                RoleDefinitionBindingCollection rdefColl = member1.RoleDefinitionBindings;
                                clientcontextSub.Load(rdefColl);
                                clientcontextSub.ExecuteQuery();

                                string Design = string.Empty;
                                string Contribute = string.Empty;
                                string Read = string.Empty;
                                string FullControl = string.Empty;
                                string Edit = string.Empty;
                                string ViewOnly = string.Empty;
                                string Approve = string.Empty;
                                string ContributeLimited = string.Empty;
                                string OtherPermissions = string.Empty;

                                foreach (RoleDefinition rdef in rdefColl)
                                {
                                    clientcontextSub.Load(rdef);
                                    clientcontextSub.ExecuteQuery();

                                    switch (rdef.Name)
                                    {
                                        case "Design":
                                            Design = "Yes";
                                            break;

                                        case "Contribute":
                                            Contribute = "Yes";
                                            break;

                                        case "Read":
                                            Read = "Yes";
                                            break;

                                        case "Full Control":
                                            FullControl = "Yes";
                                            break;

                                        case "Edit":
                                            Edit = "Yes";
                                            break;

                                        case "View Only":
                                            ViewOnly = "Yes";
                                            break;

                                        case "Contribute Limited":
                                            ContributeLimited = "Yes";
                                            break;

                                        case "Approve":
                                            Approve = "Yes";
                                            break;

                                        default:
                                            OtherPermissions = rdef.Name;
                                            break;
                                    }
                                }

                                #endregion

                                try
                                {
                                    //if (!foundatSiteLevel)
                                    //{
                                    clientcontextSub.Load(member1.Member);
                                    clientcontextSub.ExecuteQuery();

                                    if (member1.Member.Title.Contains("c:0u.c|tenant|"))
                                    {
                                        continue;
                                    }

                                    if (member1.Member.PrincipalType == PrincipalType.SharePointGroup)
                                    {
                                        Microsoft.SharePoint.Client.Group ouserGroup = (Microsoft.SharePoint.Client.Group)member1.Member.TypedObject;
                                        clientcontextSub.Load(ouserGroup);
                                        clientcontextSub.ExecuteQuery();

                                        UserCollection userColl = ouserGroup.Users;
                                        clientcontextSub.Load(userColl);
                                        clientcontextSub.ExecuteQuery();

                                        foreach (User xUser in userColl)
                                        {
                                            if (xUser.PrincipalType == PrincipalType.SecurityGroup)
                                            {
                                                //if (xUser.Title == "Everyone except external users")
                                                //{
                                                //AdGroupsinGroupWeb += ouserGroup.Title + ";";
                                                //foundatSiteLevel = true;
                                                //break;
                                                if (lstADGroupsColl.Contains(xUser.Title.ToString().Trim().ToLower()))
                                                {
                                                    if (((!string.IsNullOrEmpty(Design)) || (!string.IsNullOrEmpty(Contribute)) || (!string.IsNullOrEmpty(Read)) || (!string.IsNullOrEmpty(FullControl)) || (!string.IsNullOrEmpty(Edit)) || (!string.IsNullOrEmpty(ViewOnly)) || (!string.IsNullOrEmpty(Approve)) || (!string.IsNullOrEmpty(ContributeLimited))) || ((!string.IsNullOrEmpty(OtherPermissions)) && (OtherPermissions != "Limited Access")))
                                                    {
                                                        if (xUser.Title.ToString().ToLower().Contains("everyone"))
                                                        {
                                                            //if (!BuiltinGroups.Contains(xUser.Title))
                                                            //{
                                                            //    BuiltinGroups.Add(xUser.Title);
                                                            //}

                                                            if (BuiltinGroups.ContainsKey(xUser.Title))
                                                            {
                                                                BuiltinGroups[xUser.Title]++;
                                                            }
                                                            else
                                                            {
                                                                BuiltinGroups.Add(xUser.Title, 1);
                                                            }
                                                        }
                                                        else
                                                        {
                                                            //if (!ADGroups.Contains(xUser.Title))
                                                            //{
                                                            //    ADGroups.Add(xUser.Title);
                                                            //}

                                                            if (ADGroups.ContainsKey(xUser.Title))
                                                            {
                                                                ADGroups[xUser.Title]++;
                                                            }
                                                            else
                                                            {
                                                                ADGroups.Add(xUser.Title, 1);
                                                            }

                                                        }

                                                        excelWriterScoringNew.WriteLine("\"" + "SubSite" + "\"" + "," + "\"" + clientcontextSub.Web.Url.ToString() + "\"" + "," + "\"" + xUser.Title + "\"" + "," + "\"" + ouserGroup.Title + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + Design + "\"" + "," + "\"" + Contribute + "\"" + "," + "\"" + Read + "\"" + "," + "\"" + FullControl + "\"" + "," + "\"" + Edit + "\"" + "," + "\"" + ViewOnly + "\"" + "," + "\"" + Approve + "\"" + "," + "\"" + ContributeLimited + "\"" + "," + "\"" + OtherPermissions + "\"");
                                                        excelWriterScoringNew.Flush();
                                                    }
                                                    //excelWriterScoringNew.WriteLine("\"" + "SubSite" + "\"" + "," + "\"" + clientcontextSub.Web.Url.ToString() + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + "Yes" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"");

                                                    //foundatSiteLevel = true;
                                                    //break;
                                                    //}
                                                }
                                            }
                                        }
                                    }
                                    if (member1.Member.PrincipalType == PrincipalType.SecurityGroup)
                                    {
                                        //if (member1.Member.Title == "Everyone except external users")
                                        //{
                                        //excelWriterScoringNew.WriteLine("\"" + "SubSite" + "\"" + "," + "\"" + clientcontextSub.Web.Url.ToString() + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + "Yes" + "\"" + "," + "\"" + webUniqueRoles + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"");
                                        //AdGroupsatSiteWeb += member1.Member.Title + ";";
                                        //foundatSiteLevel = true;
                                        if (lstADGroupsColl.Contains(member1.Member.Title.ToString().Trim().ToLower()))
                                        {
                                            if (((!string.IsNullOrEmpty(Design)) || (!string.IsNullOrEmpty(Contribute)) || (!string.IsNullOrEmpty(Read)) || (!string.IsNullOrEmpty(FullControl)) || (!string.IsNullOrEmpty(Edit)) || (!string.IsNullOrEmpty(ViewOnly)) || (!string.IsNullOrEmpty(Approve)) || (!string.IsNullOrEmpty(ContributeLimited))) || ((!string.IsNullOrEmpty(OtherPermissions)) && (OtherPermissions != "Limited Access")))
                                            {
                                                if (member1.Member.Title.ToString().ToLower().Contains("everyone"))
                                                {
                                                    //if (!BuiltinGroups.Contains(member1.Member.Title))
                                                    //{
                                                    //    BuiltinGroups.Add(member1.Member.Title);
                                                    //}

                                                    if (BuiltinGroups.ContainsKey(member1.Member.Title))
                                                    {
                                                        BuiltinGroups[member1.Member.Title]++;
                                                    }
                                                    else
                                                    {
                                                        BuiltinGroups.Add(member1.Member.Title, 1);
                                                    }
                                                }
                                                else
                                                {
                                                    //if (!ADGroups.Contains(member1.Member.Title))
                                                    //{
                                                    //    ADGroups.Add(member1.Member.Title);
                                                    //}

                                                    if (ADGroups.ContainsKey(member1.Member.Title))
                                                    {
                                                        ADGroups[member1.Member.Title]++;
                                                    }
                                                    else
                                                    {
                                                        ADGroups.Add(member1.Member.Title, 1);
                                                    }
                                                }

                                                excelWriterScoringNew.WriteLine("\"" + "SubSite" + "\"" + "," + "\"" + clientcontextSub.Web.Url.ToString() + "\"" + "," + "\"" + member1.Member.Title + "\"" + "," + "\"" + "Directly Assigned" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + Design + "\"" + "," + "\"" + Contribute + "\"" + "," + "\"" + Read + "\"" + "," + "\"" + FullControl + "\"" + "," + "\"" + Edit + "\"" + "," + "\"" + ViewOnly + "\"" + "," + "\"" + Approve + "\"" + "," + "\"" + ContributeLimited + "\"" + "," + "\"" + OtherPermissions + "\"");
                                                excelWriterScoringNew.Flush();
                                            }
                                        }
                                        //foundatSiteLevel = true;
                                        //break;
                                        //}
                                    }
                                    //}
                                    //else
                                    //{
                                    //    break;
                                    //}
                                }
                                catch (Exception ex)
                                {
                                    continue;
                                }
                            }

                            //if (foundatSiteLevel)
                            //{
                            //    excelWriterScoringNew.WriteLine("\"" + "Subsite" + "\"" + "," + "\"" + clientcontextSub.Url + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + AdGroupsatSiteWeb + "\"" + "," + "\"" + AdGroupsinGroupWeb + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"");
                            //    excelWriterScoringNew.Flush();
                            //}

                            #endregion
                        }
                    }
                    catch (Exception ex)
                    {

                    }

                    #endregion

                    #region Lists on Subsites

                    ListCollection olistColl = clientcontextSub.Web.Lists;
                    clientcontextSub.Load(olistColl);
                    clientcontextSub.ExecuteQuery();

                    foreach (List oList in olistColl)
                    {
                        string listUniqueRoles = string.Empty;

                        clientcontextSub.Load(oList);
                        clientcontextSub.Load(oList, li => li.HasUniqueRoleAssignments);
                        clientcontextSub.ExecuteQuery();

                        string AdGroupsinSubsiteListGroup = string.Empty;
                        string AdGroupsatSubsiteListSite = string.Empty;

                        if (oList.BaseType == BaseType.DocumentLibrary)
                        {
                            if (oList.Title != "Form Templates" && oList.Title != "Site Assets" && oList.Title != "SitePages" && oList.Title != "Style Library" && oList.Hidden == false && oList.IsCatalog == false && oList.BaseTemplate == 101)
                            {
                                if (oList.HasUniqueRoleAssignments)
                                {
                                    ListFoldCount = 0;
                                    ListFileCount = 0;

                                    GetCountsatListLevel(oList.RootFolder, clientcontextSub);

                                    bool foundatListLevel = false;

                                    #region Subsite Lists Permission check

                                    RoleAssignmentCollection roles = oList.RoleAssignments;
                                    clientcontextSub.Load(roles);
                                    clientcontextSub.ExecuteQuery();

                                    Web oWebx = clientcontextSub.Web;
                                    clientcontextSub.Load(oWebx);
                                    clientcontextSub.ExecuteQuery();

                                    //Get all the RoleAssignments for this document
                                    foreach (RoleAssignment rAssignment in roles)
                                    {

                                        #region Role Definations

                                        RoleDefinitionBindingCollection rdefColl = rAssignment.RoleDefinitionBindings;
                                        clientcontextSub.Load(rdefColl);
                                        clientcontextSub.ExecuteQuery();

                                        string Design = string.Empty;
                                        string Contribute = string.Empty;
                                        string Read = string.Empty;
                                        string FullControl = string.Empty;
                                        string Edit = string.Empty;
                                        string ViewOnly = string.Empty;
                                        string Approve = string.Empty;
                                        string ContributeLimited = string.Empty;
                                        string OtherPermissions = string.Empty;

                                        foreach (RoleDefinition rdef in rdefColl)
                                        {
                                            clientcontextSub.Load(rdef);
                                            clientcontextSub.ExecuteQuery();

                                            switch (rdef.Name)
                                            {
                                                case "Design":
                                                    Design = "Yes";
                                                    break;

                                                case "Contribute":
                                                    Contribute = "Yes";
                                                    break;

                                                case "Read":
                                                    Read = "Yes";
                                                    break;

                                                case "Full Control":
                                                    FullControl = "Yes";
                                                    break;

                                                case "Edit":
                                                    Edit = "Yes";
                                                    break;

                                                case "View Only":
                                                    ViewOnly = "Yes";
                                                    break;

                                                case "Contribute Limited":
                                                    ContributeLimited = "Yes";
                                                    break;

                                                case "Approve":
                                                    Approve = "Yes";
                                                    break;

                                                default:
                                                    OtherPermissions = rdef.Name;
                                                    break;
                                            }
                                        }

                                        #endregion

                                        try
                                        {
                                            //if (!foundatListLevel)
                                            //{
                                            clientcontextSub.Load(rAssignment.Member);
                                            clientcontextSub.ExecuteQuery();

                                            if (rAssignment.Member.Title.Contains("c:0u.c|tenant|"))
                                            {
                                                continue;
                                            }

                                            if (rAssignment.Member.PrincipalType == PrincipalType.SharePointGroup)
                                            {
                                                Microsoft.SharePoint.Client.Group ouserGroup = (Microsoft.SharePoint.Client.Group)rAssignment.Member.TypedObject;
                                                clientcontextSub.Load(ouserGroup);
                                                clientcontextSub.ExecuteQuery();

                                                UserCollection userColl = ouserGroup.Users;
                                                clientcontextSub.Load(userColl);
                                                clientcontextSub.ExecuteQuery();

                                                foreach (User xUser in userColl)
                                                {
                                                    if (xUser.PrincipalType == PrincipalType.SecurityGroup)
                                                    {
                                                        //clientcontextSub.Load(oList.RootFolder);
                                                        //clientcontextSub.ExecuteQuery();

                                                        //AdGroupsinSubsiteListGroup += ouserGroup.Title + ";";
                                                        //foundatListLevel = true;
                                                        //break;

                                                        if (lstADGroupsColl.Contains(xUser.Title.ToString().Trim().ToLower()))
                                                        {
                                                            if (((!string.IsNullOrEmpty(Design)) || (!string.IsNullOrEmpty(Contribute)) || (!string.IsNullOrEmpty(Read)) || (!string.IsNullOrEmpty(FullControl)) || (!string.IsNullOrEmpty(Edit)) || (!string.IsNullOrEmpty(ViewOnly)) || (!string.IsNullOrEmpty(Approve)) || (!string.IsNullOrEmpty(ContributeLimited))) || ((!string.IsNullOrEmpty(OtherPermissions)) && (OtherPermissions != "Limited Access")))
                                                            {
                                                                if (xUser.Title.ToString().ToLower().Contains("everyone"))
                                                                {
                                                                    //if (!BuiltinGroups.Contains(xUser.Title))
                                                                    //{
                                                                    //    BuiltinGroups.Add(xUser.Title);
                                                                    //}

                                                                    if (BuiltinGroups.ContainsKey(xUser.Title))
                                                                    {
                                                                        BuiltinGroups[xUser.Title]++;
                                                                    }
                                                                    else
                                                                    {
                                                                        BuiltinGroups.Add(xUser.Title, 1);
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    //if (!ADGroups.Contains(xUser.Title))
                                                                    //{
                                                                    //    ADGroups.Add(xUser.Title);
                                                                    //}

                                                                    if (ADGroups.ContainsKey(xUser.Title))
                                                                    {
                                                                        ADGroups[xUser.Title]++;
                                                                    }
                                                                    else
                                                                    {
                                                                        ADGroups.Add(xUser.Title, 1);
                                                                    }

                                                                }

                                                                excelWriterScoringNew.WriteLine("\"" + "Doc Library" + "\"" + "," + "\"" + oWebx.Url.Replace(oWebx.ServerRelativeUrl, "") + oList.RootFolder.ServerRelativeUrl + "\"" + "," + "\"" + xUser.Title + "\"" + "," + "\"" + ouserGroup.Title + "\"" + "," + "\"" + ListFoldCount.ToString() + "\"" + "," + "\"" + ListFileCount.ToString() + "\"" + "," + "\"" + Design + "\"" + "," + "\"" + Contribute + "\"" + "," + "\"" + Read + "\"" + "," + "\"" + FullControl + "\"" + "," + "\"" + Edit + "\"" + "," + "\"" + ViewOnly + "\"" + "," + "\"" + Approve + "\"" + "," + "\"" + ContributeLimited + "\"" + "," + "\"" + OtherPermissions + "\"");
                                                                excelWriterScoringNew.Flush();
                                                            }
                                                        }
                                                        //excelWriterScoringNew.WriteLine("\"" + "Doc Library" + "\"" + "," + "\"" + oWebx.Url.Replace(oWebx.ServerRelativeUrl, "") + oList.RootFolder.ServerRelativeUrl + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + "Yes" + "\"" + "," + "\"" + listUniqueRoles + "\"" + "," + "\"" + FoldCount.ToString() + "\"" + "," + "\"" + FileCount.ToString() + "\"");

                                                        //foundatListLevel = true;
                                                        //break;
                                                    }
                                                }
                                            }
                                            if (rAssignment.Member.PrincipalType == PrincipalType.SecurityGroup)
                                            {
                                                //if (rAssignment.Member.Title == "Everyone except external users")
                                                //{
                                                //clientcontextSub.Load(oList.RootFolder);
                                                //clientcontextSub.ExecuteQuery();

                                                //AdGroupsatSubsiteListSite += rAssignment.Member.Title + ";";
                                                //foundatListLevel = true;
                                                if (lstADGroupsColl.Contains(rAssignment.Member.Title.ToString().Trim().ToLower()))
                                                {
                                                    if (((!string.IsNullOrEmpty(Design)) || (!string.IsNullOrEmpty(Contribute)) || (!string.IsNullOrEmpty(Read)) || (!string.IsNullOrEmpty(FullControl)) || (!string.IsNullOrEmpty(Edit)) || (!string.IsNullOrEmpty(ViewOnly)) || (!string.IsNullOrEmpty(Approve)) || (!string.IsNullOrEmpty(ContributeLimited))) || ((!string.IsNullOrEmpty(OtherPermissions)) && (OtherPermissions != "Limited Access")))
                                                    {
                                                        if (rAssignment.Member.Title.ToString().ToLower().Contains("everyone"))
                                                        {
                                                            //if (!BuiltinGroups.Contains(rAssignment.Member.Title))
                                                            //{
                                                            //    BuiltinGroups.Add(rAssignment.Member.Title);
                                                            //}

                                                            if (BuiltinGroups.ContainsKey(rAssignment.Member.Title))
                                                            {
                                                                BuiltinGroups[rAssignment.Member.Title]++;
                                                            }
                                                            else
                                                            {
                                                                BuiltinGroups.Add(rAssignment.Member.Title, 1);
                                                            }
                                                        }
                                                        else
                                                        {
                                                            //if (!ADGroups.Contains(rAssignment.Member.Title))
                                                            //{
                                                            //    ADGroups.Add(rAssignment.Member.Title);
                                                            //}

                                                            if (ADGroups.ContainsKey(rAssignment.Member.Title))
                                                            {
                                                                ADGroups[rAssignment.Member.Title]++;
                                                            }
                                                            else
                                                            {
                                                                ADGroups.Add(rAssignment.Member.Title, 1);
                                                            }
                                                        }

                                                        excelWriterScoringNew.WriteLine("\"" + "Doc Library" + "\"" + "," + "\"" + oWebx.Url.Replace(oWebx.ServerRelativeUrl, "") + oList.RootFolder.ServerRelativeUrl + "\"" + "," + "\"" + rAssignment.Member.Title + "\"" + "," + "\"" + "Directly Assigned" + "\"" + "," + "\"" + ListFoldCount.ToString() + "\"" + "," + "\"" + ListFileCount.ToString() + "\"" + "," + "\"" + Design + "\"" + "," + "\"" + Contribute + "\"" + "," + "\"" + Read + "\"" + "," + "\"" + FullControl + "\"" + "," + "\"" + Edit + "\"" + "," + "\"" + ViewOnly + "\"" + "," + "\"" + Approve + "\"" + "," + "\"" + ContributeLimited + "\"" + "," + "\"" + OtherPermissions + "\"");
                                                        excelWriterScoringNew.Flush();
                                                    }
                                                }

                                                //excelWriterScoringNew.WriteLine("\"" + "Doc Library" + "\"" + "," + "\"" + oWebx.Url.Replace(oWebx.ServerRelativeUrl, "") + oList.RootFolder.ServerRelativeUrl + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + "Yes" + "\"" + "," + "\"" + listUniqueRoles + "\"" + "," + "\"" + FoldCount.ToString() + "\"" + "," + "\"" + FileCount.ToString() + "\"");

                                                //foundatListLevel = true;
                                                //break;
                                                //}
                                            }
                                            //}
                                            //else
                                            //{
                                            //    break;
                                            //}
                                        }
                                        catch (Exception ex)
                                        {
                                            continue;
                                        }
                                    }

                                    //if (foundatListLevel)
                                    //{
                                    //    ListFoldCount = 0;
                                    //    ListFileCount = 0;

                                    //    GetCountsatListLevel(oList.RootFolder, clientcontextSub);

                                    //    excelWriterScoringNew.WriteLine("\"" + "Doc Library" + "\"" + "," + "\"" + oWebx.Url.Replace(oWebx.ServerRelativeUrl, "") + oList.RootFolder.ServerRelativeUrl + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + AdGroupsatSubsiteListSite + "\"" + "," + "\"" + AdGroupsinSubsiteListGroup + "\"" + "," + "\"" + ListFoldCount.ToString() + "\"" + "," + "\"" + ListFileCount.ToString() + "\"");
                                    //    excelWriterScoringNew.Flush();
                                    //}

                                    #endregion
                                }

                                clientcontextSub.Load(oList.RootFolder.Folders);
                                clientcontextSub.ExecuteQuery();

                                foreach (Folder sFolder in oList.RootFolder.Folders)
                                {
                                    clientcontextSub.Load(sFolder);
                                    clientcontextSub.ExecuteQuery();

                                    if (sFolder.Name != "Forms")
                                    {
                                        GetCounts(sFolder, clientcontextSub, excelWriterScoringNew);
                                    }
                                }
                            }
                        }
                    }

                    #endregion

                    #region Recursive

                    WebCollection oWebs = clientcontextSub.Web.Webs;
                    clientcontextSub.Load(oWebs);
                    clientcontextSub.ExecuteQuery();

                    foreach (Web oWeb in oWebs)
                    {
                        clientcontextSub.Load(oWeb);
                        clientcontextSub.ExecuteQuery();

                        getWeb(oWeb.Url, excelWriterScoringNew);
                    }

                    #endregion
                }
            }
            catch (Exception ex)
            {

            }
        }


        public static ClientContext createContext(string sitCollURL)
        {
            siteTitle = string.Empty;

            ClientContext contxt = new ClientContext(sitCollURL);
            SecureString secureStrPwd = new SecureString();

            //foreach (char x in "verinon@123".ToString())//need to change according to admin user credentials
            //{
            //    secureStrPwd.AppendChar(x);
            //}

            //SharePointOnlineCredentials credentials = new SharePointOnlineCredentials("spo.admin.verinon@agilent.onmicrosoft.com", secureStrPwd);

            foreach (char x in "Lot62215##".ToString())//need to change according to admin user credentials
            {
                secureStrPwd.AppendChar(x);
            }

            SharePointOnlineCredentials credentials = new SharePointOnlineCredentials("adam.a@VerinonTechnology.onmicrosoft.com", secureStrPwd);

            contxt.Credentials = credentials;
            contxt.ExecuteQuery();
            contxt.RequestTimeout = -1;

            web = contxt.Web;
            contxt.Load(web);
            contxt.ExecuteQuery();
            try
            {
                siteTitle = web.Title;
            }
            catch
            {

            }
            return contxt;
        }
        public void WriteToErrorLog(string msg, string stkTrace, string title)
        {
            //log it
            FileStream fs1 = new FileStream("errorlog.txt", FileMode.Append, FileAccess.Write);
            StreamWriter s1 = new StreamWriter(fs1);
            //s1.WriteLine("Title: " + title);
            s1.WriteLine("Message: " + msg);
            s1.WriteLine("StackTrace: " + stkTrace);
            s1.WriteLine("Title: " + title);
            s1.WriteLine("Date/Time: " + System.DateTime.Now.ToString());
            s1.WriteLine("===========================================================================================");
            s1.Close();
            fs1.Close();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                textBox3.Text = openFileDialog2.FileName;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {


            #region AD Groups CSV Reading

            //lstADGroupsColl.Clear();

            //if (!string.IsNullOrEmpty(textBox3.Text))
            //{
            //    StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox3.Text));

            //    while (!sr.EndOfStream)
            //    {
            //        try
            //        {
            //            lstADGroupsColl.Add(sr.ReadLine().Trim().ToLower());
            //        }
            //        catch
            //        {
            //            continue;
            //        }
            //    }
            //}
            //else
            //{
            //    MessageBox.Show("Please browse the path for ADGroups.csv");
            //}

            #endregion

            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();

            //if (!string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrEmpty(textBox2.Text))
            {
                StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

                while (!sr.EndOfStream)
                {
                    try
                    {
                        lstSiteColl.Add(sr.ReadLine().Trim());
                    }
                    catch
                    {
                        continue;
                    }
                }
            }

            //else
            //{
            //    MessageBox.Show("Please browse the path for SiteColl.csv / Reports folder");
            //}

            #endregion

            //StreamWriter excelWriterScoringMatrixNew = null;

            //excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "ScoringMatrix" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");

            //excelWriterScoringMatrixNew.WriteLine("Filename" + "," + "URL" + "," + "Owners" + "," + "Built-in-Groups" + "," + "AD Groups" + "," + "Start Time" + "," + "End Date" + "," + "Remarks");
            //excelWriterScoringMatrixNew.Flush();            

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();
                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString().Trim(), "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        clientcontext.Load(clientcontext.Web);
                        clientcontext.ExecuteQuery();

                        ListCollection _Lists = clientcontext.Web.Lists;
                        clientcontext.Load(_Lists);
                        clientcontext.ExecuteQuery();

                        foreach (List Pagelist in _Lists)
                        {
                            try
                            {


                                clientcontext.Load(Pagelist);
                                clientcontext.ExecuteQuery();

                                bool _dListExist = clientcontext.Web.Lists.Cast<List>().Any(list => string.Equals(list.Title, Pagelist.Title));

                                if (_dListExist)
                                {
                                    //try
                                    //{
                                    //    if (Pagelist.Title.ToLower() == "1_Uploaded Files".ToLower() || Pagelist.Title.ToLower() == "Site Assets".ToLower())
                                    //    {
                                    //        ContentTypeCollection contentTypesFields = Pagelist.ContentTypes;
                                    //        clientcontext.Load(contentTypesFields);
                                    //        clientcontext.ExecuteQuery();

                                    //        clientcontext.Load(Pagelist, l => l.ContentTypes, l => l.RootFolder.UniqueContentTypeOrder);
                                    //        clientcontext.ExecuteQuery();
                                    //        var contentTypeOrder = (from ct in contentTypesFields where ct.Name == "Document" select ct.Id).ToList();
                                    //        Pagelist.RootFolder.UniqueContentTypeOrder = contentTypeOrder;
                                    //        Pagelist.RootFolder.Update();
                                    //        clientcontext.ExecuteQuery();
                                    //    }
                                    //}catch(Exception ex)
                                    //{ }

                                    FieldCollection collFields = Pagelist.Fields;
                                    clientcontext.Load(collFields);
                                    clientcontext.ExecuteQuery();

                                    bool CategoriesAlreadyCreated = false;

                                    foreach (Field spField in collFields)
                                    {
                                        if (spField.StaticName.ToLower() == "Categorization".ToLower())
                                        {
                                            Field fld5 = collFields.GetByInternalNameOrTitle("Categorization");
                                            clientcontext.Load(fld5);
                                            clientcontext.ExecuteQuery();

                                            fld5.DeleteObject();
                                            clientcontext.ExecuteQuery();

                                            CategoriesAlreadyCreated = true;

                                            break;
                                        }
                                    }

                                    if (CategoriesAlreadyCreated)
                                    {
                                        try
                                        {

                                            List list = clientcontext.Web.Lists.GetByTitle("Manage Categories"); //Categories  //CategoriesList
                                            clientcontext.Load(list);
                                            clientcontext.ExecuteQuery();
                                            string schemaLookupField = "<Field Type='LookupMulti' Name='Categorization' StaticName='Categorization' DisplayName='Categorization' List = '" + list.Id + "' ShowField = 'Title' Mult = 'TRUE'/>";
                                            Field lookupField = Pagelist.Fields.AddFieldAsXml(schemaLookupField, true, AddFieldOptions.AddFieldInternalNameHint);
                                            Pagelist.Update();
                                            clientcontext.ExecuteQuery();

                                        }
                                        catch (Exception es)
                                        {
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                continue;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }
            this.Text = "Completed..";
        }

        private void button6_Click(object sender, EventArgs e)
        {
            #region AD Groups CSV Reading

            //lstADGroupsColl.Clear();

            //if (!string.IsNullOrEmpty(textBox3.Text))
            //{
            //    StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox3.Text));

            //    while (!sr.EndOfStream)
            //    {
            //        try
            //        {
            //            lstADGroupsColl.Add(sr.ReadLine().Trim().ToLower());
            //        }
            //        catch
            //        {
            //            continue;
            //        }
            //    }
            //}
            //else
            //{
            //    MessageBox.Show("Please browse the path for ADGroups.csv");
            //}

            #endregion

            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();

            //if (!string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrEmpty(textBox2.Text))
            {
                StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

                while (!sr.EndOfStream)
                {
                    try
                    {
                        lstSiteColl.Add(sr.ReadLine().Trim());
                    }
                    catch
                    {
                        continue;
                    }
                }
            }

            //else
            //{
            //    MessageBox.Show("Please browse the path for SiteColl.csv / Reports folder");
            //}

            #endregion

            StreamWriter excelWriterScoringMatrixNew = null;

            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "ScoringMatrix" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");

            excelWriterScoringMatrixNew.WriteLine("SiteURL" + "," + "PageName" + "," + "PageURL" + "," + "WebPartTitle");
            excelWriterScoringMatrixNew.Flush();

            //Append_LinkFixedObjects("Item id" + "," + "Uniqid" + "," + "Item Url" + ","
            //                           + "Web url" + "," + "Web id");

            List<string> lists = new List<string>();
            lists.Add("1_Uploaded Files");
            lists.Add("Discussions");
            lists.Add("Events");
            lists.Add("Ideas");
            lists.Add("Posts");
            lists.Add("Tasks");
            lists.Add("Site Assets");
            //Migrated Documents Metadata

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();
                try
                {
                    string siteUrl = lstSiteColl[j].ToString().Trim();

                    //string siteUrl = "https://rsharepoint.sharepoint.com/sites/RSpace/departments/global-accounts/presales";

                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(siteUrl, "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        clientcontext.Load(clientcontext.Web, w1 => w1.ServerRelativeUrl, w => w.Url, wde => wde.Id);
                        clientcontext.ExecuteQuery();

                        //clientcontext.Load(clientcontext.Web.Lists);
                        //clientcontext.ExecuteQuery();

                        #region Document Webpart Title Report                        

                        try
                        {
                            //List doc1 = clientcontext.Web.Lists.GetByTitle("2_Documents and Pages");

                            //clientcontext.Load(doc1);
                            //clientcontext.ExecuteQuery();

                            //clientcontext.Load(doc1, d => d.RootFolder);
                            //clientcontext.ExecuteQuery();
                            //Folder f1 = doc1.RootFolder.Folders.GetByUrl("Documents");
                            //clientcontext.Load(f1);
                            //clientcontext.ExecuteQuery();

                            //FileCollection oFiles = f1.Files;
                            //clientcontext.Load(oFiles);
                            //clientcontext.ExecuteQuery();

                            Folder documentFolder = clientcontext.Web.GetFolderByServerRelativeUrl(clientcontext.Web.ServerRelativeUrl + "/Pages/Documents");
                            clientcontext.Load(documentFolder);
                            clientcontext.ExecuteQuery();


                            FileCollection oFiles = documentFolder.Files;
                            clientcontext.Load(oFiles);
                            clientcontext.ExecuteQuery();

                            foreach (Microsoft.SharePoint.Client.File oFile in oFiles)
                            {
                                clientcontext.Load(oFile);
                                clientcontext.ExecuteQuery();

                                LimitedWebPartManager wpm = oFile.GetLimitedWebPartManager(PersonalizationScope.Shared);

                                WebPartDefinitionCollection wdc = wpm.WebParts;
                                clientcontext.Load(wdc);
                                clientcontext.ExecuteQuery();

                                foreach (WebPartDefinition wpd in wpm.WebParts)
                                {
                                    WebPart wp = wpd.WebPart;

                                    clientcontext.Load(wpd);
                                    clientcontext.Load(wp);
                                    clientcontext.ExecuteQuery();

                                    string webpartTitle = wp.Title;

                                    string modifiedTitle = ConvertStringToUTF8(webpartTitle);

                                    if (modifiedTitle != string.Empty)
                                    {
                                        //User documentModifiedUser = oFile.ModifiedBy;
                                        //clientcontext.Load(documentModifiedUser);
                                        //clientcontext.ExecuteQuery();

                                        //DateTime documentLastUpdatedDate = oFile.TimeLastModified;

                                        oFile.CheckOut();
                                        wp.Title = modifiedTitle;

                                        wpd.SaveWebPartChanges();
                                        clientcontext.ExecuteQuery();



                                        //ListItem item = oFile.ListItemAllFields;
                                        //clientcontext.Load(item);
                                        //clientcontext.ExecuteQuery();

                                        //item["Modified"] = documentLastUpdatedDate;
                                        //item["Editor"] = documentModifiedUser;
                                        //item.Update();

                                        oFile.CheckIn("generalcheckin", CheckinType.MajorCheckIn);

                                        clientcontext.ExecuteQuery();

                                        excelWriterScoringMatrixNew.WriteLine(clientcontext.Web.Url + "," + oFile.Name + "," + oFile.ServerRelativeUrl + "," + wp.Title);
                                        excelWriterScoringMatrixNew.Flush();
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {

                        }
                        #endregion

                        #region EVENTS LIST DELETE

                        //try
                        //{
                        //    List doc1 = clientcontext.Web.Lists.GetByTitle("Events");
                        //    clientcontext.Load(doc1);
                        //    clientcontext.ExecuteQuery();

                        //    doc1.DeleteObject();
                        //    clientcontext.ExecuteQuery();
                        //}
                        //catch (Exception ex)
                        //{

                        //}

                        #endregion

                        #region JIVE DOCUMENTS DELETION

                        //try
                        //{
                        //    List doc = clientcontext.Web.Lists.GetByTitle("Site Assets");

                        //    clientcontext.Load(doc);
                        //    clientcontext.ExecuteQuery();

                        //    clientcontext.Load(doc, d => d.RootFolder);
                        //    clientcontext.ExecuteQuery();
                        //    Folder f = doc.RootFolder.Folders.GetByUrl("Migrated Documents Metadata");
                        //    clientcontext.Load(f);
                        //    clientcontext.ExecuteQuery();

                        //    f.DeleteObject();
                        //    clientcontext.ExecuteQuery();
                        //}
                        //catch (Exception ex)
                        //{

                        //}

                        //try
                        //{

                        //    List doc1 = clientcontext.Web.Lists.GetByTitle("2_Documents and Pages");

                        //    clientcontext.Load(doc1);
                        //    clientcontext.ExecuteQuery();

                        //    clientcontext.Load(doc1, d => d.RootFolder);
                        //    clientcontext.ExecuteQuery();
                        //    Folder f1 = doc1.RootFolder.Folders.GetByUrl("Documents");
                        //    clientcontext.Load(f1);
                        //    clientcontext.ExecuteQuery();

                        //    f1.DeleteObject();
                        //    clientcontext.ExecuteQuery();
                        //}
                        //catch (Exception ex)
                        //{

                        //}
                        #endregion

                        #region OVERVIEW, ACTIVITY PAGES DELETION

                        //try
                        //{
                        //    List doc1 = clientcontext.Web.Lists.GetByTitle("2_Documents and Pages");

                        //    clientcontext.Load(doc1);
                        //    clientcontext.ExecuteQuery();

                        //    clientcontext.Load(doc1, d => d.RootFolder);
                        //    clientcontext.ExecuteQuery();
                        //    try
                        //    {

                        //        Microsoft.SharePoint.Client.File f1 = doc1.RootFolder.Files.GetByUrl("Overview.aspx");
                        //        clientcontext.Load(f1);
                        //        clientcontext.ExecuteQuery();

                        //        f1.DeleteObject();
                        //        clientcontext.ExecuteQuery();
                        //    }
                        //    catch
                        //    {

                        //    }
                        //    try
                        //    {
                        //        Microsoft.SharePoint.Client.File f2 = doc1.RootFolder.Files.GetByUrl("Activity.aspx");
                        //        clientcontext.Load(f2);
                        //        clientcontext.ExecuteQuery();

                        //        f2.DeleteObject();
                        //        clientcontext.ExecuteQuery();
                        //    }
                        //    catch
                        //    {

                        //    }
                        //}
                        //catch (Exception ex)
                        //{

                        //} 
                        #endregion

                        #region OLD

                        //foreach (string list in lists)
                        //{

                        //    List doc = clientcontext.Web.Lists.GetByTitle(list);

                        //    clientcontext.Load(doc);
                        //    clientcontext.ExecuteQuery();

                        //    if (list == "Site Assets")
                        //    {
                        //        clientcontext.Load(doc, d => d.RootFolder);
                        //        clientcontext.ExecuteQuery();
                        //        Folder f = doc.RootFolder.Folders.GetByUrl("Migrated Documents Metadata");
                        //        clientcontext.Load(f);
                        //        clientcontext.ExecuteQuery();
                        //        CamlQuery caml2 = new CamlQuery();
                        //        caml2.ViewXml = "<View Scope=\"RecursiveAll\"><Query></Query></View>";
                        //        Microsoft.SharePoint.Client.ListItemCollection collection2 = doc.GetItems(caml2);
                        //        clientcontext.Load(collection2);
                        //        clientcontext.ExecuteQuery();

                        //        foreach (ListItem item in collection2)
                        //        {
                        //            try
                        //            {
                        //                clientcontext.Load(item, i => i.Id, i => i.File, i => i["Created"], i => i.DisplayName, i => i["FileRef"]);
                        //                clientcontext.ExecuteQuery();

                        //                if (item["FileRef"].ToString().Contains("Migrated Documents Metadata/"))
                        //                {

                        //                    try
                        //                    {
                        //                        Append_LinkFixedObjects(item.Id + "," + item.Id + "," + "https://rsharepoint.sharepoint.com" + item["FileRef"].ToString() + ","
                        //                                       + clientcontext.Web.Url + "," + clientcontext.Web.Id);
                        //                    }
                        //                    catch
                        //                    {

                        //                    }
                        //                    ListItem it = doc.GetItemById(item.Id);

                        //                    clientcontext.Load(it);
                        //                    clientcontext.ExecuteQuery();
                        //                    it.DeleteObject();
                        //                    clientcontext.ExecuteQuery();
                        //                }
                        //            }
                        //            catch (Exception ex)
                        //            {
                        //                continue;
                        //            }


                        //        }
                        //        continue;
                        //    }


                        //    CamlQuery caml = new CamlQuery();
                        //    caml.ViewXml = "<View Scope=\"RecursiveAll\"><Query></Query></View>";
                        //    Microsoft.SharePoint.Client.ListItemCollection collection = doc.GetItems(caml);
                        //    clientcontext.Load(collection);
                        //    clientcontext.ExecuteQuery();



                        //    foreach (ListItem item in collection)
                        //    {
                        //        try
                        //        {
                        //            clientcontext.Load(item, i => i.Id, i => i.File, i => i["Created"], i => i.DisplayName, i => i["FileRef"]);
                        //            clientcontext.ExecuteQuery();

                        //            //    if (Convert.ToDateTime(item["Created"]) >= Convert.ToDateTime("8/6/2018") && (!item["FileRef"].ToString().Contains("Blog Home.aspx")))
                        //            {
                        //                try
                        //                {
                        //                    Append_LinkFixedObjects(item.Id + "," + item.Id + "," + "https://rsharepoint.sharepoint.com" + item["FileRef"].ToString() + ","
                        //                                   + clientcontext.Web.Url + "," + clientcontext.Web.Id);
                        //                }
                        //                catch
                        //                {

                        //                }
                        //                ListItem it = doc.GetItemById(item.Id);

                        //                clientcontext.Load(it);
                        //                clientcontext.ExecuteQuery();
                        //                it.DeleteObject();
                        //                clientcontext.ExecuteQuery();

                        //            }




                        //        }
                        //        catch (Exception ex)
                        //        {

                        //            continue;
                        //        }



                        //    }
                        //} 

                        #endregion
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }

        private string ConvertStringToUTF8(string inputString)
        {

            string returnredString = string.Empty;

            byte[] bytes = Encoding.Default.GetBytes(inputString);

            string myString = Encoding.UTF8.GetString(bytes);

            myString = System.Web.HttpUtility.HtmlDecode(myString);


            bool isContainsAfter = myString.Contains("�");

            if (!isContainsAfter)
            {
                returnredString = myString;
            }

            return returnredString;

        }

        private void Append_LinkFixedObjects(string strLogfile)
        {

            StreamWriter fileWriter = default(StreamWriter);

            try
            {
                fileWriter = System.IO.File.AppendText(textBox2.Text + @"\deleted.csv");
                fileWriter.WriteLine(strLogfile);
                fileWriter.Close();

            }
            catch (Exception ex)
            {
                //backgroundWorker1.ReportProgress(0, ex.Message.ToString());
            }
        }

        private void btnGetHref_Click(object sender, EventArgs e)
        {


            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();

            //if (!string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrEmpty(textBox2.Text))
            {
                StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

                while (!sr.EndOfStream)
                {
                    try
                    {
                        lstSiteColl.Add(sr.ReadLine().Trim());
                    }
                    catch
                    {
                        continue;
                    }
                }
            }

            //else
            //{
            //    MessageBox.Show("Please browse the path for SiteColl.csv / Reports folder");
            //}

            #endregion

            StreamWriter excelWriterScoringMatrixNew = null;

            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "ScoringMatrix" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");

            excelWriterScoringMatrixNew.WriteLine("SiteURL" + "," + "PageName" + "," + "PageURL" + "," + "PageId" + "," + "Type" + "," + "WebPartTitle" + "," + "Url");
            excelWriterScoringMatrixNew.Flush();

            //Append_LinkFixedObjects("Item id" + "," + "Uniqid" + "," + "Item Url" + ","
            //                           + "Web url" + "," + "Web id");

            List<string> lists = new List<string>();
            lists.Add("1_Uploaded Files");
            lists.Add("Discussions");
            lists.Add("Events");
            lists.Add("Ideas");
            lists.Add("Posts");
            lists.Add("Tasks");
            lists.Add("Site Assets");
            //Migrated Documents Metadata

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();
                try
                {
                    //string siteUrl = lstSiteColl[j].ToString().Trim();

                    string siteUrl = "https://rsharepoint.sharepoint.com/sites/RSpaceShared/rla-resources/ricoh-latin-america-spaces/puerto-rico";

                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(siteUrl, "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        clientcontext.Load(clientcontext.Web, w1 => w1.ServerRelativeUrl, w => w.Url, wde => wde.Id);
                        clientcontext.ExecuteQuery();

                        //clientcontext.Load(clientcontext.Web.Lists);
                        //clientcontext.ExecuteQuery();

                        #region Document Webpart Title Report                        

                        try
                        {
                            List<Microsoft.SharePoint.Client.ListItem> items = new List<Microsoft.SharePoint.Client.ListItem>();

                            List doc1 = clientcontext.Web.Lists.GetByTitle("2_Documents and Pages");

                            CamlQuery camlQuery = new CamlQuery();
                            camlQuery.ViewXml = @"<View Scope='RecursiveAll'><Query></Query></View>";

                            Microsoft.SharePoint.Client.ListItemCollection listItems = doc1.GetItems(camlQuery);
                            clientcontext.Load(listItems);
                            clientcontext.ExecuteQuery();

                            foreach (ListItem item in listItems)
                            {
                                if (item.FileSystemObjectType == FileSystemObjectType.File)
                                {

                                    Microsoft.SharePoint.Client.File page = item.File;
                                    clientcontext.Load(page);
                                    clientcontext.ExecuteQuery();

                                    if (page.Name == "Overview.aspx")
                                    {
                                        #region Wepart Operations
                                        LimitedWebPartManager wpm = page.GetLimitedWebPartManager(PersonalizationScope.Shared);
                                        WebPartDefinitionCollection wdc = wpm.WebParts;
                                        clientcontext.Load(wdc);
                                        clientcontext.ExecuteQuery();

                                        foreach (WebPartDefinition wpd in wpm.WebParts)
                                        {
                                            WebPart wp = wpd.WebPart;
                                            clientcontext.Load(wp);
                                            clientcontext.ExecuteQuery();

                                            //let props = parseWebPartSchema(webPartXml.get_value());
                                            //console.log(props.TypeName);

                                            try
                                            {
                                                ClientResult<string> webPartXml = wpm.ExportWebPart(wpd.Id);
                                                clientcontext.ExecuteQuery();

                                                XmlDocument xmlDoc = new XmlDocument();
                                                xmlDoc.LoadXml(webPartXml.Value);

                                                XmlNodeList xmlList = xmlDoc.ChildNodes;

                                                if (xmlList.Count > 1)
                                                {
                                                    XmlNode xmlNode = xmlList[1];

                                                    foreach (XmlNode node in xmlNode.ChildNodes)
                                                    {
                                                        if (node.Name == "Content")
                                                        {
                                                            string cdataText = node.InnerText;

                                                            #region Get href


                                                            try
                                                            {

                                                                HtmlAgilityPack.HtmlDocument doc12 = new HtmlAgilityPack.HtmlDocument();
                                                                doc12.LoadHtml(cdataText);

                                                                HtmlNodeCollection link12 = doc12.DocumentNode.SelectNodes("//a");

                                                                if (link12 != null)
                                                                {

                                                                    List<string> link1 = link12.Select(x => x.Attributes["href"].Value).ToList<string>();

                                                                    foreach (string link in link1)
                                                                    {
                                                                        excelWriterScoringMatrixNew.WriteLine(clientcontext.Web.Url + "," + page.Name + "," + page.ServerRelativeUrl + "," + item.Id + "," + "anchor" + "," + wp.Title + "," + link);
                                                                        excelWriterScoringMatrixNew.Flush();
                                                                    }
                                                                }
                                                            }
                                                            catch (Exception ex)
                                                            {
                                                                Library.WriteLog("Error at getting href urls at withinwebpart:- PageName: " + page.ServerRelativeUrl + ";WebpartTitle: " + wp.Title, ex);
                                                            }



                                                            //string pattern = @"<a [^>]*?>(.*?)</a>";
                                                            //MatchCollection matches = Regex.Matches(cdataText, pattern, RegexOptions.IgnoreCase | RegexOptions.Singleline);

                                                            //foreach (Match match in matches)
                                                            //{
                                                            //    try
                                                            //    {

                                                            //        //HtmlAgilityPack.HtmlDocument doc3 = new HtmlAgilityPack.HtmlDocument();                                                            
                                                            //        //doc3.LoadHtml(match.Value);
                                                            //        //string link = doc3.DocumentNode.SelectSingleNode("//a").Attributes["href"].Value;

                                                            //        string anchorTag = match.Value;


                                                            //        string hrefUrl = XElement.Parse(anchorTag).Attribute("href").Value;

                                                            //        //var regex = new Regex("<a [^>]*href=(?:'(?<href>.*?)')|(?:\"(?<href>.*?)\")", RegexOptions.IgnoreCase);
                                                            //        //var urls = regex.Matches(anchorTag).OfType<Match>().Select(m => m.Groups["href"].Value).SingleOrDefault();

                                                            //        excelWriterScoringMatrixNew.WriteLine(clientcontext.Web.Url + "," + page.Name + "," + page.ServerRelativeUrl + "," + item.Id + "," + wp.Title + "," + hrefUrl);
                                                            //        excelWriterScoringMatrixNew.Flush();

                                                            //    }
                                                            //    catch (Exception ex)
                                                            //    {
                                                            //        Library.WriteLog("Error at getting href urls at withinwebpart:- PageName: " + page.ServerRelativeUrl + ";WebpartTitle: " + wp.Title, ex);
                                                            //    }

                                                            //}

                                                            #endregion

                                                            #region Get BGImge

                                                            List<string> listImageUrls = new List<string>();
                                                            MatchCollection mt = Regex.Matches(cdataText, @"[^}]?([^{]*{[^}]*})", RegexOptions.Multiline);

                                                            foreach (Match match in mt)
                                                            {
                                                                try
                                                                {
                                                                    string anchorTag = match.Value;

                                                                    String[] allClassAtrrs = (anchorTag.Split('{'))[1].Remove((anchorTag.Split('{'))[1].Length - 1).Split(';');
                                                                    foreach (string strAttr in allClassAtrrs)
                                                                    {
                                                                        string propertyName = strAttr.Split(':')[0].Trim();
                                                                        if ((propertyName == "background"))
                                                                        {
                                                                            string propertyValue = strAttr.Trim().Split(':')[1].Trim().Substring(5, strAttr.Trim().Split(':')[1].Trim().Length - 7);
                                                                            listImageUrls.Add(propertyValue);
                                                                        }
                                                                        else if ((propertyName == "background-image"))
                                                                        {

                                                                            string propertyValue = strAttr.Trim().Substring(23, strAttr.Trim().Length - 25);
                                                                            listImageUrls.Add(propertyValue);
                                                                        }

                                                                    }

                                                                    //string hrefUrl = XElement.Parse(anchorTag).Attribute("background").Value;
                                                                }
                                                                catch (Exception ex)
                                                                {
                                                                    Library.WriteLog("Error at getting cssImage urls:- PageName: " + page.ServerRelativeUrl + ";WebpartTitle: " + wp.Title, ex);
                                                                }
                                                            }

                                                            if (listImageUrls.Count > 0)
                                                            {
                                                                foreach (string cssUrl in listImageUrls)
                                                                {
                                                                    excelWriterScoringMatrixNew.WriteLine(clientcontext.Web.Url + "," + page.Name + "," + page.ServerRelativeUrl + "," + item.Id + "," + "css" + "," + wp.Title + "," + cssUrl);
                                                                    excelWriterScoringMatrixNew.Flush();
                                                                }

                                                            }

                                                            #endregion


                                                        }
                                                    }
                                                }

                                            }
                                            catch (Exception ex)
                                            {
                                                Library.WriteLog("Error at getting href urls at webpartLevel:- PageName: " + page.ServerRelativeUrl + ";WebpartTitle: " + wp.Title, ex);
                                            }
                                        }

                                        #endregion

                                    }

                                }
                            }


                        }
                        catch (Exception ex)
                        {
                            Library.WriteLog("Error at reading file and webparts:- PageName:" + siteUrl, ex);

                        }
                        #endregion

                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }

        private void btnStartPublish_Click(object sender, EventArgs e)
        {

            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();

            //if (!string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrEmpty(textBox2.Text))
            {
                StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

                while (!sr.EndOfStream)
                {
                    try
                    {
                        lstSiteColl.Add(sr.ReadLine().Trim());
                    }
                    catch
                    {
                        continue;
                    }
                }
            }

            //else
            //{
            //    MessageBox.Show("Please browse the path for SiteColl.csv / Reports folder");
            //}

            #endregion

            StreamWriter excelWriterScoringMatrixNew = null;

            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "ScoringMatrix" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");

            excelWriterScoringMatrixNew.WriteLine("SiteURL" + "," + "PageName" + "," + "PageURL" + "," + "PageId");
            excelWriterScoringMatrixNew.Flush();



            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();
                try
                {
                    string siteUrl = lstSiteColl[j].ToString().Trim();

                    //string siteUrl = "https://rsharepoint.sharepoint.com/sites/rworldgroups2/bnsf";

                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(siteUrl, "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        clientcontext.Load(clientcontext.Web, w1 => w1.ServerRelativeUrl, w => w.Url, wde => wde.Id);
                        clientcontext.ExecuteQuery();

                        //clientcontext.Load(clientcontext.Web.Lists);
                        //clientcontext.ExecuteQuery();

                        #region Publish Page                       

                        try
                        {
                            List<Microsoft.SharePoint.Client.ListItem> items = new List<Microsoft.SharePoint.Client.ListItem>();

                            List doc1 = clientcontext.Web.Lists.GetByTitle("2_Documents and Pages");

                            CamlQuery camlQuery = new CamlQuery();
                            camlQuery.ViewXml = @"<View Scope='RecursiveAll'><Query></Query></View>";

                            Microsoft.SharePoint.Client.ListItemCollection listItems = doc1.GetItems(camlQuery);
                            clientcontext.Load(listItems);
                            clientcontext.ExecuteQuery();

                            foreach (ListItem item in listItems)
                            {
                                if (item.FileSystemObjectType == FileSystemObjectType.File)
                                {

                                    Microsoft.SharePoint.Client.File page = item.File;
                                    clientcontext.Load(page);
                                    clientcontext.ExecuteQuery();

                                    if (page.Name != "Search.aspx")
                                    {
                                        if (item.File.Level == FileLevel.Draft)
                                        {

                                            try
                                            {
                                                clientcontext.Load(item, i => i["Modified"], i => i["Editor"]);
                                                clientcontext.ExecuteQuery();

                                                var modifiedDate = item["Modified"];
                                                var modifiedBy = item["Editor"];

                                                item.File.Publish("publishingcsom");
                                                clientcontext.ExecuteQuery();

                                                doc1.EnableMinorVersions = false;
                                                doc1.Update();


                                                clientcontext.Load(item);
                                                clientcontext.ExecuteQuery();

                                                item["Modified"] = modifiedDate;
                                                item["Editor"] = modifiedBy;

                                                item.Update();
                                                //item.SystemUpdate();
                                                clientcontext.ExecuteQuery();


                                                doc1.EnableMinorVersions = true;
                                                doc1.DraftVersionVisibility = DraftVisibilityType.Author;
                                                doc1.Update();
                                                clientcontext.ExecuteQuery();

                                                excelWriterScoringMatrixNew.WriteLine(clientcontext.Web.Url + "," + page.Name + "," + page.ServerRelativeUrl + "," + item.Id);
                                                excelWriterScoringMatrixNew.Flush();
                                            }
                                            catch (Exception ex)
                                            {
                                                Library.WriteLog("Error at publish webpage:- PageName:" + page.Name + "; PageUrl :" + page.ServerRelativeUrl, ex);
                                            }
                                        }
                                    }
                                }
                            }


                        }
                        catch (Exception ex)
                        {
                            Library.WriteLog("Error at reading file and webparts:- PageName:" + siteUrl, ex);

                        }
                        #endregion

                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }

        private void btnGetIFrameUrls_Click(object sender, EventArgs e)
        {

            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();

            //if (!string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrEmpty(textBox2.Text))
            {
                StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

                while (!sr.EndOfStream)
                {
                    try
                    {
                        lstSiteColl.Add(sr.ReadLine().Trim());
                    }
                    catch
                    {
                        continue;
                    }
                }
            }

            //else
            //{
            //    MessageBox.Show("Please browse the path for SiteColl.csv / Reports folder");
            //}

            #endregion

            StreamWriter excelWriterScoringMatrixNew = null;

            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "ScoringMatrix" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");

            excelWriterScoringMatrixNew.WriteLine("SiteURL" + "," + "PageName" + "," + "PageURL" + "," + "PageId" + "," + "Type" + "," + "WebPartTitle" + "," + "Url");
            excelWriterScoringMatrixNew.Flush();

            //Append_LinkFixedObjects("Item id" + "," + "Uniqid" + "," + "Item Url" + ","
            //                           + "Web url" + "," + "Web id");

            List<string> lists = new List<string>();
            lists.Add("1_Uploaded Files");
            lists.Add("Discussions");
            lists.Add("Events");
            lists.Add("Ideas");
            lists.Add("Posts");
            lists.Add("Tasks");
            lists.Add("Site Assets");
            //Migrated Documents Metadata

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();
                try
                {
                    string siteUrl = lstSiteColl[j].ToString().Trim();

                    //string siteUrl = @"https://rsharepoint.sharepoint.com/sites/RSpace/rla-resources/events/events-archive/rla-incentive-trip-archive/rla-incentive-trip-fy16/";

                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(siteUrl, "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        clientcontext.Load(clientcontext.Web, w1 => w1.ServerRelativeUrl, w => w.Url, wde => wde.Id);
                        clientcontext.ExecuteQuery();

                        //clientcontext.Load(clientcontext.Web.Lists);
                        //clientcontext.ExecuteQuery();

                        #region Document Webpart Title Report                        

                        try
                        {
                            List<Microsoft.SharePoint.Client.ListItem> items = new List<Microsoft.SharePoint.Client.ListItem>();

                            List doc1 = clientcontext.Web.Lists.GetByTitle("2_Documents and Pages");

                            CamlQuery camlQuery = new CamlQuery();
                            camlQuery.ViewXml = @"<View Scope='RecursiveAll'><Query></Query></View>";

                            Microsoft.SharePoint.Client.ListItemCollection listItems = doc1.GetItems(camlQuery);
                            clientcontext.Load(listItems);
                            clientcontext.ExecuteQuery();

                            foreach (ListItem item in listItems)
                            {
                                if (item.FileSystemObjectType == FileSystemObjectType.File)
                                {


                                    Microsoft.SharePoint.Client.File page = item.File;
                                    clientcontext.Load(page);
                                    clientcontext.ExecuteQuery();



                                    //if (page.Name == "Blog Home.aspx")
                                    //{
                                    #region Wepart Operations
                                    LimitedWebPartManager wpm = page.GetLimitedWebPartManager(PersonalizationScope.Shared);
                                    WebPartDefinitionCollection wdc = wpm.WebParts;
                                    clientcontext.Load(wdc);
                                    clientcontext.ExecuteQuery();

                                    foreach (WebPartDefinition wpd in wpm.WebParts)
                                    {
                                        WebPart wp = wpd.WebPart;
                                        clientcontext.Load(wp);
                                        clientcontext.ExecuteQuery();

                                        //let props = parseWebPartSchema(webPartXml.get_value());
                                        //console.log(props.TypeName);

                                        try
                                        {
                                            ClientResult<string> webPartXml = wpm.ExportWebPart(wpd.Id);
                                            clientcontext.ExecuteQuery();

                                            XmlDocument xmlDoc = new XmlDocument();
                                            xmlDoc.LoadXml(webPartXml.Value);

                                            XmlNodeList xmlList = xmlDoc.ChildNodes;

                                            if (xmlList.Count > 1)
                                            {
                                                XmlNode xmlNode = xmlList[1];

                                                foreach (XmlNode node in xmlNode.ChildNodes)
                                                {
                                                    if (node.Name == "Content")
                                                    {
                                                        string cdataText = node.InnerText;

                                                        #region Get href


                                                        try
                                                        {

                                                            HtmlAgilityPack.HtmlDocument doc12 = new HtmlAgilityPack.HtmlDocument();
                                                            doc12.LoadHtml(cdataText);

                                                            HtmlNodeCollection link12 = doc12.DocumentNode.SelectNodes("//iframe");

                                                            if (link12 != null)
                                                            {

                                                                List<string> link1 = link12.Select(x => x.Attributes["src"].Value).ToList<string>();

                                                                foreach (string link in link1)
                                                                {
                                                                    excelWriterScoringMatrixNew.WriteLine(clientcontext.Web.Url + "," + page.Name + "," + page.ServerRelativeUrl + "," + item.Id + "," + "iFrame" + "," + wp.Title + "," + link);
                                                                    excelWriterScoringMatrixNew.Flush();
                                                                }
                                                            }
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            Library.WriteLog("Error at getting href urls at withinwebpart:- PageName: " + page.ServerRelativeUrl + ";WebpartTitle: " + wp.Title, ex);
                                                        }

                                                        #endregion


                                                    }
                                                }
                                            }

                                        }
                                        catch (Exception ex)
                                        {
                                            Library.WriteLog("Error at getting href urls at webpartLevel:- PageName: " + page.ServerRelativeUrl + ";WebpartTitle: " + wp.Title, ex);
                                        }
                                    }

                                    #endregion

                                    //}

                                }
                            }


                        }
                        catch (Exception ex)
                        {
                            Library.WriteLog("Error at reading file and webparts:- PageName:" + siteUrl, ex);

                        }
                        #endregion

                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }

        private void btnGetUrlsByDate_Click(object sender, EventArgs e)
        {
            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();

            //if (!string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrEmpty(textBox2.Text))
            {
                StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

                while (!sr.EndOfStream)
                {
                    try
                    {
                        lstSiteColl.Add(sr.ReadLine().Trim());
                    }
                    catch
                    {
                        continue;
                    }
                }
            }

            //else
            //{
            //    MessageBox.Show("Please browse the path for SiteColl.csv / Reports folder");
            //}

            #endregion

            StreamWriter excelWriterScoringMatrixNew = null;

            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "ScoringMatrix" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");

            excelWriterScoringMatrixNew.WriteLine("SiteURL" + "," + "PageName" + "," + "PageURL" + "," + "PageId" + "," + "ModifiedAt");
            excelWriterScoringMatrixNew.Flush();

            //Append_LinkFixedObjects("Item id" + "," + "Uniqid" + "," + "Item Url" + ","
            //                           + "Web url" + "," + "Web id");

            List<string> lists = new List<string>();
            lists.Add("1_Uploaded Files");
            lists.Add("Discussions");
            lists.Add("Events");
            lists.Add("Ideas");
            lists.Add("Posts");
            lists.Add("Tasks");
            lists.Add("Site Assets");
            //Migrated Documents Metadata

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();
                try
                {
                    string siteUrl = lstSiteColl[j].ToString().Trim();

                    //string siteUrl = @"https://rsharepoint.sharepoint.com/sites/rspaceshared/ricoh-portfolio/hardware-product-portfolio/smart-operation-panel/application-site";

                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(siteUrl, "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        clientcontext.Load(clientcontext.Web, w1 => w1.ServerRelativeUrl, w => w.Url, wde => wde.Id);
                        clientcontext.ExecuteQuery();

                        //clientcontext.Load(clientcontext.Web.Lists);
                        //clientcontext.ExecuteQuery();

                        #region Document Webpart Title Report                        

                        try
                        {
                            List<Microsoft.SharePoint.Client.ListItem> items = new List<Microsoft.SharePoint.Client.ListItem>();

                            List doc1 = clientcontext.Web.Lists.GetByTitle("2_Documents and Pages");

                            CamlQuery camlQuery = new CamlQuery();
                            camlQuery.ViewXml = @"<View Scope='RecursiveAll'><Query></Query></View>";

                            Microsoft.SharePoint.Client.ListItemCollection listItems = doc1.GetItems(camlQuery);
                            clientcontext.Load(listItems);
                            clientcontext.ExecuteQuery();

                            foreach (ListItem item in listItems)
                            {
                                if (item.FileSystemObjectType == FileSystemObjectType.File)
                                {


                                    Microsoft.SharePoint.Client.File page = item.File;
                                    clientcontext.Load(page);
                                    clientcontext.ExecuteQuery();

                                    ListItem pageItem = page.ListItemAllFields;
                                    clientcontext.Load(pageItem, o => o["Editor"]);
                                    clientcontext.ExecuteQuery();

                                    Microsoft.SharePoint.Client.FieldUserValue ed = (Microsoft.SharePoint.Client.FieldUserValue)pageItem["Editor"];






                                    //  DateTime dateToCompare = new DateTime(2018,9,1);

                                    if (ed.LookupValue.Contains("svc jivemigration"))
                                    // if(page.TimeLastModified.Date >= dateToCompare)
                                    {
                                        excelWriterScoringMatrixNew.WriteLine(clientcontext.Web.Url + "," + page.Name + "," + page.ServerRelativeUrl + "," + item.Id + "," + page.TimeLastModified);
                                        excelWriterScoringMatrixNew.Flush();
                                    }



                                }
                            }


                        }
                        catch (Exception ex)
                        {
                            Library.WriteLog("Error at reading file and webparts:- PageName:" + siteUrl, ex);

                        }
                        #endregion

                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void btnGetAllUrlsRecursivly_Click(object sender, EventArgs e)
        {


            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();

            //if (!string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrEmpty(textBox2.Text))
            {
                StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

                while (!sr.EndOfStream)
                {
                    try
                    {
                        lstSiteColl.Add(sr.ReadLine().Trim());
                    }
                    catch
                    {
                        continue;
                    }
                }
            }

            //else
            //{
            //    MessageBox.Show("Please browse the path for SiteColl.csv / Reports folder");
            //}

            #endregion

            StreamWriter excelWriterScoringMatrixNew = null;

            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "ScoringMatrix" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");

            excelWriterScoringMatrixNew.WriteLine("SiteURL," + "PageName," + "PageURL," + "PageId," + "ListType," + "Type," + "WebPartTitle," + "Url," + "PromtedLinksTitle," + "LinkLocation," + "BackgroundImageLocation," + "DiscussionTitle," + "DiscusionId," + "DiscusionReplyId," + "DiscusionUrlType," + "discoussionUrl," + "PostTitle," + "PostId," + "PostUrl");
            excelWriterScoringMatrixNew.Flush();

            //Append_LinkFixedObjects("Item id" + "," + "Uniqid" + "," + "Item Url" + ","
            //                           + "Web url" + "," + "Web id");

            List<string> lists = new List<string>();
            lists.Add("1_Uploaded Files");
            lists.Add("Discussions");
            lists.Add("Events");
            lists.Add("Ideas");
            lists.Add("Posts");
            lists.Add("Tasks");
            lists.Add("Site Assets");


            DataTable dtResult = new DataTable();
            dtResult.Columns.Add("SiteURL");//0
            dtResult.Columns.Add("PageName");//1
            dtResult.Columns.Add("PageURL");//2
            dtResult.Columns.Add("PageId");//3
            dtResult.Columns.Add("ListType");//4
            dtResult.Columns.Add("Type");//5
            dtResult.Columns.Add("WebPartTitle");//6
            dtResult.Columns.Add("Url");//7

            dtResult.Columns.Add("PromtedLinksTitle");//8
            dtResult.Columns.Add("LinkLocation");//9
            dtResult.Columns.Add("BackgroundImageLocation");//10

            dtResult.Columns.Add("DiscussionTitle");//11
            dtResult.Columns.Add("DiscusionId");//12
            dtResult.Columns.Add("DiscusionReplyId");//13
            dtResult.Columns.Add("DiscusionUrlType");//14
            dtResult.Columns.Add("discoussionUrl");//15

            dtResult.Columns.Add("PostTitle");//16
            dtResult.Columns.Add("PostId");//17
            dtResult.Columns.Add("PostUrl");//18




            //Migrated Documents Metadata

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();

                string siteUrl = lstSiteColl[j].ToString().Trim();

                //string siteUrl = "https://rsharepoint.sharepoint.com/sites/RSpaceShared/rla-resources/ricoh-latin-america-spaces/puerto-rico";


                try
                {

                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(siteUrl, "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        clientcontext.Load(clientcontext.Web, w1 => w1.ServerRelativeUrl, w => w.Url, wde => wde.Id);
                        clientcontext.ExecuteQuery();

                        //clientcontext.Load(clientcontext.Web.Lists);
                        //clientcontext.ExecuteQuery();

                        #region Document Webpart Title Report                        

                        try
                        {

                            ListCollection listColl = clientcontext.Web.Lists;
                            clientcontext.Load(listColl);
                            clientcontext.ExecuteQuery();

                            #region pages


                            List<Microsoft.SharePoint.Client.ListItem> items = new List<Microsoft.SharePoint.Client.ListItem>();

                            List doc1 = clientcontext.Web.Lists.GetByTitle("2_Documents and Pages");


                            CamlQuery camlQuery = new CamlQuery();
                            camlQuery.ViewXml = @"<View Scope='RecursiveAll'><Query></Query></View>";

                            Microsoft.SharePoint.Client.ListItemCollection listItems = doc1.GetItems(camlQuery);
                            clientcontext.Load(listItems);
                            clientcontext.ExecuteQuery();

                            foreach (ListItem item in listItems)
                            {
                                if (item.FileSystemObjectType == FileSystemObjectType.File)
                                {


                                    Microsoft.SharePoint.Client.File page = item.File;
                                    clientcontext.Load(page);
                                    clientcontext.ExecuteQuery();

                                    #region Wepart Operations
                                    LimitedWebPartManager wpm = page.GetLimitedWebPartManager(PersonalizationScope.Shared);
                                    WebPartDefinitionCollection wdc = wpm.WebParts;
                                    clientcontext.Load(wdc);
                                    clientcontext.ExecuteQuery();

                                    foreach (WebPartDefinition wpd in wpm.WebParts)
                                    {
                                        WebPart wp = wpd.WebPart;
                                        clientcontext.Load(wp);
                                        clientcontext.ExecuteQuery();

                                        //let props = parseWebPartSchema(webPartXml.get_value());
                                        //console.log(props.TypeName);

                                        try
                                        {
                                            ClientResult<string> webPartXml = wpm.ExportWebPart(wpd.Id);
                                            clientcontext.ExecuteQuery();

                                            XmlDocument xmlDoc = new XmlDocument();
                                            xmlDoc.LoadXml(webPartXml.Value);

                                            XmlNodeList xmlList = xmlDoc.ChildNodes;

                                            if (xmlList.Count > 1)
                                            {
                                                XmlNode xmlNode = xmlList[1];

                                                foreach (XmlNode node in xmlNode.ChildNodes)
                                                {
                                                    if (node.Name == "Content")
                                                    {
                                                        string cdataText = node.InnerText;

                                                        #region Get href


                                                        try
                                                        {

                                                            HtmlAgilityPack.HtmlDocument doc12 = new HtmlAgilityPack.HtmlDocument();
                                                            doc12.LoadHtml(cdataText);

                                                            HtmlNodeCollection link12 = doc12.DocumentNode.SelectNodes("//a");

                                                            if (link12 != null)
                                                            {

                                                                List<string> link1 = link12.Select(x => x.Attributes["href"].Value).ToList<string>();

                                                                foreach (string link in link1)
                                                                {

                                                                    excelWriterScoringMatrixNew.WriteLine(clientcontext.Web.Url + "," + page.Name + "," + page.ServerRelativeUrl + "," + item.Id + "," + "Pages," + "anchor," + wp.Title + "," + link + "," + "," + "," + "," + "," + "," + "," + "," + "," + "," + "," + "");
                                                                    excelWriterScoringMatrixNew.Flush();

                                                                    DataRow drPageUrl = dtResult.NewRow();
                                                                    drPageUrl[0] = clientcontext.Web.Url;
                                                                    drPageUrl[1] = page.Name;
                                                                    drPageUrl[2] = page.ServerRelativeUrl;
                                                                    drPageUrl[3] = item.Id;
                                                                    drPageUrl[4] = "Pages";
                                                                    drPageUrl[5] = "anchor";
                                                                    drPageUrl[6] = wp.Title;
                                                                    drPageUrl[7] = link;

                                                                    dtResult.Rows.Add(drPageUrl);
                                                                }
                                                            }
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            Library.WriteLog("Error at getting href urls at withinwebpart:- PageName: " + page.ServerRelativeUrl + ";WebpartTitle: " + wp.Title, ex);
                                                        }


                                                        #endregion

                                                        #region Get BGImge

                                                        List<string> listImageUrls = new List<string>();
                                                        MatchCollection mt = Regex.Matches(cdataText, @"[^}]?([^{]*{[^}]*})", RegexOptions.Multiline);

                                                        foreach (Match match in mt)
                                                        {
                                                            try
                                                            {
                                                                string anchorTag = match.Value;

                                                                String[] allClassAtrrs = (anchorTag.Split('{'))[1].Remove((anchorTag.Split('{'))[1].Length - 1).Split(';');
                                                                foreach (string strAttr in allClassAtrrs)
                                                                {
                                                                    string propertyName = strAttr.Split(':')[0].Trim();
                                                                    if ((propertyName == "background"))
                                                                    {
                                                                        string propertyValue = strAttr.Trim().Split(':')[1].Trim().Substring(5, strAttr.Trim().Split(':')[1].Trim().Length - 7);
                                                                        listImageUrls.Add(propertyValue);
                                                                    }
                                                                    else if ((propertyName == "background-image"))
                                                                    {

                                                                        string propertyValue = strAttr.Trim().Substring(23, strAttr.Trim().Length - 25);
                                                                        listImageUrls.Add(propertyValue);
                                                                    }

                                                                }

                                                                //string hrefUrl = XElement.Parse(anchorTag).Attribute("background").Value;
                                                            }
                                                            catch (Exception ex)
                                                            {
                                                                Library.WriteLog("Error at getting cssImage urls:- PageName: " + page.ServerRelativeUrl + ";WebpartTitle: " + wp.Title, ex);
                                                            }
                                                        }

                                                        if (listImageUrls.Count > 0)
                                                        {
                                                            foreach (string cssUrl in listImageUrls)
                                                            {

                                                                excelWriterScoringMatrixNew.WriteLine(clientcontext.Web.Url + "," + page.Name + "," + page.ServerRelativeUrl + "," + item.Id + "," + "Pages," + "css," + wp.Title + "," + cssUrl + "," + "," + "," + "," + "," + "," + "," + "," + "," + "," + "," + "");
                                                                excelWriterScoringMatrixNew.Flush();

                                                                DataRow drPageUrl = dtResult.NewRow();
                                                                drPageUrl[0] = clientcontext.Web.Url;
                                                                drPageUrl[1] = page.Name;
                                                                drPageUrl[2] = page.ServerRelativeUrl;
                                                                drPageUrl[3] = item.Id;
                                                                drPageUrl[4] = "Pages";
                                                                drPageUrl[5] = "css";
                                                                drPageUrl[6] = wp.Title;
                                                                drPageUrl[7] = cssUrl;

                                                                dtResult.Rows.Add(drPageUrl);
                                                            }

                                                        }

                                                        #endregion

                                                    }
                                                }
                                            }

                                        }
                                        catch (Exception ex)
                                        {
                                            if (!ex.Message.Contains("This part is not exportable. To be exportable, a part must be personalizable and not have its ExportMode set to None."))
                                            {
                                                Library.WriteLog("Error at getting href urls at webpartLevel:- PageName: " + page.ServerRelativeUrl + ";WebpartTitle: " + wp.Title, ex);
                                            }
                                        }
                                    }

                                    #endregion

                                }
                            }

                            #endregion

                            #region GetPromotedLinks

                            try
                            {



                                Guid promotedLinksGuid = new Guid("{192efa95-e50c-475e-87ab-361cede5dd7f}");

                                List listPromotedList = listColl.FirstOrDefault(y => y.TemplateFeatureId == promotedLinksGuid);

                                if (listPromotedList != null)
                                {
                                    CamlQuery cmlQuery = new CamlQuery();
                                    cmlQuery.ViewXml = @"<View Scope='Recursive'><Query></Query></View>";

                                    Microsoft.SharePoint.Client.ListItemCollection listItems1 = listPromotedList.GetItems(cmlQuery);
                                    clientcontext.Load(listItems1);
                                    clientcontext.ExecuteQuery();
                                    if (listItems1.Count >= 1)
                                    {
                                        string bgUrl = ((Microsoft.SharePoint.Client.FieldUrlValue)(listItems1[0]["BackgroundImageLocation"])).Url;
                                        string locationUrl = ((Microsoft.SharePoint.Client.FieldUrlValue)(listItems1[0]["LinkLocation"])).Url;



                                        excelWriterScoringMatrixNew.WriteLine(clientcontext.Web.Url + "," + "," + "," + "," + "PromotedLinks," + "," + "," + "," + listItems1[0]["Title"] + "," + locationUrl + "," + bgUrl + "," + "," + "," + "," + "," + "," + "," + "," + "");
                                        excelWriterScoringMatrixNew.Flush();


                                        DataRow drPromotedLink = dtResult.NewRow();
                                        drPromotedLink[0] = clientcontext.Web.Url;
                                        drPromotedLink[4] = "PromotedLinks";
                                        drPromotedLink[8] = listItems1[0]["Title"];
                                        drPromotedLink[9] = locationUrl;
                                        drPromotedLink[10] = bgUrl;

                                        dtResult.Rows.Add(drPromotedLink);


                                    }

                                }

                            }
                            catch (Exception ex)
                            {
                                Library.WriteLog("Error at reading promoted links:- PageName:" + siteUrl, ex);

                            }
                            #endregion

                            #region DiscussionList

                            try
                            {



                                List lst = clientcontext.Web.Lists.GetByTitle("Discussions");
                                CamlQuery q = CamlQuery.CreateAllFoldersQuery();
                                ListItemCollection topics = lst.GetItems(q);
                                clientcontext.Load(topics);
                                clientcontext.ExecuteQuery();

                                string discusionTitlle = string.Empty;
                                string discusionId = string.Empty;
                                string discusionType = string.Empty;
                                string discussionurl = string.Empty;


                                foreach (ListItem discussionItem in topics)
                                {
                                    discusionTitlle = Convert.ToString(discussionItem["Title"]);
                                    discusionId = Convert.ToString(discussionItem.Id);
                                    string srtBody = Convert.ToString(discussionItem["Body"]);


                                    HtmlAgilityPack.HtmlDocument dicussionDoc12 = new HtmlAgilityPack.HtmlDocument();
                                    dicussionDoc12.LoadHtml(srtBody);

                                    HtmlNodeCollection dicussionlink12 = dicussionDoc12.DocumentNode.SelectNodes("//a");

                                    if (dicussionlink12 != null)
                                    {

                                        List<string> discusionlink1 = dicussionlink12.Select(x => x.Attributes["href"].Value).ToList<string>();

                                        foreach (string discoussionBodyUrl in discusionlink1)
                                        {


                                            excelWriterScoringMatrixNew.WriteLine(clientcontext.Web.Url + "," + "," + "," + "," + "Discussions," + "," + "," + "," + "," + "," + "," + discusionTitlle + "," + discusionId + "," + "," + "DiscussionBody," + discoussionBodyUrl + "," + "," + "," + "");
                                            excelWriterScoringMatrixNew.Flush();


                                            DataRow drDocumentBody = dtResult.NewRow();
                                            drDocumentBody[0] = clientcontext.Web.Url;
                                            drDocumentBody[4] = "Discussions";
                                            drDocumentBody[11] = discusionTitlle;
                                            drDocumentBody[13] = discusionId;
                                            drDocumentBody[14] = "DiscussionBody";
                                            drDocumentBody[15] = discoussionBodyUrl;

                                            dtResult.Rows.Add(drDocumentBody);
                                        }
                                    }


                                    ListItem topic = discussionItem;

                                    q = CamlQuery.CreateAllItemsQuery(100, "Title", "FileRef", "Body");
                                    q.FolderServerRelativeUrl = topic["FileRef"].ToString();
                                    ListItemCollection replies = lst.GetItems(q);
                                    clientcontext.Load(replies);
                                    clientcontext.ExecuteQuery();

                                    foreach (ListItem reply in replies)
                                    {

                                        string srtreplyBody = Convert.ToString(reply["Body"]);


                                        HtmlAgilityPack.HtmlDocument dicussionReplyDoc12 = new HtmlAgilityPack.HtmlDocument();
                                        dicussionReplyDoc12.LoadHtml(srtreplyBody);

                                        HtmlNodeCollection dicussionReplylink12 = dicussionReplyDoc12.DocumentNode.SelectNodes("//a");

                                        if (dicussionReplylink12 != null)
                                        {

                                            List<string> discusionReplylink1 = dicussionReplylink12.Select(x => x.Attributes["href"].Value).ToList<string>();


                                            foreach (string discoussionreplyUrl in discusionReplylink1)
                                            {


                                                excelWriterScoringMatrixNew.WriteLine(clientcontext.Web.Url + "," + "," + "," + "," + "Discussions," + "," + "," + "," + "," + "," + "," + discusionTitlle + "," + discusionId + "," + Convert.ToString(reply["Id"]) + "," + "DiscussionReplay," + discoussionreplyUrl + "," + "," + "," + "");
                                                excelWriterScoringMatrixNew.Flush();

                                                DataRow drDocumentReplyBody = dtResult.NewRow();

                                                drDocumentReplyBody[0] = clientcontext.Web.Url;
                                                drDocumentReplyBody[4] = "Discussions";
                                                drDocumentReplyBody[11] = discusionTitlle;
                                                drDocumentReplyBody[12] = discusionId;
                                                drDocumentReplyBody[13] = Convert.ToString(reply["Id"]);
                                                drDocumentReplyBody[14] = "DiscussionReplay";
                                                drDocumentReplyBody[14] = discoussionreplyUrl;

                                                dtResult.Rows.Add(drDocumentReplyBody);
                                            }


                                        }
                                    }

                                }
                            }
                            catch (Exception ex)
                            {
                                Library.WriteLog("Error at reading discussion links:- PageName:" + siteUrl, ex);

                            }

                            #endregion

                            #region PostList


                            try
                            {


                                List lstPosts = clientcontext.Web.Lists.GetByTitle("Posts");
                                CamlQuery qq = new CamlQuery();
                                qq.ViewXml = @"<View Scope='Recursive'><Query></Query></View>";
                                ListItemCollection postColl = lstPosts.GetItems(qq);
                                clientcontext.Load(postColl);
                                clientcontext.ExecuteQuery();

                                string postTitlle = string.Empty;
                                string postId = string.Empty;
                                foreach (ListItem PostItem in postColl)
                                {


                                    postTitlle = Convert.ToString(PostItem["Title"]);
                                    postId = Convert.ToString(PostItem.Id);
                                    string srtpostBody = Convert.ToString(PostItem["Body"]);


                                    HtmlAgilityPack.HtmlDocument postDoc12 = new HtmlAgilityPack.HtmlDocument();
                                    postDoc12.LoadHtml(srtpostBody);

                                    HtmlNodeCollection postlink12 = postDoc12.DocumentNode.SelectNodes("//a");

                                    if (postlink12 != null)
                                    {

                                        List<string> postlink1 = postlink12.Select(x => x.Attributes["href"].Value).ToList<string>();


                                        foreach (string postBodyUrl in postlink1)
                                        {

                                            excelWriterScoringMatrixNew.WriteLine(clientcontext.Web.Url + "," + "," + "," + "," + "Posts," + "," + "," + "," + "," + "," + "," + "," + "," + "," + "DiscussionReplay," + "," + postTitlle + "," + postId + "," + postBodyUrl);
                                            excelWriterScoringMatrixNew.Flush();

                                            DataRow drpostBody = dtResult.NewRow();

                                            drpostBody[0] = clientcontext.Web.Url;
                                            drpostBody[4] = "Posts";
                                            drpostBody[16] = postTitlle;
                                            drpostBody[17] = postId;
                                            drpostBody[18] = postBodyUrl;

                                            dtResult.Rows.Add(drpostBody);
                                        }


                                    }
                                }

                            }
                            catch (Exception ex)
                            {
                                Library.WriteLog("Error at reading Posts links:- PageName:" + siteUrl, ex);

                            }
                            #endregion



                        }
                        catch (Exception ex)
                        {
                            Library.WriteLog("Error at reading file and webparts:- PageName:" + siteUrl, ex);

                        }
                        #endregion

                    }
                }
                catch (Exception ex)
                {
                    Library.WriteLog("Error at creating context for site" + siteUrl, ex);
                }
            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();


            DataSet ds = new DataSet();
            ds.Tables.Add(dtResult);

            new ExportExcel().ExportToExcel(ds, textBox2.Text + "\\" + "ScoringMatrixExcel" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".xlsx");



            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }



        private void btnUpdateWebpart_Click(object sender, EventArgs e)
        {

            DeleteUpdate();



            //string siteUrl = "https://rsharepoint.sharepoint.com/sites/RSpace/spaces-control/test-spaces/html-code-test-3";
            //string pageUrl = "/sites/RSpace/spaces-control/test-spaces/html-code-test-3/Pages/Overview.aspx";

            //AuthenticationManager authManager = new AuthenticationManager();
            //using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(siteUrl, "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
            //{
            //    clientcontext.Load(clientcontext.Web, w1 => w1.ServerRelativeUrl, w => w.Url, wde => wde.Id);
            //    clientcontext.ExecuteQuery();

            //    //clientcontext.Load(clientcontext.Web.Lists);
            //    //clientcontext.ExecuteQuery();

            //    Microsoft.SharePoint.Client.File page = clientcontext.Web.GetFileByServerRelativeUrl(pageUrl);

            //    page.CheckOut();
            //    LimitedWebPartManager wpm = page.GetLimitedWebPartManager(PersonalizationScope.Shared);

            //    //Guid webPartId = new Guid("04bff9ae-ac91-4461-80a3-264ba86aa6a5");
            //    //ClientResult<string> webPartXml = wpm.ExportWebPart(webPartId);
            //    //context.ExecuteQuery();

            //    WebPartDefinitionCollection wdc = wpm.WebParts;
            //    clientcontext.Load(wdc);
            //    clientcontext.ExecuteQuery();

            //    string exportedXml = string.Empty;

            //    foreach (WebPartDefinition wpd in wpm.WebParts)
            //    {
            //        WebPart wp = wpd.WebPart;

            //        clientcontext.Load(wpd,X=>X.ZoneId);
            //        clientcontext.Load(wp);
            //        clientcontext.ExecuteQuery();

            //        try
            //        {


            //            if ((wpd.ZoneId == "xa78a3d6c3f2b4b29b79ae959b81d8c18") && wp.ZoneIndex == 0)
            //            {
            //                //Guid webPartId = new Guid("04bff9ae-ac91-4461-80a3-264ba86aa6a5");
            //                ClientResult<string> webPartXml = wpm.ExportWebPart(wpd.Id);
            //                clientcontext.ExecuteQuery();
            //                exportedXml = webPartXml.Value;

            //                break;
            //            }
            //        }
            //        catch(Exception ex)
            //        {

            //        }
            //    }


            //    exportedXml = exportedXml.Replace("https://rspace.ricoh-la.com/resources/statics/28952/marketing1.html?a=1482338587205", "https://rsharepoint.sharepoint.com/sites/RSpace/spaces-control/test-spaces/html-code-test-3/SiteAssets/ResourceFiles/marketing1.aspx");

            //    var importedWebpart = wpm.ImportWebPart(exportedXml);
            //    wpm.AddWebPart(importedWebpart.WebPart, "xa78a3d6c3f2b4b29b79ae959b81d8c18", Convert.ToInt32(0));


            //    page.CheckIn("generalcheckin", CheckinType.MajorCheckIn);
            //    clientcontext.ExecuteQuery();
            //}


            //zone 1 id : x6699e0e91c484c8e91390e6b909991ea
            //zone 2 id : xa78a3d6c3f2b4b29b79ae959b81d8c18
        }


        private void DeleteUpdate()
        {
            string siteUrl = "https://rsharepoint.sharepoint.com/sites/rworldgroups/solutioncenter-team";
            string pageUrl = "/sites/rworldgroups/solutioncenter-team/Pages/Overview.aspx";

            AuthenticationManager authManager = new AuthenticationManager();
            using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(siteUrl, "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
            {
                clientcontext.Load(clientcontext.Web, w1 => w1.ServerRelativeUrl, w => w.Url, wde => wde.Id);
                clientcontext.ExecuteQuery();

                //clientcontext.Load(clientcontext.Web.Lists);
                //clientcontext.ExecuteQuery();

                Microsoft.SharePoint.Client.File page = clientcontext.Web.GetFileByServerRelativeUrl(pageUrl);

                page.CheckOut();
                LimitedWebPartManager wpm = page.GetLimitedWebPartManager(PersonalizationScope.Shared);

                //Guid webPartId = new Guid("04bff9ae-ac91-4461-80a3-264ba86aa6a5");
                //ClientResult<string> webPartXml = wpm.ExportWebPart(webPartId);
                //context.ExecuteQuery();

                WebPartDefinitionCollection wdc = wpm.WebParts;
                clientcontext.Load(wdc);
                clientcontext.ExecuteQuery();

                string exportedXml = string.Empty;

                foreach (WebPartDefinition wpd in wpm.WebParts)
                {
                    WebPart wp = wpd.WebPart;

                    clientcontext.Load(wpd, X => X.ZoneId);
                    clientcontext.Load(wp);
                    clientcontext.ExecuteQuery();


                    //if(wp.Title == "HTML CSS CODE - DO NOT DELETE")
                    //{
                    //    wpd.DeleteWebPart();
                    //    clientcontext.ExecuteQuery();
                    //}

                    try
                    {


                        if ((wpd.ZoneId == "xf2ad02d0da5a43c29d8caba85a35d723") && wp.ZoneIndex == 8)
                        {
                            //Guid webPartId = new Guid("04bff9ae-ac91-4461-80a3-264ba86aa6a5");
                            ClientResult<string> webPartXml = wpm.ExportWebPart(wpd.Id);
                            clientcontext.ExecuteQuery();
                            exportedXml = webPartXml.Value;

                            wpd.DeleteWebPart();
                            clientcontext.ExecuteQuery();

                            break;
                        }
                    }
                    catch (Exception ex)
                    {

                    }
                }


                exportedXml = exportedXml.Replace("<Title>Content</Title>", "<Title>Formatted Text</Title>");

                exportedXml = exportedXml.Replace("<FrameType>None</FrameType>", "<FrameType>Default</FrameType>");

                

                var importedWebpart = wpm.ImportWebPart(exportedXml);
                wpm.AddWebPart(importedWebpart.WebPart, "xf2ad02d0da5a43c29d8caba85a35d723", Convert.ToInt32(8));


                page.CheckIn("generalcheckin", CheckinType.MajorCheckIn);
                clientcontext.ExecuteQuery();
            }

        }

        private void DeleteAndAddWebpart(string webPartTitle, string zoneId, string contentOfWebpart, string pageUrl, ClientContext context)
        {
            string zoneIndex = string.Empty;

            Microsoft.SharePoint.Client.File page = context.Web.GetFileByServerRelativeUrl(pageUrl);

            page.CheckOut();

            LimitedWebPartManager wpm = page.GetLimitedWebPartManager(PersonalizationScope.Shared);

            //Guid webPartId = new Guid("04bff9ae-ac91-4461-80a3-264ba86aa6a5");
            //ClientResult<string> webPartXml = wpm.ExportWebPart(webPartId);
            //context.ExecuteQuery();

            WebPartDefinitionCollection wdc = wpm.WebParts;
            context.Load(wdc);
            context.ExecuteQuery();


            foreach (WebPartDefinition wpd in wpm.WebParts)
            {
                WebPart wp = wpd.WebPart;

                context.Load(wpd);
                context.Load(wp);
                context.ExecuteQuery();


                string modifiedText = ConvertStringToUTF8(wp.Title);

                wp.Title = modifiedText;

                wpd.SaveWebPartChanges();
                context.ExecuteQuery();

                var x = wp.ExportMode;

                if (wp.Title.ToUpper().Trim() == webPartTitle.ToUpper().Trim())
                {
                    zoneIndex = Convert.ToString(wp.ZoneIndex);
                    wpd.DeleteWebPart();
                    context.ExecuteQuery();

                    break;
                }
            }

            string xmlWebPart = "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
               "<WebPart xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"" +
               " xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\"" +
               " xmlns=\"http://schemas.microsoft.com/WebPart/v2\">" +
               "<Title>" + webPartTitle + "</Title><FrameType>None</FrameType>" +

               "<Description>Use for formatted text, tables, and images.</Description>" +
               "<IsIncluded>true</IsIncluded><ZoneID></ZoneID><PartOrder>0</PartOrder>" +
               "<FrameState>Normal</FrameState><Height /><Width /><AllowRemove>true</AllowRemove>" +
               "<AllowZoneChange>true</AllowZoneChange><AllowMinimize>true</AllowMinimize>" +
               "<AllowConnect>true</AllowConnect><AllowEdit>true</AllowEdit>" +
               "<AllowHide>true</AllowHide><IsVisible>true</IsVisible><DetailLink /><HelpLink />" +
               "<HelpMode>Modeless</HelpMode><Dir>Default</Dir><PartImageSmall />" +
               "<MissingAssembly>Cannot import this Web Part.</MissingAssembly>" +
               "<PartImageLarge>/_layouts/images/mscontl.gif</PartImageLarge><IsIncludedFilter />" +
               "<Assembly>Microsoft.SharePoint, Version=13.0.0.0, Culture=neutral, " +
               "PublicKeyToken=94de0004b6e3fcc5</Assembly>" +
               "<TypeName>Microsoft.SharePoint.WebPartPages.ContentEditorWebPart</TypeName>" +
               "<ContentLink xmlns=\"http://schemas.microsoft.com/WebPart/v2/ContentEditor\" />" +
               "<Content xmlns=\"http://schemas.microsoft.com/WebPart/v2/ContentEditor\">" +
               "<![CDATA[" + contentOfWebpart + "]]></Content>" +
               "<PartStorage xmlns=\"http://schemas.microsoft.com/WebPart/v2/ContentEditor\" /></WebPart>";



            //var importedWebpart = wpm.ImportWebPart(xmlWebPart);
            //wpm.AddWebPart(importedWebpart.WebPart, zoneId, Convert.ToInt32(zoneIndex));

            page.CheckIn("generalcheckin", CheckinType.MajorCheckIn);

            context.ExecuteQuery();
        }

    }
}
