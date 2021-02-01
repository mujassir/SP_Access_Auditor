using AccessAuditor.BusinessLogic;
using AccessAuditor.Models;
using ClosedXML.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AccessAuditor
{
    public partial class MainForm : BaseForm
    {
        LoginForm loginForm = new LoginForm();

        private bool _isUserAuthenticated = false;
        private string _spSiteURL = "";
        private string _username = "";
        private string _password = "";


        private List<PermissionSummary> permissionSummaries;
        private Timer _timer;
        public MainForm()
        {
            InitializeComponent();

            _timer = new Timer();
            _timer.Interval = 10;
            _timer.Tick += _timer_Tick;
        }

        private void _timer_Tick(object sender, EventArgs e)
        {
            LoadSiteInfo();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Hide();
            loginForm.OnValidate += Form_OnValidate;
            loginForm.FormClosed += LoginForm_FormClosed;
            loginForm.ShowDialog();



        }

        private void LoginForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (_isUserAuthenticated == false)
            {
                Application.Exit();
            }
        }

        private void Form_OnValidate(string spSiteURL, string username, string password)
        {
            loginForm.Hide();
            this.Show();
            btnLoad.Enabled = false;
            txtSPSiteURL.Text = _spSiteURL = spSiteURL;
            _username = username;
            _password = password;
            LoadPermissions();
        }


        private void btnLoad_Click(object sender, EventArgs e)
        {
            DataManager.ClearCache();
            _spSiteURL = txtSPSiteURL.Text.Trim();
            if (_spSiteURL.Length == 0)
            {
                MessageBox.Show("Please provide SharePoint Site URL to load permissions", "Warning!");
            }
            LoadPermissions();
        }

        private void LoadPermissions()
        {
            var debug = Convert.ToBoolean(System.Configuration.ConfigurationSettings.AppSettings["debug"]);
            SPSite testData = null;

            if (debug)
                testData = LoadJsonTestData();

            Task.Run(() =>
        {
            btnExport.BeginInvoke((MethodInvoker)delegate () { btnExport.Enabled = false; });
            btnLoad.BeginInvoke((MethodInvoker)delegate () { btnLoad.Enabled = false; });
            ShowLoader();
            if (debug && testData != null)
            {

                DataManager.SiteCache = testData;
                LoadSiteInfo();

            }
            else
            {

                DataManager.LoadSiteDetails(_spSiteURL, _username, _password);
                LoadSiteInfo();
                if (debug)
                    SaveJsonTestData();

            }

            HideLoader();
            btnLoad.BeginInvoke((MethodInvoker)delegate () { btnLoad.Enabled = true; });
            btnExport.BeginInvoke((MethodInvoker)delegate () { btnExport.Enabled = true; });

        });
        }

        public SPSite LoadJsonTestData()
        {
            SPSite data = null;
            using (StreamReader r = new StreamReader("testdata.json"))
            {
                string json = r.ReadToEnd();
                data = JsonConvert.DeserializeObject<SPSite>(json);
            }
            return data;
        }
        public void SaveJsonTestData()
        {
            var debug = Convert.ToBoolean(System.Configuration.ConfigurationSettings.AppSettings["debug"]);


            if (debug)
            {
                SPSite data = null;
                using (var sr = new StreamWriter("testdata.json"))
                {
                    sr.WriteLine(JsonConvert.SerializeObject(DataManager.SiteCache));
                }
            }

        }

        private void LoadSiteInfo()
        {


            var siteInfo = DataManager.SiteCache;
            //treeViewSiteContent.Nodes.Clear();

            treeViewSiteContent.BeginInvoke((MethodInvoker)delegate () { treeViewSiteContent.Nodes.Clear(); });

            permissionSummaries = new List<PermissionSummary>();

            var rootNode = GetSiteInfoTreeNodes(siteInfo);

            //BuildRecursiveSiteNodes(siteInfo, sitesNodes);


            treeViewSiteContent.BeginInvoke((MethodInvoker)delegate ()
            {
                treeViewSiteContent.Nodes.Add(rootNode);
                treeViewSiteContent.SelectedNode = rootNode;
            });

            gridViewPermissionSummary.BeginInvoke((MethodInvoker)delegate ()
            {
                gridViewPermissionSummary.DataSource = permissionSummaries.Select(p => new
                {
                    p.Member,
                    p.Sites,
                    p.Lists,
                    p.FullControl,
                    p.Edit,
                    p.Read,
                    p.LimitedAccess

                }).ToList();
            });
        }

        private TreeNode GetSiteInfoTreeNodes(SPSite siteInfo)
        {
            TreeNode rootNode = new TreeNode(siteInfo.Title)
            {
                Tag = siteInfo.Permissions
            };
            AddPermissionsToSummary(siteInfo);

            if (siteInfo.Sites.Count() == 0 && siteInfo.Lists.Count() == 0) return rootNode;

            LoadSiteInfo(siteInfo, rootNode);
            return rootNode;
        }
        private void LoadSiteInfo(SPSite siteInfo, TreeNode rootNode)
        {
            if (siteInfo.Sites.Count() == 0 && siteInfo.Lists.Count() == 0)
                return;

            var listsNodes = new TreeNode("Lists");
            foreach (var list in siteInfo.Lists)
            {
                var child = new TreeNode(list.Title)
                {
                    Tag = list.Permissions
                };
                AddPermissionsToSummary(list);
                listsNodes.Nodes.Add(child);
            }
            if (listsNodes.Nodes.Count > 0)
                rootNode.Nodes.Add(listsNodes);

            var sitesNodes = new TreeNode("Subsites");

            foreach (var subsSite in siteInfo.Sites)
            {
                var child = new TreeNode(subsSite.Title);

                child.Tag = subsSite.Permissions;
                AddPermissionsToSummary(subsSite);
                sitesNodes.Nodes.Add(child);
                LoadSiteInfo(subsSite, child);
            }
            if (sitesNodes.Nodes.Count > 0)
                rootNode.Nodes.Add(sitesNodes);

        }


        private TreeNode GetSiteInfoTreeNodes(SPSite siteInfo, string filterByMemberAccess)
        {
            TreeNode rootNode = new TreeNode();
            var access = siteInfo.Permissions.FirstOrDefault(p => p.Member == filterByMemberAccess);
            if (access != null)
                rootNode = new TreeNode(siteInfo.Title + "(" + access.Permissions + ")");
            else
                rootNode = new TreeNode("Root");

            if (siteInfo.Sites.Count() == 0 && siteInfo.Lists.Count() == 0)
                return rootNode;

            LoadSiteInfo(siteInfo, rootNode, filterByMemberAccess);

            return rootNode;
        }



        private void LoadSiteInfo(SPSite siteInfo, TreeNode rootNode, string filterByMemberAccess)
        {
            if (siteInfo.Sites.Count() == 0 && siteInfo.Lists.Count() == 0)
                return;

            var listsNodes = new TreeNode("Lists");
            foreach (var list in siteInfo.Lists)
            {
                var access = list.Permissions.FirstOrDefault(p => p.Member == filterByMemberAccess);
                if (access == null) continue;
                listsNodes.Nodes.Add(new TreeNode(list.Title + "(" + access.Permissions + ")"));
            }
            if (listsNodes.Nodes.Count > 0)
                rootNode.Nodes.Add(listsNodes);

            var sitesNodes = new TreeNode("Subsites");

            foreach (var subsSite in siteInfo.Sites)
            {
                var access = subsSite.Permissions.FirstOrDefault(p => p.Member == filterByMemberAccess);
                if (access != null)
                {
                    var child = new TreeNode(subsSite.Title + "(" + access.Permissions + ")");
                    sitesNodes.Nodes.Add(child);
                    LoadSiteInfo(subsSite, child, filterByMemberAccess);
                }
                else
                {
                    LoadSiteInfo(subsSite, rootNode, filterByMemberAccess);
                }
            }
            if (sitesNodes.Nodes.Count > 0)
                rootNode.Nodes.Add(sitesNodes);
        }

        //private void BuildRecursiveSiteNodes(SPSite site, TreeNode node)
        //{
        //    foreach (var subsSite in site.Sites)
        //    {
        //        var child = new TreeNode(subsSite.Title);
        //        child.Tag = subsSite.Permissions;
        //        AddPermissionsToSummary(subsSite);
        //        node.Nodes.Add(child);
        //        BuildRecursiveSiteNodes(subsSite, child);
        //    }

        //}

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {

            //lblMessage.Text = "You selected: " + e.Node.FullPath;
            gridViewPermissions.DataSource = e.Node.Tag;
            gbPermissions.Text = string.Format("Permissions ({0})", e.Node.Text);

        }



        private void AddPermissionsToSummary(SPList list)
        {
            foreach (var permission in list.Permissions)
            {
                var entry = this.permissionSummaries.FirstOrDefault(p => p.Member == permission.Member);
                if (entry == null)
                {
                    entry = new PermissionSummary() { Member = permission.Member };
                    this.permissionSummaries.Add(entry);
                }
                entry.Lists++;
                if (permission.Permissions.Contains("Full Control")) entry.FullControl++;
                if (permission.Permissions.Contains("Edit")) entry.Edit++;
                if (permission.Permissions.Contains("Read")) entry.Read++;
                if (permission.Permissions.Contains("Limited Access")) entry.LimitedAccess++;
                entry.ListCollection.Add(list);
            }
        }

        private void AddPermissionsToSummary(SPSite site)
        {
            foreach (var permission in site.Permissions)
            {
                var entry = this.permissionSummaries.FirstOrDefault(p => p.Member == permission.Member);
                if (entry == null)
                {
                    entry = new PermissionSummary() { Member = permission.Member };
                    this.permissionSummaries.Add(entry);
                }
                entry.Sites++;
                if (permission.Permissions.Contains("Full Control")) entry.FullControl++;
                if (permission.Permissions.Contains("Edit")) entry.Edit++;
                if (permission.Permissions.Contains("Read")) entry.Read++;
                if (permission.Permissions.Contains("LimitedAccess")) entry.LimitedAccess++;
                entry.SiteCollection.Add(site);
            }
        }

        private void gridViewPermissionSummary_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

            //LoadMemberAccess(permissionSummary);

        }

        private string GetPermissionString(string member, IEnumerable<SPPermission> permissions)
        {
            var res = permissions.FirstOrDefault(p => p.Member == member);
            return res != null ? res.Permissions : "";
        }



        private void ShowLoader()
        {
            imageLoader.BeginInvoke((MethodInvoker)delegate ()
            {
                imageLoader.Show();
            });
        }

        private void HideLoader()
        {
            imageLoader.BeginInvoke((MethodInvoker)delegate ()
            {
                imageLoader.Hide();
            });
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }


        private MemoryStream SaveWorkbookToMemoryStream(XLWorkbook workbook)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                workbook.SaveAs(stream, new SaveOptions { EvaluateFormulasBeforeSaving = false, GenerateCalculationChain = false, ValidatePackage = false });
                return stream;
            }
        }


        private void gridViewPermissionSummary_RowEnter(object sender, DataGridViewCellEventArgs e)
        {

            if (e.RowIndex < 0) return;
            var permissionSummary = this.permissionSummaries[e.RowIndex];
            var member = permissionSummary.Member;
            treeViewMemberAccess.Nodes.Clear();
            var rootNote = GetSiteInfoTreeNodes(DataManager.SiteCache, member);
            treeViewMemberAccess.Nodes.Add(rootNote);
        }

        private void btnExport_Click(object sender, EventArgs e)
        {

            // ExportManager.ExportToExcel();
            var siteInfo = DataManager.SiteCache;
            var currentRow = 1;
            saveFileDialog1.DefaultExt = ".xlsx";
            saveFileDialog1.FileName = "*.xlsx";
            saveFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            var dialogResult = saveFileDialog1.ShowDialog();
            if (dialogResult != DialogResult.OK) return;

            //var fileName = saveFileDialog1.FileName;
            var filePath = saveFileDialog1.FileName;
            using (XLWorkbook wb = new XLWorkbook())
            {
                //Add DataTable in worksheet  


                var memberPermissionSummarySheet = wb.Worksheets.Add("Member Permission Summary");

                currentRow = 1;
                memberPermissionSummarySheet.Cell(currentRow, 1).Value = "Member";
                memberPermissionSummarySheet.Cell(currentRow, 2).Value = "Site";
                memberPermissionSummarySheet.Cell(currentRow, 3).Value = "Lists";
                memberPermissionSummarySheet.Cell(currentRow, 4).Value = "Full Control";
                memberPermissionSummarySheet.Cell(currentRow, 5).Value = "Edit";
                memberPermissionSummarySheet.Cell(currentRow, 6).Value = "Read";
                memberPermissionSummarySheet.Cell(currentRow, 7).Value = "Limited Access";

                memberPermissionSummarySheet.Row(currentRow).Style.Font.Bold = true;
                memberPermissionSummarySheet.Row(currentRow).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                memberPermissionSummarySheet.Row(currentRow).Style.Fill.BackgroundColor = XLColor.Yellow;

                foreach (var summary in this.permissionSummaries)
                {
                    currentRow++;
                    memberPermissionSummarySheet.Cell(currentRow, 1).Value = summary.Member;
                    memberPermissionSummarySheet.Cell(currentRow, 2).Value = summary.Sites;
                    memberPermissionSummarySheet.Cell(currentRow, 3).Value = summary.Lists;
                    memberPermissionSummarySheet.Cell(currentRow, 4).Value = summary.FullControl;
                    memberPermissionSummarySheet.Cell(currentRow, 5).Value = summary.Edit;
                    memberPermissionSummarySheet.Cell(currentRow, 6).Value = summary.Read;
                    memberPermissionSummarySheet.Cell(currentRow, 7).Value = summary.LimitedAccess;
                }


                memberPermissionSummarySheet.Columns().AdjustToContents();
                memberPermissionSummarySheet.Rows().AdjustToContents();

                currentRow = 1;
                var contenWisePermissionSheet = wb.Worksheets.Add("Content Wise Permissions");

                contenWisePermissionSheet.Cell(currentRow, 1).Value = "Site";
                contenWisePermissionSheet.Cell(currentRow, 2).Value = "Subsite";
                contenWisePermissionSheet.Cell(currentRow, 3).Value = "List/Library";
                contenWisePermissionSheet.Cell(currentRow, 4).Value = "Members";
                contenWisePermissionSheet.Cell(currentRow, 5).Value = "Permission";

                contenWisePermissionSheet.Row(currentRow).Style.Font.Bold = true;
                contenWisePermissionSheet.Row(currentRow).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                contenWisePermissionSheet.Row(currentRow).Style.Fill.BackgroundColor = XLColor.Yellow;

                //first data row will start here
                currentRow++;
                contenWisePermissionSheet.Cell(currentRow, 1).Value = siteInfo.Title;
                foreach (var subSite in siteInfo.Sites)
                {


                    foreach (var subSitePermission in subSite.Permissions)
                    {
                        currentRow++;
                        contenWisePermissionSheet.Cell(currentRow, 1).Value = siteInfo.Title;
                        contenWisePermissionSheet.Cell(currentRow, 2).Value = subSite.Title;
                        contenWisePermissionSheet.Cell(currentRow, 4).Value = subSitePermission.Member;
                        contenWisePermissionSheet.Cell(currentRow, 5).Value = subSitePermission.Permissions;
                    }
                }
                foreach (var list in siteInfo.Lists)
                {
                    foreach (var listPermission in list.Permissions)
                    {
                        currentRow++;
                        contenWisePermissionSheet.Cell(currentRow, 1).Value = siteInfo.Title;
                        contenWisePermissionSheet.Cell(currentRow, 3).Value = list.Title;
                        contenWisePermissionSheet.Cell(currentRow, 4).Value = listPermission.Member;
                        contenWisePermissionSheet.Cell(currentRow, 5).Value = listPermission.Permissions;

                    }
                }


                contenWisePermissionSheet.Columns().AdjustToContents();
                contenWisePermissionSheet.Rows().AdjustToContents();


                var memberWisePermissionSheet = wb.Worksheets.Add("Members Wise Permission");

                currentRow = 1;
                memberWisePermissionSheet.Cell(currentRow, 1).Value = "Member";
                memberWisePermissionSheet.Cell(currentRow, 2).Value = "Site";
                memberWisePermissionSheet.Cell(currentRow, 3).Value = "List";
                memberWisePermissionSheet.Cell(currentRow, 4).Value = "Permission";

                memberWisePermissionSheet.Row(currentRow).Style.Font.Bold = true;
                memberWisePermissionSheet.Row(currentRow).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                memberWisePermissionSheet.Row(currentRow).Style.Fill.BackgroundColor = XLColor.Yellow;

                foreach (var summary in this.permissionSummaries)
                {
                    foreach (var site in summary.SiteCollection)
                    {
                        currentRow++;
                        memberWisePermissionSheet.Cell(currentRow, 1).Value = summary.Member;
                        memberWisePermissionSheet.Cell(currentRow, 2).Value = site.Title;
                        memberWisePermissionSheet.Cell(currentRow, 4).Value = site.Permissions;
                    }

                    foreach (var list in summary.ListCollection)
                    {
                        currentRow++;
                        memberWisePermissionSheet.Cell(currentRow, 1).Value = summary.Member;
                        memberWisePermissionSheet.Cell(currentRow, 3).Value = list.Title;
                        memberWisePermissionSheet.Cell(currentRow, 4).Value = list.Permissions;
                    }


                }
                memberWisePermissionSheet.Columns().AdjustToContents();
                memberWisePermissionSheet.Rows().AdjustToContents();



                //using (MemoryStream memoryStream = SaveWorkbookToMemoryStream(wb))
                //{
                //    File.WriteAllBytes("permissions.xlsx", memoryStream.ToArray());
                //}

                var saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel files|*.xlsx",
                    Title = "Save an Excel File"
                };
                //if (!String.IsNullOrWhiteSpace(saveFileDialog.FileName))
                //    wb.SaveAs(saveFileDialog.FileName);
                wb.SaveAs(filePath);


            }
        }

    }
}
