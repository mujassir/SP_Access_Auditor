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
        }

        private string GetPermissionString(string member, IEnumerable<SPPermission> permissions)
        {
            var res = permissions.FirstOrDefault(p => p.Member == member);
            return res != null ? res.Permissions : "";
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
            var siteInfo = DataManager.SiteCache;
            saveFileDialog1.DefaultExt = ".xlsx";
            saveFileDialog1.FileName = "*.xlsx";
            saveFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            var dialogResult = saveFileDialog1.ShowDialog();
            if (dialogResult != DialogResult.OK) return;
            var filePath = saveFileDialog1.FileName;

            ExportManager.ExportData(DataManager.SiteCache, permissionSummaries, filePath);
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

    }
}
