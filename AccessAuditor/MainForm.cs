using AccessAuditor.BusinessLogic;
using AccessAuditor.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AccessAuditor
{
    public partial class MainForm : Form
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
            Task.Run(() =>
            {
                btnLoad.BeginInvoke((MethodInvoker)delegate () { btnLoad.Enabled = false; });
                ShowLoader();
                DataManager.LoadSiteDetails(_spSiteURL, _username, _password);
                LoadSiteInfo();
                HideLoader();
                btnLoad.BeginInvoke((MethodInvoker)delegate () { btnLoad.Enabled = true; });

            });
        }


        private void LoadSiteInfo()
        {


            var siteInfo = DataManager.SiteCache;
            treeViewSiteContent.Nodes.Clear();
            permissionSummaries = new List<PermissionSummary>();

            var rootNode = new TreeNode(siteInfo.Title);
            rootNode.Tag = siteInfo.Permissions;
            AddPermissionsToSummary(siteInfo);

            if (siteInfo.Sites.Count() == 0 && siteInfo.Lists.Count() == 0)
                return;

            var sitesNodes = new TreeNode("Sites");
            var listsNodes = new TreeNode("Lists");

            rootNode.Nodes.Add(sitesNodes);
            rootNode.Nodes.Add(listsNodes);

            foreach (var list in siteInfo.Lists)
            {
                var child = new TreeNode(list.Title);
                child.Tag = list.Permissions;
                AddPermissionsToSummary(list);
                listsNodes.Nodes.Add(child);
            }

            foreach (var site in siteInfo.Sites)
            {
                var child = new TreeNode(site.Title);
                child.Tag = site.Permissions;
                AddPermissionsToSummary(site);
                sitesNodes.Nodes.Add(child);
            }

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
                if (permission.Permissions.Contains("LimitedAccess")) entry.LimitedAccess++;
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
            if (e.RowIndex < 0) return;
            var permissionSummary = this.permissionSummaries[e.RowIndex];
            LoadMemberAccess(permissionSummary);

        }

        private void LoadMemberAccess(PermissionSummary permission)
        {


            treeViewMemberAccess.Nodes.Clear();


            if (permission.Sites == 0 && permission.Lists == 0)
                return;

            var sitesNodes = new TreeNode("Sites");
            var listsNodes = new TreeNode("Lists");
            foreach (var list in permission.ListCollection)
            {
                var title = string.Format("{0} ({1})", list.Title, GetPermissionString(permission.Member, list.Permissions));
                var child = new TreeNode(title);
                child.Tag = list.Permissions;
                listsNodes.Nodes.Add(child);
            }

            foreach (var site in permission.SiteCollection)
            {
                var title = string.Format("{0} ({1})", site.Title, GetPermissionString(permission.Member, site.Permissions));
                var child = new TreeNode(title);
                child.Tag = site.Permissions;
                sitesNodes.Nodes.Add(child);
            }
            treeViewMemberAccess.BeginInvoke((MethodInvoker)delegate ()
            {
                treeViewMemberAccess.Nodes.Add(sitesNodes);
                treeViewMemberAccess.Nodes.Add(listsNodes);
                treeViewMemberAccess.ExpandAll();
            });
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

    }
}
