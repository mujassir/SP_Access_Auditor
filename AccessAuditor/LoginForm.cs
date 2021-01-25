using AccessAuditor.BusinessLogic;
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
    public partial class LoginForm : Form
    {
        public delegate void ValidateDelegate(string spSiteURL, string username, string password);
        public event ValidateDelegate OnValidate;
        public LoginForm()
        {
            InitializeComponent();
        }

        private void LoginForm_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            imageLoader.Show();
            Task.Run(() =>
            {
                ValidateUser();
            });


        }

        private void ValidateUser()
        {
            try
            {
                var web = DataManager.ValidateCredentials(txtSPSiteURL.Text, txtUsername.Text, txtPassword.Text);
                this.BeginInvoke((MethodInvoker)delegate ()
                 {
                     if (OnValidate != null)
                     {
                         OnValidate.Invoke(txtSPSiteURL.Text, txtUsername.Text, txtPassword.Text);
                     }
                 });
               
            }
            catch (Exception ex)
            {
                MessageBox.Show("Username/password is incorrect or you may not have access to the given SharePoint Site", "Authentication Error");
                //throw;
            }
        }
    }
}
