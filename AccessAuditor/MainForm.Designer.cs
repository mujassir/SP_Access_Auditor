
namespace AccessAuditor
{
    partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.treeViewSiteContent = new System.Windows.Forms.TreeView();
            this.gridViewPermissions = new System.Windows.Forms.DataGridView();
            this.gridViewPermissionSummary = new System.Windows.Forms.DataGridView();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnExport = new System.Windows.Forms.Button();
            this.imageLoader = new System.Windows.Forms.PictureBox();
            this.btnLoad = new System.Windows.Forms.Button();
            this.txtSPSiteURL = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.gbPermissions = new System.Windows.Forms.GroupBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.treeViewMemberAccess = new System.Windows.Forms.TreeView();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            ((System.ComponentModel.ISupportInitialize)(this.gridViewPermissions)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridViewPermissionSummary)).BeginInit();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.imageLoader)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.gbPermissions.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.SuspendLayout();
            // 
            // treeViewSiteContent
            // 
            this.treeViewSiteContent.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treeViewSiteContent.Font = new System.Drawing.Font("Calibri", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.treeViewSiteContent.Location = new System.Drawing.Point(3, 27);
            this.treeViewSiteContent.Name = "treeViewSiteContent";
            this.treeViewSiteContent.Size = new System.Drawing.Size(677, 226);
            this.treeViewSiteContent.TabIndex = 0;
            this.treeViewSiteContent.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.treeView1_AfterSelect);
            // 
            // gridViewPermissions
            // 
            this.gridViewPermissions.AllowUserToAddRows = false;
            this.gridViewPermissions.AllowUserToDeleteRows = false;
            this.gridViewPermissions.AllowUserToResizeRows = false;
            this.gridViewPermissions.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.gridViewPermissions.BackgroundColor = System.Drawing.Color.White;
            this.gridViewPermissions.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gridViewPermissions.Dock = System.Windows.Forms.DockStyle.Fill;
            this.gridViewPermissions.Location = new System.Drawing.Point(3, 27);
            this.gridViewPermissions.MultiSelect = false;
            this.gridViewPermissions.Name = "gridViewPermissions";
            this.gridViewPermissions.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.gridViewPermissions.Size = new System.Drawing.Size(678, 226);
            this.gridViewPermissions.TabIndex = 2;
            // 
            // gridViewPermissionSummary
            // 
            this.gridViewPermissionSummary.AllowUserToAddRows = false;
            this.gridViewPermissionSummary.AllowUserToDeleteRows = false;
            this.gridViewPermissionSummary.AllowUserToResizeRows = false;
            this.gridViewPermissionSummary.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells;
            this.gridViewPermissionSummary.BackgroundColor = System.Drawing.Color.White;
            this.gridViewPermissionSummary.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gridViewPermissionSummary.Dock = System.Windows.Forms.DockStyle.Fill;
            this.gridViewPermissionSummary.Location = new System.Drawing.Point(3, 27);
            this.gridViewPermissionSummary.MultiSelect = false;
            this.gridViewPermissionSummary.Name = "gridViewPermissionSummary";
            this.gridViewPermissionSummary.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.gridViewPermissionSummary.Size = new System.Drawing.Size(677, 226);
            this.gridViewPermissionSummary.TabIndex = 2;
            this.gridViewPermissionSummary.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.gridViewPermissionSummary_CellContentClick);
            this.gridViewPermissionSummary.RowEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.gridViewPermissionSummary_RowEnter);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnExport);
            this.groupBox1.Controls.Add(this.imageLoader);
            this.groupBox1.Controls.Add(this.btnLoad);
            this.groupBox1.Controls.Add(this.txtSPSiteURL);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Top;
            this.groupBox1.Font = new System.Drawing.Font("Calibri", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(10);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(1379, 106);
            this.groupBox1.TabIndex = 3;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "SharePoint Site";
            // 
            // btnExport
            // 
            this.btnExport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExport.Location = new System.Drawing.Point(1186, 44);
            this.btnExport.Name = "btnExport";
            this.btnExport.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.btnExport.Size = new System.Drawing.Size(181, 39);
            this.btnExport.TabIndex = 3;
            this.btnExport.Text = "Export";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // imageLoader
            // 
            this.imageLoader.Image = global::AccessAuditor.Properties.Resources.ajax_loader__1_;
            this.imageLoader.Location = new System.Drawing.Point(882, 30);
            this.imageLoader.Name = "imageLoader";
            this.imageLoader.Size = new System.Drawing.Size(120, 58);
            this.imageLoader.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.imageLoader.TabIndex = 2;
            this.imageLoader.TabStop = false;
            this.imageLoader.Visible = false;
            // 
            // btnLoad
            // 
            this.btnLoad.Location = new System.Drawing.Point(695, 44);
            this.btnLoad.Name = "btnLoad";
            this.btnLoad.Size = new System.Drawing.Size(181, 39);
            this.btnLoad.TabIndex = 1;
            this.btnLoad.Text = "Load Permissions";
            this.btnLoad.UseVisualStyleBackColor = true;
            this.btnLoad.Click += new System.EventHandler(this.btnLoad_Click);
            // 
            // txtSPSiteURL
            // 
            this.txtSPSiteURL.Font = new System.Drawing.Font("Calibri", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtSPSiteURL.Location = new System.Drawing.Point(20, 46);
            this.txtSPSiteURL.Name = "txtSPSiteURL";
            this.txtSPSiteURL.Size = new System.Drawing.Size(663, 37);
            this.txtSPSiteURL.TabIndex = 0;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.gridViewPermissionSummary);
            this.groupBox2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox2.Font = new System.Drawing.Font("Calibri", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox2.Location = new System.Drawing.Point(3, 265);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(683, 256);
            this.groupBox2.TabIndex = 4;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Permissions Summary";
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Controls.Add(this.groupBox2, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.groupBox3, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.gbPermissions, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.groupBox4, 1, 1);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 106);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(1379, 524);
            this.tableLayoutPanel1.TabIndex = 5;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.treeViewSiteContent);
            this.groupBox3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox3.Font = new System.Drawing.Font("Calibri", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox3.Location = new System.Drawing.Point(3, 3);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(683, 256);
            this.groupBox3.TabIndex = 5;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Site Content";
            // 
            // gbPermissions
            // 
            this.gbPermissions.Controls.Add(this.gridViewPermissions);
            this.gbPermissions.Dock = System.Windows.Forms.DockStyle.Fill;
            this.gbPermissions.Font = new System.Drawing.Font("Calibri", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.gbPermissions.Location = new System.Drawing.Point(692, 3);
            this.gbPermissions.Name = "gbPermissions";
            this.gbPermissions.Size = new System.Drawing.Size(684, 256);
            this.gbPermissions.TabIndex = 6;
            this.gbPermissions.TabStop = false;
            this.gbPermissions.Text = "Permissions";
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.treeViewMemberAccess);
            this.groupBox4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox4.Font = new System.Drawing.Font("Calibri", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox4.Location = new System.Drawing.Point(692, 265);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(684, 256);
            this.groupBox4.TabIndex = 7;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Member Access";
            // 
            // treeViewMemberAccess
            // 
            this.treeViewMemberAccess.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treeViewMemberAccess.Location = new System.Drawing.Point(3, 27);
            this.treeViewMemberAccess.Name = "treeViewMemberAccess";
            this.treeViewMemberAccess.Size = new System.Drawing.Size(678, 226);
            this.treeViewMemberAccess.TabIndex = 0;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(239)))), ((int)(((byte)(239)))), ((int)(((byte)(239)))));
            this.ClientSize = new System.Drawing.Size(1379, 630);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Controls.Add(this.groupBox1);
            this.Name = "MainForm";
            this.Text = "Access Auditor";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.gridViewPermissions)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridViewPermissionSummary)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.imageLoader)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.gbPermissions.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TreeView treeViewSiteContent;
        private System.Windows.Forms.DataGridView gridViewPermissions;
        private System.Windows.Forms.DataGridView gridViewPermissionSummary;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox txtSPSiteURL;
        private System.Windows.Forms.Button btnLoad;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.GroupBox gbPermissions;
        private System.Windows.Forms.PictureBox imageLoader;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.TreeView treeViewMemberAccess;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
    }
}

