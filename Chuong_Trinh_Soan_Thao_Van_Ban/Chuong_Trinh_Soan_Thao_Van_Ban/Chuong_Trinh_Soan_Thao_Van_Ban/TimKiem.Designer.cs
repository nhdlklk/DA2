namespace Chuong_Trinh_Soan_Thao_Van_Ban
{
    partial class frmtimkiem
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmtimkiem));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.txtsearch = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.fileSystemWatcher1 = new System.IO.FileSystemWatcher();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.raddown = new System.Windows.Forms.RadioButton();
            this.radup = new System.Windows.Forms.RadioButton();
            this.btnsearch = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.fileSystemWatcher1)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.txtsearch);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(282, 71);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Thông Tin Tìm Kiếm";
            // 
            // txtsearch
            // 
            this.txtsearch.Location = new System.Drawing.Point(76, 34);
            this.txtsearch.Name = "txtsearch";
            this.txtsearch.Size = new System.Drawing.Size(189, 20);
            this.txtsearch.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 37);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(50, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Tìm Kiếm";
            // 
            // fileSystemWatcher1
            // 
            this.fileSystemWatcher1.EnableRaisingEvents = true;
            this.fileSystemWatcher1.SynchronizingObject = this;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.raddown);
            this.groupBox2.Controls.Add(this.radup);
            this.groupBox2.Location = new System.Drawing.Point(134, 89);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(160, 52);
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Chức Năng";
            // 
            // raddown
            // 
            this.raddown.AutoSize = true;
            this.raddown.Checked = true;
            this.raddown.Location = new System.Drawing.Point(71, 21);
            this.raddown.Name = "raddown";
            this.raddown.Size = new System.Drawing.Size(56, 17);
            this.raddown.TabIndex = 3;
            this.raddown.TabStop = true;
            this.raddown.Text = "Xuống";
            this.raddown.UseVisualStyleBackColor = true;
            // 
            // radup
            // 
            this.radup.AutoSize = true;
            this.radup.Location = new System.Drawing.Point(15, 21);
            this.radup.Name = "radup";
            this.radup.Size = new System.Drawing.Size(43, 17);
            this.radup.TabIndex = 3;
            this.radup.Text = "Lên";
            this.radup.UseVisualStyleBackColor = true;
            // 
            // btnsearch
            // 
            this.btnsearch.Location = new System.Drawing.Point(12, 89);
            this.btnsearch.Name = "btnsearch";
            this.btnsearch.Size = new System.Drawing.Size(91, 33);
            this.btnsearch.TabIndex = 3;
            this.btnsearch.Text = "&Tìm";
            this.btnsearch.UseVisualStyleBackColor = true;
            this.btnsearch.Click += new System.EventHandler(this.btnsearch_Click);
            // 
            // frmtimkiem
            // 
            this.AcceptButton = this.btnsearch;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(328, 151);
            this.Controls.Add(this.btnsearch);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmtimkiem";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "TimKiem";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.fileSystemWatcher1)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox txtsearch;
        private System.Windows.Forms.Label label1;
        private System.IO.FileSystemWatcher fileSystemWatcher1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.RadioButton raddown;
        private System.Windows.Forms.RadioButton radup;
        private System.Windows.Forms.Button btnsearch;
    }
}