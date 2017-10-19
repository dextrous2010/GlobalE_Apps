using System.Drawing;

namespace GE_Merchant_Picker
{
    partial class GE_Merchant_Picker_Form
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
            this.merchantsListBox = new System.Windows.Forms.ListBox();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.goToSiteBtn = new System.Windows.Forms.Button();
            this.goToAdminBtn = new System.Windows.Forms.Button();
            this.returnPortalBtn = new System.Windows.Forms.Button();
            this.trackingPortalBtn = new System.Windows.Forms.Button();
            this.QaBtn = new System.Windows.Forms.Button();
            this.stagingBtn = new System.Windows.Forms.Button();
            this.productionBtn = new System.Windows.Forms.Button();
            this.goToGEAdminBtn = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // merchantsListBox
            // 
            this.merchantsListBox.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.merchantsListBox.FormattingEnabled = true;
            this.merchantsListBox.ItemHeight = 18;
            this.merchantsListBox.Location = new System.Drawing.Point(13, 49);
            this.merchantsListBox.Name = "merchantsListBox";
            this.merchantsListBox.Size = new System.Drawing.Size(192, 274);
            this.merchantsListBox.TabIndex = 0;
            this.merchantsListBox.SelectedIndexChanged += new System.EventHandler(this.merchantsListBox_SelectedIndexChanged);
            // 
            // richTextBox1
            // 
            this.richTextBox1.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.richTextBox1.Location = new System.Drawing.Point(212, 13);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.ReadOnly = true;
            this.richTextBox1.Size = new System.Drawing.Size(359, 167);
            this.richTextBox1.TabIndex = 1;
            this.richTextBox1.Text = "";
            // 
            // goToSiteBtn
            // 
            this.goToSiteBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.goToSiteBtn.Location = new System.Drawing.Point(212, 195);
            this.goToSiteBtn.Name = "goToSiteBtn";
            this.goToSiteBtn.Size = new System.Drawing.Size(359, 36);
            this.goToSiteBtn.TabIndex = 2;
            this.goToSiteBtn.Text = "Launch Merchant Site";
            this.goToSiteBtn.UseVisualStyleBackColor = true;
            this.goToSiteBtn.Click += new System.EventHandler(this.goToSiteBtn_Click);
            // 
            // goToAdminBtn
            // 
            this.goToAdminBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.goToAdminBtn.Location = new System.Drawing.Point(212, 282);
            this.goToAdminBtn.Name = "goToAdminBtn";
            this.goToAdminBtn.Size = new System.Drawing.Size(171, 39);
            this.goToAdminBtn.TabIndex = 2;
            this.goToAdminBtn.Text = "Plarform Admin";
            this.goToAdminBtn.UseVisualStyleBackColor = true;
            this.goToAdminBtn.Click += new System.EventHandler(this.goToAdminBtn_Click);
            // 
            // returnPortalBtn
            // 
            this.returnPortalBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.returnPortalBtn.Location = new System.Drawing.Point(400, 237);
            this.returnPortalBtn.Name = "returnPortalBtn";
            this.returnPortalBtn.Size = new System.Drawing.Size(171, 40);
            this.returnPortalBtn.TabIndex = 2;
            this.returnPortalBtn.Text = "Return Portal";
            this.returnPortalBtn.UseVisualStyleBackColor = true;
            this.returnPortalBtn.Click += new System.EventHandler(this.returnPortalBtn_Click);
            // 
            // trackingPortalBtn
            // 
            this.trackingPortalBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.trackingPortalBtn.Location = new System.Drawing.Point(400, 283);
            this.trackingPortalBtn.Name = "trackingPortalBtn";
            this.trackingPortalBtn.Size = new System.Drawing.Size(171, 40);
            this.trackingPortalBtn.TabIndex = 2;
            this.trackingPortalBtn.Text = "Tracking Portal";
            this.trackingPortalBtn.UseVisualStyleBackColor = true;
            this.trackingPortalBtn.Click += new System.EventHandler(this.trackingPortalBtn_Click);
            // 
            // QaBtn
            // 
            this.QaBtn.Location = new System.Drawing.Point(13, 13);
            this.QaBtn.Name = "QaBtn";
            this.QaBtn.Size = new System.Drawing.Size(52, 23);
            this.QaBtn.TabIndex = 3;
            this.QaBtn.Text = "QA";
            this.QaBtn.UseVisualStyleBackColor = true;
            this.QaBtn.Click += new System.EventHandler(this.QaBtn_Click);
            this.QaBtn.BackColor = Color.LightGreen;
            // 
            // stagingBtn
            // 
            this.stagingBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.stagingBtn.Location = new System.Drawing.Point(65, 13);
            this.stagingBtn.Name = "stagingBtn";
            this.stagingBtn.Size = new System.Drawing.Size(70, 23);
            this.stagingBtn.TabIndex = 3;
            this.stagingBtn.Text = "Staging";
            this.stagingBtn.UseVisualStyleBackColor = true;
            this.stagingBtn.Click += new System.EventHandler(this.stagingBtn_Click);
            // 
            // productionBtn
            // 
            this.productionBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.productionBtn.Location = new System.Drawing.Point(135, 13);
            this.productionBtn.Name = "productionBtn";
            this.productionBtn.Size = new System.Drawing.Size(71, 23);
            this.productionBtn.TabIndex = 3;
            this.productionBtn.Text = "Production";
            this.productionBtn.UseVisualStyleBackColor = true;
            this.productionBtn.Click += new System.EventHandler(this.productionBtn_Click);
            // 
            // goToGEAdminBtn
            // 
            this.goToGEAdminBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.goToGEAdminBtn.Location = new System.Drawing.Point(212, 238);
            this.goToGEAdminBtn.Name = "goToGEAdminBtn";
            this.goToGEAdminBtn.Size = new System.Drawing.Size(171, 40);
            this.goToGEAdminBtn.TabIndex = 2;
            this.goToGEAdminBtn.Text = "Global-E Admin";
            this.goToGEAdminBtn.UseVisualStyleBackColor = true;
            this.goToGEAdminBtn.Click += new System.EventHandler(this.goToGEAdminBtn_Click);
            // 
            // GE_Merchant_Picker_Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(583, 334);
            this.Controls.Add(this.productionBtn);
            this.Controls.Add(this.stagingBtn);
            this.Controls.Add(this.QaBtn);
            this.Controls.Add(this.trackingPortalBtn);
            this.Controls.Add(this.returnPortalBtn);
            this.Controls.Add(this.goToGEAdminBtn);
            this.Controls.Add(this.goToAdminBtn);
            this.Controls.Add(this.goToSiteBtn);
            this.Controls.Add(this.richTextBox1);
            this.Controls.Add(this.merchantsListBox);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "GE_Merchant_Picker_Form";
            this.Text = "GE Merchant Picker";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListBox merchantsListBox;
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.Button goToSiteBtn;
        private System.Windows.Forms.Button goToAdminBtn;
        private System.Windows.Forms.Button returnPortalBtn;
        private System.Windows.Forms.Button trackingPortalBtn;
        private System.Windows.Forms.Button QaBtn;
        private System.Windows.Forms.Button stagingBtn;
        private System.Windows.Forms.Button productionBtn;
        private System.Windows.Forms.Button goToGEAdminBtn;
    }
}

