namespace UploadBio
{
    partial class frmUploadBio
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
            this.btnUpload = new System.Windows.Forms.Button();
            this.dgvwLog = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.dgvwLog)).BeginInit();
            this.SuspendLayout();
            // 
            // btnUpload
            // 
            this.btnUpload.Location = new System.Drawing.Point(3, 12);
            this.btnUpload.Name = "btnUpload";
            this.btnUpload.Size = new System.Drawing.Size(279, 23);
            this.btnUpload.TabIndex = 0;
            this.btnUpload.Text = "Upload Bio Log";
            this.btnUpload.UseVisualStyleBackColor = true;
            this.btnUpload.Click += new System.EventHandler(this.btnUpload_Click);
            // 
            // dgvwLog
            // 
            this.dgvwLog.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvwLog.Location = new System.Drawing.Point(3, 41);
            this.dgvwLog.Name = "dgvwLog";
            this.dgvwLog.Size = new System.Drawing.Size(279, 209);
            this.dgvwLog.TabIndex = 1;
            // 
            // frmUploadBio
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 262);
            this.Controls.Add(this.dgvwLog);
            this.Controls.Add(this.btnUpload);
            this.Name = "frmUploadBio";
            this.Text = "Upload";
            this.Load += new System.EventHandler(this.frmUploadBio_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvwLog)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnUpload;
        private System.Windows.Forms.DataGridView dgvwLog;
    }
}

