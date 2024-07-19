namespace e_booking_import_data
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            btnBrowse = new Button();
            btnUpload = new Button();
            txtPath = new TextBox();
            txtSname = new TextBox();
            dataGridView1 = new DataGridView();
            dataGridView2 = new DataGridView();
            btnSave = new Button();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            ((System.ComponentModel.ISupportInitialize)dataGridView2).BeginInit();
            SuspendLayout();
            // 
            // btnBrowse
            // 
            btnBrowse.Location = new Point(87, 48);
            btnBrowse.Name = "btnBrowse";
            btnBrowse.Size = new Size(94, 29);
            btnBrowse.TabIndex = 0;
            btnBrowse.Text = "Browse";
            btnBrowse.UseVisualStyleBackColor = true;
            btnBrowse.Click += btnBrowse_Click;
            // 
            // btnUpload
            // 
            btnUpload.Location = new Point(87, 107);
            btnUpload.Name = "btnUpload";
            btnUpload.Size = new Size(94, 29);
            btnUpload.TabIndex = 1;
            btnUpload.Text = "Upload";
            btnUpload.UseVisualStyleBackColor = true;
            btnUpload.Click += btnUpload_Click;
            // 
            // txtPath
            // 
            txtPath.Location = new Point(218, 50);
            txtPath.Name = "txtPath";
            txtPath.Size = new Size(434, 27);
            txtPath.TabIndex = 2;
            // 
            // txtSname
            // 
            txtSname.Location = new Point(218, 109);
            txtSname.Name = "txtSname";
            txtSname.Size = new Size(434, 27);
            txtSname.TabIndex = 3;
            // 
            // dataGridView1
            // 
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Location = new Point(26, 164);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.RowHeadersWidth = 51;
            dataGridView1.Size = new Size(331, 299);
            dataGridView1.TabIndex = 4;
            // 
            // dataGridView2
            // 
            dataGridView2.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView2.Location = new Point(395, 164);
            dataGridView2.Name = "dataGridView2";
            dataGridView2.RowHeadersWidth = 51;
            dataGridView2.Size = new Size(331, 299);
            dataGridView2.TabIndex = 5;
            // 
            // btnSave
            // 
            btnSave.Location = new Point(330, 480);
            btnSave.Name = "btnSave";
            btnSave.Size = new Size(94, 29);
            btnSave.TabIndex = 6;
            btnSave.Text = "Save To DB";
            btnSave.UseVisualStyleBackColor = true;
            btnSave.Click += btnSave_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(754, 621);
            Controls.Add(btnSave);
            Controls.Add(dataGridView2);
            Controls.Add(dataGridView1);
            Controls.Add(txtSname);
            Controls.Add(txtPath);
            Controls.Add(btnUpload);
            Controls.Add(btnBrowse);
            Name = "Form1";
            Text = "Form1";
            Load += Form1_Load;
            ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
            ((System.ComponentModel.ISupportInitialize)dataGridView2).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button btnBrowse;
        private Button btnUpload;
        private TextBox txtPath;
        private TextBox txtSname;
        private DataGridView dataGridView1;
        private DataGridView dataGridView2;
        private Button btnSave;
    }
}
