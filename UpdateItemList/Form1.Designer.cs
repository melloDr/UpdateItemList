using System.Threading.Tasks;

namespace UpdateItemList
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
        private async Task InitializeComponent()
        {
            txtFolder = new TextBox();
            btnFolder = new Button();
            grdSumItem = new DataGridView();
            btnWriteData = new Button();
            txtErr = new TextBox();
            btnCreateSheet = new Button();
            ((System.ComponentModel.ISupportInitialize)grdSumItem).BeginInit();
            SuspendLayout();
            // 
            // txtFolder
            // 
            txtFolder.Location = new Point(12, 12);
            txtFolder.Name = "txtFolder";
            txtFolder.Size = new Size(526, 27);
            txtFolder.TabIndex = 0;
            // 
            // btnFolder
            // 
            btnFolder.Location = new Point(12, 45);
            btnFolder.Name = "btnFolder";
            btnFolder.Size = new Size(94, 29);
            btnFolder.TabIndex = 1;
            btnFolder.Text = "Folder";
            btnFolder.UseVisualStyleBackColor = true;
            btnFolder.Click += btnFolder_Click;
            // 
            // grdSumItem
            // 
            grdSumItem.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            grdSumItem.Location = new Point(12, 148);
            grdSumItem.Name = "grdSumItem";
            grdSumItem.RowHeadersWidth = 51;
            grdSumItem.Size = new Size(231, 86);
            grdSumItem.TabIndex = 2;
            // 
            // btnWriteData
            // 
            btnWriteData.Location = new Point(112, 45);
            btnWriteData.Name = "btnWriteData";
            btnWriteData.Size = new Size(94, 29);
            btnWriteData.TabIndex = 3;
            btnWriteData.Text = "Write Data";
            btnWriteData.UseVisualStyleBackColor = true;
            btnWriteData.Click += btnWriteData_Click;
            // 
            // txtErr
            // 
            txtErr.Location = new Point(274, 130);
            txtErr.Multiline = true;
            txtErr.Name = "txtErr";
            txtErr.Size = new Size(616, 357);
            txtErr.TabIndex = 4;
            // 
            // btnCreateSheet
            // 
            btnCreateSheet.Location = new Point(212, 45);
            btnCreateSheet.Name = "btnCreateSheet";
            btnCreateSheet.Size = new Size(112, 29);
            btnCreateSheet.TabIndex = 5;
            btnCreateSheet.Text = "Create Sheet";
            btnCreateSheet.UseVisualStyleBackColor = true;
            btnCreateSheet.Click += btnCreateSheet_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(958, 530);
            Controls.Add(btnCreateSheet);
            Controls.Add(txtErr);
            Controls.Add(btnWriteData);
            Controls.Add(grdSumItem);
            Controls.Add(btnFolder);
            Controls.Add(txtFolder);
            Name = "Form1";
            Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)grdSumItem).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private TextBox txtFolder;
        private Button btnFolder;
        private DataGridView grdSumItem;
        private Button btnWriteData;
        private TextBox txtErr;
        private Button btnCreateSheet;
    }
}
