namespace SpreadSheetConnector
{
    partial class AddForm
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
            this.NameTextBox = new System.Windows.Forms.TextBox();
            this.PickFolderButton = new System.Windows.Forms.Button();
            this.GoogleSheetPathTextBox = new System.Windows.Forms.TextBox();
            this.ActionComboBox = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.FromRangeTextBox = new System.Windows.Forms.TextBox();
            this.ToRangeTextBox = new System.Windows.Forms.TextBox();
            this.AddButton = new System.Windows.Forms.Button();
            this.folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.SuspendLayout();
            // 
            // NameTextBox
            // 
            this.NameTextBox.Location = new System.Drawing.Point(13, 13);
            this.NameTextBox.Name = "NameTextBox";
            this.NameTextBox.Size = new System.Drawing.Size(259, 20);
            this.NameTextBox.TabIndex = 0;
            this.NameTextBox.Text = "Name";
            // 
            // PickFolderButton
            // 
            this.PickFolderButton.Location = new System.Drawing.Point(13, 40);
            this.PickFolderButton.Name = "PickFolderButton";
            this.PickFolderButton.Size = new System.Drawing.Size(259, 23);
            this.PickFolderButton.TabIndex = 1;
            this.PickFolderButton.Text = "Pick folder";
            this.PickFolderButton.UseVisualStyleBackColor = true;
            this.PickFolderButton.Click += new System.EventHandler(this.PickFolderButton_Click);
            // 
            // GoogleSheetPathTextBox
            // 
            this.GoogleSheetPathTextBox.Location = new System.Drawing.Point(13, 70);
            this.GoogleSheetPathTextBox.Name = "GoogleSheetPathTextBox";
            this.GoogleSheetPathTextBox.Size = new System.Drawing.Size(259, 20);
            this.GoogleSheetPathTextBox.TabIndex = 2;
            this.GoogleSheetPathTextBox.Text = "Google url";
            // 
            // ActionComboBox
            // 
            this.ActionComboBox.FormattingEnabled = true;
            this.ActionComboBox.Items.AddRange(new object[] {
            "Overwrite",
            "Append"});
            this.ActionComboBox.Location = new System.Drawing.Point(13, 97);
            this.ActionComboBox.Name = "ActionComboBox";
            this.ActionComboBox.Size = new System.Drawing.Size(259, 21);
            this.ActionComboBox.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 125);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(108, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Remove header rows";
            // 
            // FromRangeTextBox
            // 
            this.FromRangeTextBox.Location = new System.Drawing.Point(128, 125);
            this.FromRangeTextBox.Name = "FromRangeTextBox";
            this.FromRangeTextBox.Size = new System.Drawing.Size(32, 20);
            this.FromRangeTextBox.TabIndex = 5;
            // 
            // ToRangeTextBox
            // 
            this.ToRangeTextBox.Location = new System.Drawing.Point(167, 125);
            this.ToRangeTextBox.Name = "ToRangeTextBox";
            this.ToRangeTextBox.Size = new System.Drawing.Size(34, 20);
            this.ToRangeTextBox.TabIndex = 6;
            // 
            // AddButton
            // 
            this.AddButton.Location = new System.Drawing.Point(13, 167);
            this.AddButton.Name = "AddButton";
            this.AddButton.Size = new System.Drawing.Size(259, 23);
            this.AddButton.TabIndex = 7;
            this.AddButton.Text = "Add";
            this.AddButton.UseVisualStyleBackColor = true;
            this.AddButton.Click += new System.EventHandler(this.AddButton_Click);
            // 
            // AddForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 202);
            this.Controls.Add(this.AddButton);
            this.Controls.Add(this.ToRangeTextBox);
            this.Controls.Add(this.FromRangeTextBox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.ActionComboBox);
            this.Controls.Add(this.GoogleSheetPathTextBox);
            this.Controls.Add(this.PickFolderButton);
            this.Controls.Add(this.NameTextBox);
            this.Name = "AddForm";
            this.Text = "AddForm";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox NameTextBox;
        private System.Windows.Forms.Button PickFolderButton;
        private System.Windows.Forms.TextBox GoogleSheetPathTextBox;
        private System.Windows.Forms.ComboBox ActionComboBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox FromRangeTextBox;
        private System.Windows.Forms.TextBox ToRangeTextBox;
        private System.Windows.Forms.Button AddButton;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog;
    }
}