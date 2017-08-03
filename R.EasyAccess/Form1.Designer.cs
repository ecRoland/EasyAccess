namespace R.EasyAccess
{
    partial class Form1
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
            this.open = new System.Windows.Forms.Button();
            this.list = new System.Windows.Forms.ComboBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.excel = new System.Windows.Forms.RadioButton();
            this.word = new System.Windows.Forms.RadioButton();
            this.path = new System.Windows.Forms.RadioButton();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // open
            // 
            this.open.Location = new System.Drawing.Point(212, 112);
            this.open.Name = "open";
            this.open.Size = new System.Drawing.Size(75, 23);
            this.open.TabIndex = 0;
            this.open.Text = "Open";
            this.open.UseVisualStyleBackColor = true;
            this.open.Click += new System.EventHandler(this.open_Click);
            // 
            // list
            // 
            this.list.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.list.FormattingEnabled = true;
            this.list.Items.AddRange(new object[] {
            "Item-1",
            "Item-2",
            "Item-3",
            "Item-4",
            "Item-5"});
            this.list.Location = new System.Drawing.Point(212, 12);
            this.list.Name = "list";
            this.list.Size = new System.Drawing.Size(75, 21);
            this.list.TabIndex = 1;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.path);
            this.groupBox1.Controls.Add(this.word);
            this.groupBox1.Controls.Add(this.excel);
            this.groupBox1.Location = new System.Drawing.Point(12, 6);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(93, 129);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            // 
            // excel
            // 
            this.excel.AutoSize = true;
            this.excel.Location = new System.Drawing.Point(6, 10);
            this.excel.Name = "excel";
            this.excel.Size = new System.Drawing.Size(51, 17);
            this.excel.TabIndex = 0;
            this.excel.TabStop = true;
            this.excel.Text = "Excel";
            this.excel.UseVisualStyleBackColor = true;
            // 
            // word
            // 
            this.word.AutoSize = true;
            this.word.Location = new System.Drawing.Point(6, 56);
            this.word.Name = "word";
            this.word.Size = new System.Drawing.Size(51, 17);
            this.word.TabIndex = 1;
            this.word.TabStop = true;
            this.word.Text = "Word";
            this.word.UseVisualStyleBackColor = true;
            // 
            // path
            // 
            this.path.AutoSize = true;
            this.path.Location = new System.Drawing.Point(6, 105);
            this.path.Name = "path";
            this.path.Size = new System.Drawing.Size(47, 17);
            this.path.TabIndex = 2;
            this.path.TabStop = true;
            this.path.Text = "Path";
            this.path.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(299, 146);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.list);
            this.Controls.Add(this.open);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "Form1";
            this.Text = "R>EasyAccess";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button open;
        private System.Windows.Forms.ComboBox list;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton path;
        private System.Windows.Forms.RadioButton word;
        private System.Windows.Forms.RadioButton excel;
    }
}

