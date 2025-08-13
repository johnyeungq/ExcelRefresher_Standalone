namespace ExcelRefresher_Standalone
{
    partial class ExcelRefresherForm
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
            ExcelPathLB = new ListBox();
            button1 = new Button();
            ExcelNameLB = new ListBox();
            tabControl1 = new TabControl();
            ManualPage = new TabPage();
            button3 = new Button();
            button2 = new Button();
            label1 = new Label();
            AddexcelPathTb = new TextBox();
            AddexcelNameTB = new TextBox();
            AutoPage = new TabPage();
            AutoLogLB = new ListBox();
            ExcelCountLB = new Label();
            AutoLB = new ListBox();
            DataPage = new TabPage();
            logfolderPathTB = new TextBox();
            ConfigPathTB = new TextBox();
            XMLpathTB = new TextBox();
            progressBar1 = new ProgressBar();
            tabControl1.SuspendLayout();
            ManualPage.SuspendLayout();
            AutoPage.SuspendLayout();
            DataPage.SuspendLayout();
            SuspendLayout();
            // 
            // ExcelPathLB
            // 
            ExcelPathLB.FormattingEnabled = true;
            ExcelPathLB.Location = new Point(248, 16);
            ExcelPathLB.Name = "ExcelPathLB";
            ExcelPathLB.Size = new Size(709, 144);
            ExcelPathLB.TabIndex = 0;
            ExcelPathLB.SelectedIndexChanged += ExcelPathLB_SelectedIndexChanged;
            // 
            // button1
            // 
            button1.Location = new Point(963, 53);
            button1.Name = "button1";
            button1.Size = new Size(131, 107);
            button1.TabIndex = 1;
            button1.Text = "Refresh";
            button1.UseVisualStyleBackColor = true;
            button1.Click += RefreshExcelBtn_Click;
            // 
            // ExcelNameLB
            // 
            ExcelNameLB.FormattingEnabled = true;
            ExcelNameLB.Location = new Point(6, 16);
            ExcelNameLB.Name = "ExcelNameLB";
            ExcelNameLB.Size = new Size(214, 144);
            ExcelNameLB.TabIndex = 0;
            ExcelNameLB.SelectedIndexChanged += ExcelNameLB_SelectedIndexChanged;
            // 
            // tabControl1
            // 
            tabControl1.Controls.Add(ManualPage);
            tabControl1.Controls.Add(AutoPage);
            tabControl1.Controls.Add(DataPage);
            tabControl1.Location = new Point(12, 12);
            tabControl1.Name = "tabControl1";
            tabControl1.SelectedIndex = 0;
            tabControl1.Size = new Size(1108, 322);
            tabControl1.TabIndex = 2;
            // 
            // ManualPage
            // 
            ManualPage.Controls.Add(button3);
            ManualPage.Controls.Add(button2);
            ManualPage.Controls.Add(label1);
            ManualPage.Controls.Add(AddexcelPathTb);
            ManualPage.Controls.Add(AddexcelNameTB);
            ManualPage.Controls.Add(ExcelNameLB);
            ManualPage.Controls.Add(button1);
            ManualPage.Controls.Add(ExcelPathLB);
            ManualPage.Location = new Point(4, 29);
            ManualPage.Name = "ManualPage";
            ManualPage.Padding = new Padding(3);
            ManualPage.Size = new Size(1100, 289);
            ManualPage.TabIndex = 0;
            ManualPage.Text = "Manual";
            ManualPage.UseVisualStyleBackColor = true;
            // 
            // button3
            // 
            button3.Location = new Point(963, 16);
            button3.Name = "button3";
            button3.Size = new Size(131, 29);
            button3.TabIndex = 3;
            button3.Text = "Run Auto";
            button3.UseVisualStyleBackColor = true;
            button3.Click += button3_Click;
            // 
            // button2
            // 
            button2.Location = new Point(963, 166);
            button2.Name = "button2";
            button2.Size = new Size(131, 74);
            button2.TabIndex = 4;
            button2.Text = "Add To Database";
            button2.UseVisualStyleBackColor = true;
            button2.Click += AddExcelBtn_Click;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(6, 190);
            label1.Name = "label1";
            label1.Size = new Size(87, 20);
            label1.TabIndex = 3;
            label1.Text = "Excel Name";
            // 
            // AddexcelPathTb
            // 
            AddexcelPathTb.Location = new Point(248, 213);
            AddexcelPathTb.Name = "AddexcelPathTb";
            AddexcelPathTb.Size = new Size(709, 27);
            AddexcelPathTb.TabIndex = 2;
            // 
            // AddexcelNameTB
            // 
            AddexcelNameTB.Location = new Point(6, 213);
            AddexcelNameTB.Name = "AddexcelNameTB";
            AddexcelNameTB.Size = new Size(214, 27);
            AddexcelNameTB.TabIndex = 2;
            // 
            // AutoPage
            // 
            AutoPage.Controls.Add(AutoLogLB);
            AutoPage.Controls.Add(ExcelCountLB);
            AutoPage.Controls.Add(AutoLB);
            AutoPage.Location = new Point(4, 29);
            AutoPage.Name = "AutoPage";
            AutoPage.Padding = new Padding(3);
            AutoPage.Size = new Size(1100, 289);
            AutoPage.TabIndex = 1;
            AutoPage.Text = "Auto";
            AutoPage.UseVisualStyleBackColor = true;
            AutoPage.Click += AutoPage_Click;
            // 
            // AutoLogLB
            // 
            AutoLogLB.FormattingEnabled = true;
            AutoLogLB.Location = new Point(15, 179);
            AutoLogLB.Name = "AutoLogLB";
            AutoLogLB.Size = new Size(1066, 104);
            AutoLogLB.TabIndex = 2;
            // 
            // ExcelCountLB
            // 
            ExcelCountLB.AutoSize = true;
            ExcelCountLB.Location = new Point(15, 24);
            ExcelCountLB.Name = "ExcelCountLB";
            ExcelCountLB.Size = new Size(97, 20);
            ExcelCountLB.TabIndex = 1;
            ExcelCountLB.Text = "Excel Count : ";
            // 
            // AutoLB
            // 
            AutoLB.FormattingEnabled = true;
            AutoLB.Location = new Point(15, 47);
            AutoLB.Name = "AutoLB";
            AutoLB.Size = new Size(939, 124);
            AutoLB.TabIndex = 0;
            // 
            // DataPage
            // 
            DataPage.Controls.Add(logfolderPathTB);
            DataPage.Controls.Add(ConfigPathTB);
            DataPage.Controls.Add(XMLpathTB);
            DataPage.Location = new Point(4, 29);
            DataPage.Name = "DataPage";
            DataPage.Size = new Size(1100, 289);
            DataPage.TabIndex = 2;
            DataPage.Text = "Data";
            DataPage.UseVisualStyleBackColor = true;
            // 
            // logfolderPathTB
            // 
            logfolderPathTB.Location = new Point(21, 178);
            logfolderPathTB.Name = "logfolderPathTB";
            logfolderPathTB.ReadOnly = true;
            logfolderPathTB.Size = new Size(450, 27);
            logfolderPathTB.TabIndex = 0;
            // 
            // ConfigPathTB
            // 
            ConfigPathTB.Location = new Point(21, 115);
            ConfigPathTB.Name = "ConfigPathTB";
            ConfigPathTB.ReadOnly = true;
            ConfigPathTB.Size = new Size(450, 27);
            ConfigPathTB.TabIndex = 0;
            // 
            // XMLpathTB
            // 
            XMLpathTB.Location = new Point(21, 49);
            XMLpathTB.Name = "XMLpathTB";
            XMLpathTB.ReadOnly = true;
            XMLpathTB.Size = new Size(450, 27);
            XMLpathTB.TabIndex = 0;
            // 
            // progressBar1
            // 
            progressBar1.Dock = DockStyle.Bottom;
            progressBar1.Location = new Point(0, 339);
            progressBar1.Name = "progressBar1";
            progressBar1.Size = new Size(1125, 29);
            progressBar1.TabIndex = 1;
            // 
            // ExcelRefresherForm
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1125, 368);
            Controls.Add(progressBar1);
            Controls.Add(tabControl1);
            Name = "ExcelRefresherForm";
            Text = "ExcelRefresher_CS";
            Load += ExcelRefresherForm_Load;
            tabControl1.ResumeLayout(false);
            ManualPage.ResumeLayout(false);
            ManualPage.PerformLayout();
            AutoPage.ResumeLayout(false);
            AutoPage.PerformLayout();
            DataPage.ResumeLayout(false);
            DataPage.PerformLayout();
            ResumeLayout(false);
        }

        #endregion

        private ListBox ExcelPathLB;
        private Button button1;
        private ListBox ExcelNameLB;
        private TabControl tabControl1;
        private TabPage ManualPage;
        private TabPage AutoPage;
        private TextBox AddexcelNameTB;
        private Label label1;
        private TextBox AddexcelPathTb;
        private Button button2;
        private ListBox AutoLB;
        private Label ExcelCountLB;
        private TabPage DataPage;
        private TextBox XMLpathTB;
        private TextBox ConfigPathTB;
        private ProgressBar progressBar1;
        private ListBox AutoLogLB;
        private TextBox logfolderPathTB;
        private Button button3;
    }
}
