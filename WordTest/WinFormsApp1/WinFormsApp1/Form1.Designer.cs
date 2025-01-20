namespace WinFormsApp1
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
            textBox1 = new TextBox();
            button1 = new Button();
            openFileDialog1 = new OpenFileDialog();
            groupBox1 = new GroupBox();
            upStartRow = new NumericUpDown();
            button2 = new Button();
            btnNewTable = new Button();
            udColumns = new NumericUpDown();
            btnColumns = new Button();
            udRows = new NumericUpDown();
            btnAddRows = new Button();
            btnCapitalPlanYear = new Button();
            btnCapitalQuarter = new Button();
            groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)upStartRow).BeginInit();
            ((System.ComponentModel.ISupportInitialize)udColumns).BeginInit();
            ((System.ComponentModel.ISupportInitialize)udRows).BeginInit();
            SuspendLayout();
            // 
            // textBox1
            // 
            textBox1.Location = new Point(138, 128);
            textBox1.Margin = new Padding(6, 6, 6, 6);
            textBox1.Name = "textBox1";
            textBox1.Size = new Size(1222, 38);
            textBox1.TabIndex = 0;
            // 
            // button1
            // 
            button1.Location = new Point(1376, 126);
            button1.Margin = new Padding(6, 6, 6, 6);
            button1.Name = "button1";
            button1.Size = new Size(150, 46);
            button1.TabIndex = 2;
            button1.Text = "Open File";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // openFileDialog1
            // 
            openFileDialog1.DefaultExt = "*.docx";
            openFileDialog1.FileName = "openFileDialog1";
            openFileDialog1.Filter = "Word 檔案|*.docx";
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(upStartRow);
            groupBox1.Controls.Add(button2);
            groupBox1.Controls.Add(btnNewTable);
            groupBox1.Controls.Add(udColumns);
            groupBox1.Controls.Add(btnColumns);
            groupBox1.Controls.Add(udRows);
            groupBox1.Controls.Add(btnAddRows);
            groupBox1.Location = new Point(138, 232);
            groupBox1.Margin = new Padding(6, 6, 6, 6);
            groupBox1.Name = "groupBox1";
            groupBox1.Padding = new Padding(6, 6, 6, 6);
            groupBox1.Size = new Size(392, 626);
            groupBox1.TabIndex = 5;
            groupBox1.TabStop = false;
            groupBox1.Text = "groupBox1";
            // 
            // upStartRow
            // 
            upStartRow.Location = new Point(44, 378);
            upStartRow.Margin = new Padding(6, 6, 6, 6);
            upStartRow.Name = "upStartRow";
            upStartRow.Size = new Size(122, 38);
            upStartRow.TabIndex = 11;
            // 
            // button2
            // 
            button2.Location = new Point(178, 378);
            button2.Margin = new Padding(6, 6, 6, 6);
            button2.Name = "button2";
            button2.Size = new Size(150, 46);
            button2.TabIndex = 10;
            button2.Text = "button2";
            button2.UseVisualStyleBackColor = true;
            // 
            // btnNewTable
            // 
            btnNewTable.Location = new Point(80, 64);
            btnNewTable.Margin = new Padding(6, 6, 6, 6);
            btnNewTable.Name = "btnNewTable";
            btnNewTable.Size = new Size(182, 58);
            btnNewTable.TabIndex = 9;
            btnNewTable.Text = "建新表";
            btnNewTable.UseVisualStyleBackColor = true;
            btnNewTable.Click += btnNewTable_Click;
            // 
            // udColumns
            // 
            udColumns.Location = new Point(40, 236);
            udColumns.Margin = new Padding(6, 6, 6, 6);
            udColumns.Name = "udColumns";
            udColumns.Size = new Size(126, 38);
            udColumns.TabIndex = 8;
            // 
            // btnColumns
            // 
            btnColumns.Location = new Point(178, 236);
            btnColumns.Margin = new Padding(6, 6, 6, 6);
            btnColumns.Name = "btnColumns";
            btnColumns.Size = new Size(150, 46);
            btnColumns.TabIndex = 7;
            btnColumns.Text = "建 n 欄";
            btnColumns.UseVisualStyleBackColor = true;
            // 
            // udRows
            // 
            udRows.Location = new Point(40, 156);
            udRows.Margin = new Padding(6, 6, 6, 6);
            udRows.Name = "udRows";
            udRows.Size = new Size(126, 38);
            udRows.TabIndex = 6;
            // 
            // btnAddRows
            // 
            btnAddRows.Location = new Point(178, 156);
            btnAddRows.Margin = new Padding(6, 6, 6, 6);
            btnAddRows.Name = "btnAddRows";
            btnAddRows.Size = new Size(150, 46);
            btnAddRows.TabIndex = 5;
            btnAddRows.Text = "建 n 列";
            btnAddRows.UseVisualStyleBackColor = true;
            // 
            // btnCapitalPlanYear
            // 
            btnCapitalPlanYear.Location = new Point(750, 271);
            btnCapitalPlanYear.Margin = new Padding(6, 6, 6, 6);
            btnCapitalPlanYear.Name = "btnCapitalPlanYear";
            btnCapitalPlanYear.Size = new Size(378, 64);
            btnCapitalPlanYear.TabIndex = 6;
            btnCapitalPlanYear.Text = "群益永續計畫-年度 Word 檔匯出";
            btnCapitalPlanYear.UseVisualStyleBackColor = true;
            btnCapitalPlanYear.Click += btnCapitalPlanYear_Click;
            // 
            // btnCapitalQuarter
            // 
            btnCapitalQuarter.Location = new Point(750, 369);
            btnCapitalQuarter.Margin = new Padding(6, 6, 6, 6);
            btnCapitalQuarter.Name = "btnCapitalQuarter";
            btnCapitalQuarter.Size = new Size(378, 64);
            btnCapitalQuarter.TabIndex = 7;
            btnCapitalQuarter.Text = "群益永續計畫-季度 Word 檔匯出";
            btnCapitalQuarter.UseVisualStyleBackColor = true;
            btnCapitalQuarter.Click += btnCapitalQuarter_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(14F, 30F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1600, 900);
            Controls.Add(btnCapitalQuarter);
            Controls.Add(btnCapitalPlanYear);
            Controls.Add(groupBox1);
            Controls.Add(button1);
            Controls.Add(textBox1);
            Margin = new Padding(6, 6, 6, 6);
            Name = "Form1";
            Text = "Form1";
            groupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)upStartRow).EndInit();
            ((System.ComponentModel.ISupportInitialize)udColumns).EndInit();
            ((System.ComponentModel.ISupportInitialize)udRows).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private TextBox textBox1;
        private Button button1;
        private OpenFileDialog openFileDialog1;
        private GroupBox groupBox1;
        private NumericUpDown udRows;
        private Button btnAddRows;
        private Button btnNewTable;
        private NumericUpDown udColumns;
        private Button btnColumns;
        private NumericUpDown upStartRow;
        private Button button2;
        private Button btnCapitalPlanYear;
        private Button btnCapitalQuarter;
    }
}
