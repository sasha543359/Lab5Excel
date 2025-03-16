namespace Lab5Excel
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
            writetoexcel = new Button();
            button2 = new Button();
            button3 = new Button();
            button4 = new Button();
            button5 = new Button();
            button6 = new Button();
            button7 = new Button();
            textBox1 = new TextBox();
            textBox2 = new TextBox();
            label1 = new Label();
            label2 = new Label();
            dataGridView1 = new DataGridView();
            dataGridView2 = new DataGridView();
            richTextBox1 = new RichTextBox();
            txtExcelSum = new TextBox();
            label3 = new Label();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            ((System.ComponentModel.ISupportInitialize)dataGridView2).BeginInit();
            SuspendLayout();
            // 
            // writetoexcel
            // 
            writetoexcel.Location = new Point(19, 38);
            writetoexcel.Name = "writetoexcel";
            writetoexcel.Size = new Size(217, 41);
            writetoexcel.TabIndex = 0;
            writetoexcel.Text = "1. Write to excel";
            writetoexcel.UseVisualStyleBackColor = true;
            writetoexcel.Click += button1_Click;
            // 
            // button2
            // 
            button2.Location = new Point(325, 38);
            button2.Name = "button2";
            button2.Size = new Size(217, 41);
            button2.TabIndex = 1;
            button2.Text = "2. ReadFromExcel";
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click;
            // 
            // button3
            // 
            button3.Location = new Point(19, 163);
            button3.Name = "button3";
            button3.Size = new Size(203, 41);
            button3.TabIndex = 2;
            button3.Text = "1. Text to grid";
            button3.UseVisualStyleBackColor = true;
            button3.Click += button3_Click;
            // 
            // button4
            // 
            button4.Location = new Point(325, 163);
            button4.Name = "button4";
            button4.Size = new Size(203, 41);
            button4.TabIndex = 3;
            button4.Text = "2. Grid to excel";
            button4.UseVisualStyleBackColor = true;
            button4.Click += button4_Click;
            // 
            // button5
            // 
            button5.Location = new Point(593, 163);
            button5.Name = "button5";
            button5.Size = new Size(203, 41);
            button5.TabIndex = 4;
            button5.Text = "3. Excel to grid";
            button5.UseVisualStyleBackColor = true;
            button5.Click += button5_Click;
            // 
            // button6
            // 
            button6.Location = new Point(593, 38);
            button6.Name = "button6";
            button6.Size = new Size(203, 41);
            button6.TabIndex = 5;
            button6.Text = "4. Grid To Text File";
            button6.UseVisualStyleBackColor = true;
            button6.Click += button6_Click;
            // 
            // button7
            // 
            button7.Location = new Point(593, 97);
            button7.Name = "button7";
            button7.Size = new Size(203, 41);
            button7.TabIndex = 6;
            button7.Text = "5. WrlnCellsExcel";
            button7.UseVisualStyleBackColor = true;
            button7.Click += button7_Click;
            // 
            // textBox1
            // 
            textBox1.Location = new Point(122, 107);
            textBox1.Name = "textBox1";
            textBox1.Size = new Size(100, 23);
            textBox1.TabIndex = 7;
            // 
            // textBox2
            // 
            textBox2.Location = new Point(428, 107);
            textBox2.Name = "textBox2";
            textBox2.Size = new Size(100, 23);
            textBox2.TabIndex = 8;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(38, 110);
            label1.Name = "label1";
            label1.Size = new Size(48, 15);
            label1.TabIndex = 9;
            label1.Text = "txtWrite";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(339, 115);
            label2.Name = "label2";
            label2.Size = new Size(46, 15);
            label2.TabIndex = 10;
            label2.Text = "txtRead";
            // 
            // dataGridView1
            // 
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Location = new Point(19, 244);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.Size = new Size(331, 433);
            dataGridView1.TabIndex = 11;
            // 
            // dataGridView2
            // 
            dataGridView2.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView2.Location = new Point(458, 236);
            dataGridView2.Name = "dataGridView2";
            dataGridView2.Size = new Size(325, 441);
            dataGridView2.TabIndex = 12;
            // 
            // richTextBox1
            // 
            richTextBox1.Location = new Point(882, 38);
            richTextBox1.Name = "richTextBox1";
            richTextBox1.Size = new Size(506, 588);
            richTextBox1.TabIndex = 13;
            richTextBox1.Text = "";
            // 
            // txtExcelSum
            // 
            txtExcelSum.Location = new Point(1000, 654);
            txtExcelSum.Name = "txtExcelSum";
            txtExcelSum.Size = new Size(100, 23);
            txtExcelSum.TabIndex = 14;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(882, 657);
            label3.Name = "label3";
            label3.Size = new Size(70, 15);
            label3.TabIndex = 15;
            label3.Text = "txtExcelSum";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1429, 793);
            Controls.Add(label3);
            Controls.Add(txtExcelSum);
            Controls.Add(richTextBox1);
            Controls.Add(dataGridView2);
            Controls.Add(dataGridView1);
            Controls.Add(label2);
            Controls.Add(label1);
            Controls.Add(textBox2);
            Controls.Add(textBox1);
            Controls.Add(button7);
            Controls.Add(button6);
            Controls.Add(button5);
            Controls.Add(button4);
            Controls.Add(button3);
            Controls.Add(button2);
            Controls.Add(writetoexcel);
            Name = "Form1";
            Text = "Form1";
            Load += Form1_Load;
            ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
            ((System.ComponentModel.ISupportInitialize)dataGridView2).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button writetoexcel;
        private Button button2;
        private Button button3;
        private Button button4;
        private Button button5;
        private Button button6;
        private Button button7;
        private TextBox textBox1;
        private TextBox textBox2;
        private Label label1;
        private Label label2;
        private DataGridView dataGridView1;
        private DataGridView dataGridView2;
        private RichTextBox richTextBox1;
        private TextBox txtExcelSum;
        private Label label3;
    }
}
