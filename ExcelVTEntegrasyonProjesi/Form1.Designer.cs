namespace ExcelVTEntegrasyonProjesi
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            btnReadData = new Button();
            richTextBox1 = new RichTextBox();
            btnReadExcel = new Button();
            richTextBox2 = new RichTextBox();
            SuspendLayout();
            // 
            // btnReadData
            // 
            btnReadData.BackColor = Color.FromArgb(0, 192, 0);
            btnReadData.ForeColor = Color.White;
            btnReadData.Location = new Point(39, 60);
            btnReadData.Name = "btnReadData";
            btnReadData.Size = new Size(129, 42);
            btnReadData.TabIndex = 0;
            btnReadData.Text = "Read Data And Wirte To Excel";
            btnReadData.UseVisualStyleBackColor = false;
            btnReadData.Click += btnReadData_Click;
            // 
            // richTextBox1
            // 
            richTextBox1.BackColor = Color.FromArgb(0, 192, 0);
            richTextBox1.ForeColor = Color.White;
            richTextBox1.Location = new Point(211, 12);
            richTextBox1.Name = "richTextBox1";
            richTextBox1.Size = new Size(577, 185);
            richTextBox1.TabIndex = 1;
            richTextBox1.Text = "";
            // 
            // btnReadExcel
            // 
            btnReadExcel.BackColor = Color.Blue;
            btnReadExcel.ForeColor = Color.White;
            btnReadExcel.Location = new Point(39, 306);
            btnReadExcel.Name = "btnReadExcel";
            btnReadExcel.Size = new Size(129, 42);
            btnReadExcel.TabIndex = 2;
            btnReadExcel.Text = "Read  From Excel And Add to DataBase";
            btnReadExcel.UseVisualStyleBackColor = false;
            btnReadExcel.Click += btnReadExcel_Click;
            // 
            // richTextBox2
            // 
            richTextBox2.BackColor = Color.Blue;
            richTextBox2.ForeColor = Color.White;
            richTextBox2.Location = new Point(211, 228);
            richTextBox2.Name = "richTextBox2";
            richTextBox2.Size = new Size(577, 185);
            richTextBox2.TabIndex = 3;
            richTextBox2.Text = "";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.FromArgb(64, 64, 64);
            ClientSize = new Size(800, 450);
            Controls.Add(richTextBox2);
            Controls.Add(btnReadExcel);
            Controls.Add(richTextBox1);
            Controls.Add(btnReadData);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "Form1";
            Text = "Data";
            ResumeLayout(false);
        }

        #endregion

        private Button btnReadData;
        private RichTextBox richTextBox1;
        private Button btnReadExcel;
        private RichTextBox richTextBox2;
    }
}
