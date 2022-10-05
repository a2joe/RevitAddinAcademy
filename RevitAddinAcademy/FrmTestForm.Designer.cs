namespace RevitAddinAcademy
{
    partial class FrmTestForm
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
            this.label1 = new System.Windows.Forms.Label();
            this.lbxText = new System.Windows.Forms.ListBox();
            this.button1 = new System.Windows.Forms.Button();
            this.btnButton2 = new System.Windows.Forms.Button();
            this.btnButton3 = new System.Windows.Forms.Button();
            this.tbxTextBox = new System.Windows.Forms.TextBox();
            this.lbxText2 = new System.Windows.Forms.ListBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(71, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "This is a label";
            // 
            // lbxText
            // 
            this.lbxText.FormattingEnabled = true;
            this.lbxText.Location = new System.Drawing.Point(12, 34);
            this.lbxText.Name = "lbxText";
            this.lbxText.Size = new System.Drawing.Size(348, 394);
            this.lbxText.TabIndex = 1;
            this.lbxText.DoubleClick += new System.EventHandler(this.lbxText_DoubleClick);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(456, 34);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(142, 74);
            this.button1.TabIndex = 2;
            this.button1.Text = "Do Something";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // btnButton2
            // 
            this.btnButton2.Location = new System.Drawing.Point(456, 150);
            this.btnButton2.Name = "btnButton2";
            this.btnButton2.Size = new System.Drawing.Size(142, 74);
            this.btnButton2.TabIndex = 3;
            this.btnButton2.Text = "Do Something2";
            this.btnButton2.UseVisualStyleBackColor = true;
            this.btnButton2.Click += new System.EventHandler(this.btnButton2_Click);
            // 
            // btnButton3
            // 
            this.btnButton3.Location = new System.Drawing.Point(456, 266);
            this.btnButton3.Name = "btnButton3";
            this.btnButton3.Size = new System.Drawing.Size(142, 74);
            this.btnButton3.TabIndex = 4;
            this.btnButton3.Text = "Do Something3";
            this.btnButton3.UseVisualStyleBackColor = true;
            this.btnButton3.Click += new System.EventHandler(this.btnButton3_Click);
            // 
            // tbxTextBox
            // 
            this.tbxTextBox.Location = new System.Drawing.Point(456, 384);
            this.tbxTextBox.Name = "tbxTextBox";
            this.tbxTextBox.Size = new System.Drawing.Size(311, 20);
            this.tbxTextBox.TabIndex = 5;
            this.tbxTextBox.Text = "This is default text";
            // 
            // lbxText2
            // 
            this.lbxText2.FormattingEnabled = true;
            this.lbxText2.Location = new System.Drawing.Point(651, 34);
            this.lbxText2.Name = "lbxText2";
            this.lbxText2.Size = new System.Drawing.Size(137, 225);
            this.lbxText2.TabIndex = 6;
            this.lbxText2.SelectedIndexChanged += new System.EventHandler(this.listBox1_SelectedIndexChanged);
            // 
            // FrmTestForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.lbxText2);
            this.Controls.Add(this.tbxTextBox);
            this.Controls.Add(this.btnButton3);
            this.Controls.Add(this.btnButton2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.lbxText);
            this.Controls.Add(this.label1);
            this.Name = "FrmTestForm";
            this.Text = "This is a test form";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListBox lbxText;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button btnButton2;
        private System.Windows.Forms.Button btnButton3;
        private System.Windows.Forms.TextBox tbxTextBox;
        private System.Windows.Forms.ListBox lbxText2;
    }
}