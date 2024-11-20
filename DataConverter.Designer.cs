using System.Resources;

namespace GBC
{
    partial class DataConverter
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DataConverter));
            button1 = new Button();
            button2 = new Button();
            button3 = new Button();
            label1 = new Label();
            label2 = new Label();
            radioButton1 = new RadioButton();
            radioButton2 = new RadioButton();
            button4 = new Button();
            SuspendLayout();
            // 
            // button1
            // 
            button1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            button1.Location = new Point(12, 60);
            button1.Name = "button1";
            button1.Size = new Size(290, 108);
            button1.TabIndex = 0;
            button1.Text = "选择文件";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // button2
            // 
            button2.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            button2.Location = new Point(12, 187);
            button2.Name = "button2";
            button2.Size = new Size(290, 111);
            button2.TabIndex = 1;
            button2.Text = "导出路径";
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click;
            // 
            // button3
            // 
            button3.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            button3.Location = new Point(172, 318);
            button3.Name = "button3";
            button3.Size = new Size(130, 100);
            button3.TabIndex = 2;
            button3.Text = "开始";
            button3.UseVisualStyleBackColor = true;
            button3.Click += button3_Click;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(172, 421);
            label1.Name = "label1";
            label1.Size = new Size(130, 20);
            label1.TabIndex = 3;
            label1.Text = "Made by Whiter.";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Cursor = Cursors.Hand;
            label2.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Underline);
            label2.Location = new Point(12, 421);
            label2.Name = "label2";
            label2.Size = new Size(39, 20);
            label2.TabIndex = 4;
            label2.Text = "v0.4";
            label2.Click += label2_Click;
            // 
            // radioButton1
            // 
            radioButton1.AutoSize = true;
            radioButton1.Checked = true;
            radioButton1.Location = new Point(64, 21);
            radioButton1.Name = "radioButton1";
            radioButton1.Size = new Size(60, 24);
            radioButton1.TabIndex = 5;
            radioButton1.TabStop = true;
            radioButton1.Text = "去重";
            radioButton1.UseVisualStyleBackColor = true;
            // 
            // radioButton2
            // 
            radioButton2.AutoSize = true;
            radioButton2.Location = new Point(188, 21);
            radioButton2.Name = "radioButton2";
            radioButton2.Size = new Size(75, 24);
            radioButton2.TabIndex = 6;
            radioButton2.Text = "不去重";
            radioButton2.UseVisualStyleBackColor = true;
            // 
            // button4
            // 
            button4.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            button4.Location = new Point(12, 318);
            button4.Name = "button4";
            button4.Size = new Size(130, 100);
            button4.TabIndex = 7;
            button4.Text = "队伍名单\n（可选）";
            button4.UseVisualStyleBackColor = true;
            button4.Click += button4_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(9F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(318, 450);
            Controls.Add(button4);
            Controls.Add(radioButton2);
            Controls.Add(radioButton1);
            Controls.Add(label2);
            Controls.Add(label1);
            Controls.Add(button3);
            Controls.Add(button2);
            Controls.Add(button1);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "Data-Converter";
            Text = "GBC数据转换器";
            TopMost = true;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button button1;
        private Button button2;
        private Button button3;
        private Label label1;
        private Label label2;
        private RadioButton radioButton1;
        private RadioButton radioButton2;
        private Button button4;
    }
}
