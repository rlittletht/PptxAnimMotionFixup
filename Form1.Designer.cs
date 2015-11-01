namespace FixPPTLayout
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
            this.label1 = new System.Windows.Forms.Label();
            this.m_eb43 = new System.Windows.Forms.TextBox();
            this.m_eb169 = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.m_ebOutput = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.m_ebFixLayout = new System.Windows.Forms.Button();
            this.m_pbOpen = new System.Windows.Forms.Button();
            this.m_cbSlides = new System.Windows.Forms.ComboBox();
            this.m_ebFixAll = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(48, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "4x3 PPT";
            // 
            // m_eb43
            // 
            this.m_eb43.Location = new System.Drawing.Point(75, 10);
            this.m_eb43.Name = "m_eb43";
            this.m_eb43.Size = new System.Drawing.Size(335, 20);
            this.m_eb43.TabIndex = 1;
            this.m_eb43.Text = "c:\\temp\\WordTalk4x3.pptx";
            // 
            // m_eb169
            // 
            this.m_eb169.Location = new System.Drawing.Point(75, 36);
            this.m_eb169.Name = "m_eb169";
            this.m_eb169.Size = new System.Drawing.Size(335, 20);
            this.m_eb169.TabIndex = 3;
            this.m_eb169.Text = "c:\\temp\\WordTalk16x9.pptx";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 39);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(54, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "16x9 PPT";
            // 
            // m_ebOutput
            // 
            this.m_ebOutput.Location = new System.Drawing.Point(75, 62);
            this.m_ebOutput.Name = "m_ebOutput";
            this.m_ebOutput.Size = new System.Drawing.Size(335, 20);
            this.m_ebOutput.TabIndex = 5;
            this.m_ebOutput.Text = "c:\\temp\\WordTalkOut.pptx";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(13, 65);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(39, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "Output";
            // 
            // m_ebFixLayout
            // 
            this.m_ebFixLayout.Location = new System.Drawing.Point(473, 100);
            this.m_ebFixLayout.Name = "m_ebFixLayout";
            this.m_ebFixLayout.Size = new System.Drawing.Size(75, 23);
            this.m_ebFixLayout.TabIndex = 6;
            this.m_ebFixLayout.Text = "Fix Layout";
            this.m_ebFixLayout.UseVisualStyleBackColor = true;
            this.m_ebFixLayout.Click += new System.EventHandler(this.DoFixup);
            // 
            // m_pbOpen
            // 
            this.m_pbOpen.Location = new System.Drawing.Point(473, 10);
            this.m_pbOpen.Name = "m_pbOpen";
            this.m_pbOpen.Size = new System.Drawing.Size(75, 23);
            this.m_pbOpen.TabIndex = 7;
            this.m_pbOpen.Text = "Open Files";
            this.m_pbOpen.UseVisualStyleBackColor = true;
            this.m_pbOpen.Click += new System.EventHandler(this.DoOpenFiles);
            // 
            // m_cbSlides
            // 
            this.m_cbSlides.FormattingEnabled = true;
            this.m_cbSlides.Location = new System.Drawing.Point(229, 102);
            this.m_cbSlides.Name = "m_cbSlides";
            this.m_cbSlides.Size = new System.Drawing.Size(238, 21);
            this.m_cbSlides.TabIndex = 8;
            // 
            // m_ebFixAll
            // 
            this.m_ebFixAll.Location = new System.Drawing.Point(473, 129);
            this.m_ebFixAll.Name = "m_ebFixAll";
            this.m_ebFixAll.Size = new System.Drawing.Size(75, 23);
            this.m_ebFixAll.TabIndex = 9;
            this.m_ebFixAll.Text = "Fix All Layouts";
            this.m_ebFixAll.UseVisualStyleBackColor = true;
            this.m_ebFixAll.Click += new System.EventHandler(this.DoFixAll);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(565, 164);
            this.Controls.Add(this.m_ebFixAll);
            this.Controls.Add(this.m_cbSlides);
            this.Controls.Add(this.m_pbOpen);
            this.Controls.Add(this.m_ebFixLayout);
            this.Controls.Add(this.m_ebOutput);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.m_eb169);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.m_eb43);
            this.Controls.Add(this.label1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox m_eb43;
        private System.Windows.Forms.TextBox m_eb169;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox m_ebOutput;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button m_ebFixLayout;
        private System.Windows.Forms.Button m_pbOpen;
        private System.Windows.Forms.ComboBox m_cbSlides;
        private System.Windows.Forms.Button m_ebFixAll;
    }
}

