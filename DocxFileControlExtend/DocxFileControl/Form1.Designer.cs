namespace WindowsFormsApplication1
{
    partial class Form1
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージ リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.btnNew = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.txtReplace3 = new System.Windows.Forms.TextBox();
            this.txtReplace2 = new System.Windows.Forms.TextBox();
            this.txtReplace1 = new System.Windows.Forms.TextBox();
            this.btnReplace = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnNew
            // 
            this.btnNew.Location = new System.Drawing.Point(317, 12);
            this.btnNew.Name = "btnNew";
            this.btnNew.Size = new System.Drawing.Size(100, 23);
            this.btnNew.TabIndex = 0;
            this.btnNew.Text = "Docx 新規作成";
            this.btnNew.UseVisualStyleBackColor = true;
            this.btnNew.Click += new System.EventHandler(this.btnNew_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.txtReplace3);
            this.groupBox1.Controls.Add(this.txtReplace2);
            this.groupBox1.Controls.Add(this.txtReplace1);
            this.groupBox1.Controls.Add(this.btnReplace);
            this.groupBox1.Location = new System.Drawing.Point(12, 44);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(423, 114);
            this.groupBox1.TabIndex = 5;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "置換";
            // 
            // txtReplace3
            // 
            this.txtReplace3.Location = new System.Drawing.Point(22, 77);
            this.txtReplace3.Name = "txtReplace3";
            this.txtReplace3.Size = new System.Drawing.Size(249, 19);
            this.txtReplace3.TabIndex = 8;
            this.txtReplace3.Text = "0120-111-2222";
            // 
            // txtReplace2
            // 
            this.txtReplace2.Location = new System.Drawing.Point(22, 52);
            this.txtReplace2.Name = "txtReplace2";
            this.txtReplace2.Size = new System.Drawing.Size(249, 19);
            this.txtReplace2.TabIndex = 7;
            this.txtReplace2.Text = "西原";
            // 
            // txtReplace1
            // 
            this.txtReplace1.Location = new System.Drawing.Point(22, 27);
            this.txtReplace1.Name = "txtReplace1";
            this.txtReplace1.Size = new System.Drawing.Size(249, 19);
            this.txtReplace1.TabIndex = 6;
            this.txtReplace1.Text = "ヤマジュン";
            // 
            // btnReplace
            // 
            this.btnReplace.Location = new System.Drawing.Point(305, 25);
            this.btnReplace.Name = "btnReplace";
            this.btnReplace.Size = new System.Drawing.Size(100, 23);
            this.btnReplace.TabIndex = 5;
            this.btnReplace.Text = "Docx 置換　　　";
            this.btnReplace.UseVisualStyleBackColor = true;
            this.btnReplace.Click += new System.EventHandler(this.btnReplace_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(447, 177);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btnNew);
            this.Name = "Form1";
            this.Text = "Form1";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnNew;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox txtReplace3;
        private System.Windows.Forms.TextBox txtReplace2;
        private System.Windows.Forms.TextBox txtReplace1;
        private System.Windows.Forms.Button btnReplace;
    }
}

