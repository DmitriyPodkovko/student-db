namespace Students
{
    partial class LovePasswordForm
    {
        /// <summary>
        /// Требуется переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(LovePasswordForm));
            this.txbLove = new System.Windows.Forms.TextBox();
            this.btnLove = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // txbLove
            // 
            this.txbLove.Location = new System.Drawing.Point(12, 12);
            this.txbLove.Name = "txbLove";
            this.txbLove.PasswordChar = '@';
            this.txbLove.Size = new System.Drawing.Size(105, 20);
            this.txbLove.TabIndex = 0;
            // 
            // btnLove
            // 
            this.btnLove.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnLove.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnLove.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.World, ((byte)(204)));
            this.btnLove.Image = ((System.Drawing.Image)(resources.GetObject("btnLove.Image")));
            this.btnLove.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnLove.Location = new System.Drawing.Point(24, 43);
            this.btnLove.Name = "btnLove";
            this.btnLove.Size = new System.Drawing.Size(80, 23);
            this.btnLove.TabIndex = 1;
            this.btnLove.Text = "O K";
            this.btnLove.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnLove.UseVisualStyleBackColor = true;
            this.btnLove.Click += new System.EventHandler(this.btnLove_Click);
            // 
            // LovePasswordForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(129, 78);
            this.Controls.Add(this.btnLove);
            this.Controls.Add(this.txbLove);
            this.Cursor = System.Windows.Forms.Cursors.Default;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.IsMdiContainer = true;
            this.Name = "LovePasswordForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Введіть пароль";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txbLove;
        private System.Windows.Forms.Button btnLove;
    }
}