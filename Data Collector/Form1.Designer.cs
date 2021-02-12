
namespace Data_Collector
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
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
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.grab_btn = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // grab_btn
            // 
            this.grab_btn.Location = new System.Drawing.Point(53, 33);
            this.grab_btn.Name = "grab_btn";
            this.grab_btn.Size = new System.Drawing.Size(110, 26);
            this.grab_btn.TabIndex = 0;
            this.grab_btn.Text = "Собрать данные";
            this.grab_btn.UseVisualStyleBackColor = true;
            this.grab_btn.Click += new System.EventHandler(this.grab_btn_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(228, 94);
            this.Controls.Add(this.grab_btn);
            this.Name = "Form1";
            this.Text = "Data Collector";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button grab_btn;
    }
}

