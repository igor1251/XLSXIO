namespace LibTest
{
    partial class TestForm
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
            this.launchTestButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // launchTestButton
            // 
            this.launchTestButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.launchTestButton.Location = new System.Drawing.Point(0, 0);
            this.launchTestButton.Name = "launchTestButton";
            this.launchTestButton.Size = new System.Drawing.Size(204, 55);
            this.launchTestButton.TabIndex = 0;
            this.launchTestButton.Text = "Запустить";
            this.launchTestButton.UseVisualStyleBackColor = true;
            this.launchTestButton.Click += new System.EventHandler(this.launchTestButton_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(204, 55);
            this.Controls.Add(this.launchTestButton);
            this.Name = "Form1";
            this.Text = "XlsxOverview";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button launchTestButton;
    }
}

