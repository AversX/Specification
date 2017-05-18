namespace Specification
{
    partial class MainForm
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
            this.FormButton = new System.Windows.Forms.Button();
            this.powerRB = new System.Windows.Forms.RadioButton();
            this.controlRB = new System.Windows.Forms.RadioButton();
            this.SuspendLayout();
            // 
            // FormButton
            // 
            this.FormButton.Location = new System.Drawing.Point(54, 67);
            this.FormButton.Name = "FormButton";
            this.FormButton.Size = new System.Drawing.Size(116, 23);
            this.FormButton.TabIndex = 0;
            this.FormButton.Text = "Сформировать";
            this.FormButton.UseVisualStyleBackColor = true;
            this.FormButton.Click += new System.EventHandler(this.FormButton_Click);
            // 
            // powerRB
            // 
            this.powerRB.AutoSize = true;
            this.powerRB.Location = new System.Drawing.Point(54, 12);
            this.powerRB.Name = "powerRB";
            this.powerRB.Size = new System.Drawing.Size(99, 17);
            this.powerRB.TabIndex = 1;
            this.powerRB.TabStop = true;
            this.powerRB.Text = "Силовая часть";
            this.powerRB.UseVisualStyleBackColor = true;
            // 
            // controlRB
            // 
            this.controlRB.AutoSize = true;
            this.controlRB.Location = new System.Drawing.Point(54, 35);
            this.controlRB.Name = "controlRB";
            this.controlRB.Size = new System.Drawing.Size(87, 17);
            this.controlRB.TabIndex = 1;
            this.controlRB.TabStop = true;
            this.controlRB.Text = "Управление";
            this.controlRB.UseVisualStyleBackColor = true;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(232, 102);
            this.Controls.Add(this.controlRB);
            this.Controls.Add(this.powerRB);
            this.Controls.Add(this.FormButton);
            this.Name = "MainForm";
            this.Text = "Specification";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button FormButton;
        private System.Windows.Forms.RadioButton powerRB;
        private System.Windows.Forms.RadioButton controlRB;
    }
}

