namespace CSharpColorPicker
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }
        #region Windows Form Designer generated code
        private void InitializeComponent()
        {
            this.excelColorPopupButton1 = new CSharpColorPicker.ExcelColorPopupButton();
            this.SuspendLayout();
            this.excelColorPopupButton1.Location = new System.Drawing.Point(313, 158);
            this.excelColorPopupButton1.Name = "excelColorPopupButton1";
            this.excelColorPopupButton1.SelectedColor = System.Drawing.Color.White;
            this.excelColorPopupButton1.Size = new System.Drawing.Size(47, 25);
            this.excelColorPopupButton1.TabIndex = 0;
            this.excelColorPopupButton1.UseVisualStyleBackColor = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.excelColorPopupButton1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
        }
        #endregion
        private ExcelColorPopupButton excelColorPopupButton1;
    }
}
