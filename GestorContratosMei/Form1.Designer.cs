
namespace GestorContratosMei
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
            this.openFDExcel = new System.Windows.Forms.OpenFileDialog();
            this.saveFDExcel = new System.Windows.Forms.SaveFileDialog();
            this.textBoxExcel = new System.Windows.Forms.TextBox();
            this.buttonOpenExcel = new System.Windows.Forms.Button();
            this.buttonExecutar = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.checkedListBoxColunas = new System.Windows.Forms.CheckedListBox();
            this.SuspendLayout();
            // 
            // textBoxExcel
            // 
            this.textBoxExcel.Enabled = false;
            this.textBoxExcel.Location = new System.Drawing.Point(28, 113);
            this.textBoxExcel.Name = "textBoxExcel";
            this.textBoxExcel.Size = new System.Drawing.Size(386, 23);
            this.textBoxExcel.TabIndex = 0;
            this.textBoxExcel.TextChanged += new System.EventHandler(this.textBoxExcel_TextChanged);
            // 
            // buttonOpenExcel
            // 
            this.buttonOpenExcel.Location = new System.Drawing.Point(28, 57);
            this.buttonOpenExcel.Name = "buttonOpenExcel";
            this.buttonOpenExcel.Size = new System.Drawing.Size(75, 23);
            this.buttonOpenExcel.TabIndex = 1;
            this.buttonOpenExcel.Text = "Open";
            this.buttonOpenExcel.UseVisualStyleBackColor = true;
            this.buttonOpenExcel.Click += new System.EventHandler(this.buttonOpenExcel_Click);
            // 
            // buttonExecutar
            // 
            this.buttonExecutar.Location = new System.Drawing.Point(28, 361);
            this.buttonExecutar.Name = "buttonExecutar";
            this.buttonExecutar.Size = new System.Drawing.Size(75, 23);
            this.buttonExecutar.TabIndex = 2;
            this.buttonExecutar.Text = "Executar";
            this.buttonExecutar.UseVisualStyleBackColor = true;
            this.buttonExecutar.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(28, 212);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(169, 15);
            this.label1.TabIndex = 5;
            this.label1.Text = "Selecione um filtro de  colunas";
            // 
            // checkedListBoxColunas
            // 
            this.checkedListBoxColunas.FormattingEnabled = true;
            this.checkedListBoxColunas.Location = new System.Drawing.Point(508, 27);
            this.checkedListBoxColunas.Name = "checkedListBoxColunas";
            this.checkedListBoxColunas.Size = new System.Drawing.Size(215, 364);
            this.checkedListBoxColunas.TabIndex = 6;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.checkedListBoxColunas);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.buttonExecutar);
            this.Controls.Add(this.buttonOpenExcel);
            this.Controls.Add(this.textBoxExcel);
            this.Name = "Form1";
            this.Text = "Importador Excel";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFDExcel;
        private System.Windows.Forms.SaveFileDialog saveFDExcel;
        private System.Windows.Forms.TextBox textBoxExcel;
        private System.Windows.Forms.Button buttonOpenExcel;
        private System.Windows.Forms.Button buttonExecutar;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckedListBox checkedListBoxColunas;
    }
}

