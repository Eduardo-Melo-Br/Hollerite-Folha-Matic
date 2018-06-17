namespace WindowsFormsApplication1
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
            this.btnCriar = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.buttonTXTPDFMULTIPLO = new System.Windows.Forms.Button();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.checkedListBoxMSGTipoPagamento = new System.Windows.Forms.CheckedListBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.comboBoxMes = new System.Windows.Forms.ComboBox();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnCriar
            // 
            this.btnCriar.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.btnCriar.Location = new System.Drawing.Point(15, 27);
            this.btnCriar.Name = "btnCriar";
            this.btnCriar.Size = new System.Drawing.Size(351, 68);
            this.btnCriar.TabIndex = 0;
            this.btnCriar.Text = "Para criar o arquivo Holerite.PDF clique aqui!";
            this.btnCriar.UseVisualStyleBackColor = true;
            this.btnCriar.Click += new System.EventHandler(this.btnCriar_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 347);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(167, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Licença para uso de Zelo Imóveis";
            // 
            // buttonTXTPDFMULTIPLO
            // 
            this.buttonTXTPDFMULTIPLO.Location = new System.Drawing.Point(15, 101);
            this.buttonTXTPDFMULTIPLO.Name = "buttonTXTPDFMULTIPLO";
            this.buttonTXTPDFMULTIPLO.Size = new System.Drawing.Size(351, 47);
            this.buttonTXTPDFMULTIPLO.TabIndex = 2;
            this.buttonTXTPDFMULTIPLO.Text = "Criar Hole*.pdf para Hole*.txt";
            this.buttonTXTPDFMULTIPLO.UseVisualStyleBackColor = true;
            this.buttonTXTPDFMULTIPLO.Click += new System.EventHandler(this.buttonTXTPDFMULTIPLO_Click);
            // 
            // folderBrowserDialog1
            // 
            this.folderBrowserDialog1.SelectedPath = "C:\\Holerites";
            this.folderBrowserDialog1.ShowNewFolderButton = false;
            // 
            // checkedListBoxMSGTipoPagamento
            // 
            this.checkedListBoxMSGTipoPagamento.FormattingEnabled = true;
            this.checkedListBoxMSGTipoPagamento.Items.AddRange(new object[] {
            "Recibo de adiantamento salarial",
            "Recibo de salário",
            "Recibo de adiantamento de 13º salário",
            "Recibo de 13º salário"});
            this.checkedListBoxMSGTipoPagamento.Location = new System.Drawing.Point(15, 175);
            this.checkedListBoxMSGTipoPagamento.Name = "checkedListBoxMSGTipoPagamento";
            this.checkedListBoxMSGTipoPagamento.Size = new System.Drawing.Size(490, 64);
            this.checkedListBoxMSGTipoPagamento.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 156);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(99, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Tipo de pagamento";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(13, 242);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(27, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "Mês";
            // 
            // comboBoxMes
            // 
            this.comboBoxMes.FormattingEnabled = true;
            this.comboBoxMes.Items.AddRange(new object[] {
            "Janeiro",
            "Fevereiro",
            "Março",
            "Abril",
            "Maio",
            "Junho",
            "Julho",
            "Agosto",
            "Setembro",
            "Outubro",
            "Novembro",
            "Dezembro"});
            this.comboBoxMes.Location = new System.Drawing.Point(16, 259);
            this.comboBoxMes.Name = "comboBoxMes";
            this.comboBoxMes.Size = new System.Drawing.Size(121, 21);
            this.comboBoxMes.TabIndex = 6;
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            "2015",
            "2016",
            "2017",
            "2018",
            "2019",
            "2020",
            "2021"});
            this.comboBox1.Location = new System.Drawing.Point(144, 258);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(121, 21);
            this.comboBox1.TabIndex = 7;
            this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(144, 241);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(26, 13);
            this.label4.TabIndex = 8;
            this.label4.Text = "Ano";
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.aboutToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(517, 24);
            this.menuStrip1.TabIndex = 9;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(44, 20);
            this.aboutToolStripMenuItem.Text = "Help";
            this.aboutToolStripMenuItem.Click += new System.EventHandler(this.aboutToolStripMenuItem_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(517, 394);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.comboBoxMes);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.checkedListBoxMSGTipoPagamento);
            this.Controls.Add(this.buttonTXTPDFMULTIPLO);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnCriar);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "Conversor TXT para PDF ";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnCriar;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button buttonTXTPDFMULTIPLO;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.CheckedListBox checkedListBoxMSGTipoPagamento;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox comboBoxMes;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
    }
}

