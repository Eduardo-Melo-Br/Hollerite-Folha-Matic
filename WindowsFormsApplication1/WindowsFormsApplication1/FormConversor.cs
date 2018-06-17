using System;
using System.IO; 
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ceTe.DynamicPDF;
using ceTe.DynamicPDF.Merger;
using ceTe.DynamicPDF.PageElements;
using ceTe.DynamicPDF.Text;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        Document document;
        ceTe.DynamicPDF.Page page;
        String cab1, cab2, cab3, cab4, cab5;
        String func1, func2, func3, func4, func5, func6;
        String nroBanco, nomBanco, ageBanco, ccBanco, cpfCli;
        Double venc, desc;

        OleDbConnection myConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.; Extended Properties = dBASE IV;"); // Conexão com a Tabela de Dados
        OleDbDataReader _VERBASReader;
        OleDbCommand _VERBASCommand = new OleDbCommand();

        //String sDtAdmissao, sDtPagamento;
        int intLinha = 951; // 851
        int intMargem = 85;
        Boolean bPagina;

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        int intHolerite = 0;

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Telefones para HELP: (19) 3819-2277 ou (19) 99672-9330 E-Mail efmelo@outlookcom ou E-Mail eduardo@MiliuApps.com");
        }

        int intMargemSuperior = 0;

        private void buttonTXTPDFMULTIPLO_Click(object sender, EventArgs e)
        {
            string[] files = {"1","2","3"};

            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                MessageBox.Show("Pasta selecionada: " + folderBrowserDialog1.SelectedPath);
            }
            else
            {
                MessageBox.Show("ERRO");
            }
                
            try
            {
                // Exception could occur due to insufficient permission.
                files = Directory.GetFiles(folderBrowserDialog1.SelectedPath, "Hole*.txt", SearchOption.TopDirectoryOnly);
            }
            catch (Exception)
            {
                MessageBox.Show("Não encontrei arquivos.");
            }

            // If matching files have been found, return the first one.
            if (files.Length > 0)
            {
                for (int i=0;i<=files.Length -1;i++)
                {
                    CriarPDF(files[i]);
                }
            }
        }


        public Form1()
        {
            InitializeComponent();
            myConnection.Open();
        }

        private String format_value(String str1)
        {
            String str2;
            int maxlen = 0;

            str2 = Convert.ToString(Convert.ToDouble(str1.Substring(0, str1.Length)));

            if (str2.Length == 1)
            {
                return "               0,00";
            }

            if (str2.Length > 3)
            {
                //str2 = str2.Substring(0,str2.Length -3) + "," + str2.Substring(str2.Length - 3, 3);
            }
                    
            switch  (str2.Length)
            {
                case 3:
                    maxlen = 6;
                    break;
                case 4:
                    maxlen = 6;
                    break;
                case 5:
                    maxlen = 5;
                    break;
                case 6:
                    maxlen = 4;
                    break;
                case 7:
                    maxlen = 3;
                    break;
                case 8:
                    maxlen = 2;
                    break;
                case 9:
                    maxlen = 1;
                    break;
            }
             
            while (str2.Length < 15 + maxlen)
            {
                str2 = " " + str2;
            }
            
            if (str2.IndexOf(",") < 1 )
            {
                str2 = str2 + ",00";
            }

            if (str2.Length <= str2.IndexOf(",") + 2)
            {
                str2 = str2 + "0";
            }

            return str2;
        }

        private void CriarPaginaDoVerso()
        {
            // Create page to place the PDF
            page = new ceTe.DynamicPDF.Page(1404, 2000, 1);

            ceTe.DynamicPDF.PageElements.Label lbl1 = new ceTe.DynamicPDF.PageElements.Label(T_Unicode(cab1), 1404 -  intMargem, 1940, 700, 35); // Nome do condomínio
            ceTe.DynamicPDF.PageElements.Label lbl6 = new ceTe.DynamicPDF.PageElements.Label(func1, 1404 - intMargem, 1980, 500, 35); // Nome do funcionário
            ceTe.DynamicPDF.PageElements.Label lblTipo = new ceTe.DynamicPDF.PageElements.Label(" ", 700, 35, 400, 50);

            foreach (int indexChecked in checkedListBoxMSGTipoPagamento.CheckedIndices)
            {
                // The indexChecked variable contains the index of the item.
                lblTipo = new ceTe.DynamicPDF.PageElements.Label(checkedListBoxMSGTipoPagamento.Items[indexChecked].ToString() + " Referênte à " + T_Unicode(cab4) + "/" + cab5, 1404 - intMargem, 1900, 700, 35); // Tipo de Pagamento
            }
            lbl1.FontSize = 16;
            lbl6.FontSize = 16;
            lblTipo.FontSize = 16;
            lbl1.Angle = 180;
            lbl6.Angle = 180;
            lblTipo.Angle = 180;
            page.Elements.Add(lbl1);
            page.Elements.Add(lbl6);
            page.Elements.Add(lblTipo);
            document.Pages.Add(page);
        }


        private void CriarNovaPagina()
        {
            // Create page to place the PDF
            page = new ceTe.DynamicPDF.Page(1404, 2100, 1);
            
            intHolerite++;

            this.btnCriar.Text = "Gerei " + Convert.ToString(intHolerite) + " Holerites.";

            // Parte de cima

            // Add rectangles to show dimensions of original          
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(1   + intMargem, 3 + intMargemSuperior,  1160, 220 + intMargemSuperior));              // Primeiro BOX
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(1   + intMargem, 120, 1160, 790));             // BOX do corpo
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Line(50 + intMargem, 200, 50 + intMargem, 810));         // Linha das referências vertical
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Line(1 + intMargem, 810, 890, 810));                     // Linha do final das verbas
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Line(710 + intMargem, 202, 710 + intMargem, 870));       // Linha vertical da referências
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Line(1   + intMargem, 200, 1160 + intMargem, 200));      // Linha Cabeçalho
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(800 + intMargem, 810, 361, 31));               // Box valor liquido
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(800 + intMargem, 200, 150, 670));              // Mensagem Valor Líquido
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(1161 + intMargem,  3, 125, 907));              // Recibo
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(1   + intMargem, 870, 1160, 40));              // Box dos Totais

            // Cabeçalho das verbas

            ceTe.DynamicPDF.PageElements.Label lblVerbas1 = new ceTe.DynamicPDF.PageElements.Label("CÓD.", 5 + intMargem, 203, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblVerbas2 = new ceTe.DynamicPDF.PageElements.Label("DESCRIÇÃO", 370 + intMargem, 203, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblVerbas3 = new ceTe.DynamicPDF.PageElements.Label("REF.", 730 + intMargem, 203, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblVerbas4 = new ceTe.DynamicPDF.PageElements.Label("VENCIMENTOS", 810 + intMargem, 203, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblVerbas5 = new ceTe.DynamicPDF.PageElements.Label("DESCONTOS", 980 + intMargem, 203, 800, 80);

            lblVerbas1.FontSize = 16;
            lblVerbas2.FontSize = 16;
            lblVerbas3.FontSize = 16;
            lblVerbas4.FontSize = 16;
            lblVerbas5.FontSize = 16;

            page.Elements.Add(lblVerbas1);
            page.Elements.Add(lblVerbas2);
            page.Elements.Add(lblVerbas3);
            page.Elements.Add(lblVerbas4);
            page.Elements.Add(lblVerbas5);


            // Suporte   
            ceTe.DynamicPDF.PageElements.Label lblp1 = new ceTe.DynamicPDF.PageElements.Label("MiliuApps.com", 10 + intMargem, 910, 800, 5);
            ceTe.DynamicPDF.PageElements.Label lblp2 = new ceTe.DynamicPDF.PageElements.Label("MiliuApps.com", 10 + intMargem, 910 + intLinha, 800, 5);
            page.Elements.Add(lblp1);
            page.Elements.Add(lblp2);

            // Recibo do Empregador

            ceTe.DynamicPDF.PageElements.Label lblr1 = new ceTe.DynamicPDF.PageElements.Label("DECLARO TER RECEBIDO A IMPORTÂNCIA LIQUÍDA DISCRIMINADA NESTE RECIBO", 1180 + intMargem, 770, 800, 80);
            lblr1.FontSize = 16;

            ceTe.DynamicPDF.PageElements.Label lblRecibo1 = lblr1;
            ceTe.DynamicPDF.PageElements.Label lblRecibo2 = new ceTe.DynamicPDF.PageElements.Label("..................../..................../....................               ..............................................................................................................", 1225 + intMargem, 760, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblRecibo3 = new ceTe.DynamicPDF.PageElements.Label("                         Data                                                                                        Assinatura", 1248 + intMargem, 760, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblRecibo4 = new ceTe.DynamicPDF.PageElements.Label("                             VIA EMPREGADOR", 1262 + intMargem, 730, 800, 80);

            lblRecibo1.Angle = -90;
            lblRecibo2.Angle = -90;
            lblRecibo3.Angle = -90;
            lblRecibo4.Angle = -90;
            lblRecibo4.FontSize = 15;

            page.Elements.Add(lblRecibo1);
            page.Elements.Add(lblRecibo2);
            page.Elements.Add(lblRecibo3);
            page.Elements.Add(lblRecibo4);

            // Recibo do empregado

            ceTe.DynamicPDF.PageElements.Label lblr5 = new ceTe.DynamicPDF.PageElements.Label("DECLARO TER RECEBIDO A IMPORTÂNCIA LIQUÍDA DISCRIMINADA NESTE RECIBO", 1180 + intMargem, 770 + intLinha, 800, 80);
            lblr5.FontSize = 16;

            ceTe.DynamicPDF.PageElements.Label lblRecibo5 = lblr5;
            ceTe.DynamicPDF.PageElements.Label lblRecibo6 = new ceTe.DynamicPDF.PageElements.Label("..................../..................../....................               ..............................................................................................................", 1230 + intMargem, 790 + intLinha, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblRecibo7 = new ceTe.DynamicPDF.PageElements.Label("               Data                                                                                                  Assinatura", 1248 + intMargem, 760 + intLinha, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblRecibo8 = new ceTe.DynamicPDF.PageElements.Label("                             VIA EMPREGADO", 1262 + intMargem, 730 + intLinha, 800, 80);

            lblRecibo5.Angle = -90;
            lblRecibo6.Angle = -90;
            lblRecibo7.Angle = -90;
            lblRecibo8.Angle = -90;
            lblRecibo8.FontSize = 15;

            page.Elements.Add(lblRecibo5);
            page.Elements.Add(lblRecibo6);
            page.Elements.Add(lblRecibo7);
            page.Elements.Add(lblRecibo8);

            // Parte de baixo

            // Add rectangles to show dimensions of original          
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(1 + intMargem, 3 + intLinha, 1160, 220));  // Primeiro BOX
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(1 + intMargem, 120 + intLinha, 1160, 790)); // BOX do corpo
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Line(50 + intMargem, 200 + intLinha, 50 + intMargem, 810 + intLinha)); // Linha das referências vertical
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Line(1 + intMargem, 810 + intLinha, 890, 810 + intLinha));             // Linha do final das verbas 
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Line(710 + intMargem, 202 + intLinha, 710 + intMargem, 870 + intLinha));        // Linha vertical da referências
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Line(1 + intMargem, 200 + intLinha, 1160 + intMargem, 200 + intLinha));         // Linha Cabeçalho
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(800 + intMargem, 810 + intLinha, 361, 31));                                      // Box valor liquido
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(800 + intMargem, 200 + intLinha, 150, 670));              // Mensagem Valor Líquido
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(800 + intMargem, 840 + intLinha, 361, 31));                            // Bases
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(1161 + intMargem, 3 + intLinha, 125, 907));                           // Recibo
            page.Elements.Add(new ceTe.DynamicPDF.PageElements.Rectangle(1 + intMargem, 870 + intLinha, 1160, 40));                            // Box dos Totais
            
            // Cabeçalho das verbas

            ceTe.DynamicPDF.PageElements.Label lblVerbas21 = new ceTe.DynamicPDF.PageElements.Label("CÓD.", 5 + intMargem, 203 + intLinha, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblVerbas22 = new ceTe.DynamicPDF.PageElements.Label("DESCRIÇÃO", 370 + intMargem, 203 + intLinha, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblVerbas23 = new ceTe.DynamicPDF.PageElements.Label("REF.", 730 + intMargem, 203 + intLinha, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblVerbas24 = new ceTe.DynamicPDF.PageElements.Label("VENCIMENTOS", 810 + intMargem, 203 + intLinha, 800, 80);
            ceTe.DynamicPDF.PageElements.Label lblVerbas25 = new ceTe.DynamicPDF.PageElements.Label("DESCONTOS", 980 + intMargem, 203 + intLinha, 800, 80);

            page.Elements.Add(lblVerbas21);
            page.Elements.Add(lblVerbas22);
            page.Elements.Add(lblVerbas23);
            page.Elements.Add(lblVerbas24);
            page.Elements.Add(lblVerbas25);
            
        }

        private string T_Verba(int iVerba)
        {
            
            _VERBASCommand.Connection = myConnection;
            _VERBASCommand.CommandText = String.Concat("SELECT N2 FROM VERBAS WHERE N1 LIKE '", iVerba.ToString().Trim()+"%'");
            _VERBASReader = _VERBASCommand.ExecuteReader();

            if (_VERBASReader.Read())
            {
                string sVerba = _VERBASReader.GetString(0);
                _VERBASReader.Close();
                return sVerba;
            } else
            {
                return "";
            }
        }
        private string T_Unicode(string lbl)
        {
            String nresult = "";

                if (lbl.IndexOf("Condom" + Convert.ToChar(65533) + "nio") >= 0)
                {
                    nresult += "Condomínio ";
                }
                if (lbl.IndexOf("condom" + Convert.ToChar(65533) + "nio") >= 0)
                {
                    nresult += "condomínio ";
                }
                if (lbl.IndexOf("condom" + Convert.ToChar(65533) + "nio") >= 0)
                {
                    nresult += "condomínio ";
                }
                if (lbl.IndexOf("Edif" + Convert.ToChar(65533) + "cio") >= 0)
                {
                    nresult += "Edifício ";
                }
                if (lbl.IndexOf("edif" + Convert.ToChar(65533) + "cio") >= 0)
                {
                    nresult += "edifício ";
                }
                if (lbl.IndexOf("Turia" + Convert.ToChar(65533) + "u") >= 0)
                {
                    nresult += "Turiaçu ";
                }
                // Mêses
            if (lbl.IndexOf("Mar" + Convert.ToChar(65533) + "o") >= 0)
            {
                nresult += "Março ";
            }

            if (nresult.Length == 0)
            {
                nresult = lbl;
            }
            return nresult;
        }
        
        private void CriarPDF(string PDFname)
        {

            // Create a merge document and set it's properties
            document = new Document();
            document.Creator = "Visual Studio 2015";
            document.Author = "Eduardo F. de Melo (19) 3819-2277 / 3504-1629";
            document.Title = "Holerith";

            intHolerite = 0;

            int Linha = 0;

            try
            {
                // Create an instance of StreamReader to read from a file.
                // The using statement also closes the StreamReader.
                StreamReader sr = new StreamReader(PDFname);
                try
                {
                    String line;
                    // Read and display lines from the file until the end of 
                    // the file is reached.
                    while ((line = sr.ReadLine()) != null)
                    {
                        Linha = Linha + 1;

                        // MessageBox.Show(line);


                        if (line.Substring(0, 1) == "1")
                        {
                            bPagina = false;
                            cab1 = line.Substring(1, 50);
                            cab2 = line.Substring(70, 35).Trim() + ", " + line.Substring(108, 114).Trim() + " " + line.Substring(223, 50).Trim() + " " + line.Substring(273, 50).Trim() + " " + line.Substring(323, 2).Trim() + " " + line.Substring(325, 9).Trim();
                            cab3 = line.Substring(50, 20).Trim();
                            cab4 = line.Substring(332, 9);
                            cab5 = line.Substring(341, 4);
                        }

                        if (line.Substring(0, 1) == "2")
                        {
                            venc = 0;
                            desc = 0;
                            CriarNovaPagina();
                            // Parte de cima

                            ceTe.DynamicPDF.PageElements.Label lbl1 = new ceTe.DynamicPDF.PageElements.Label(" ", 700, 35, 400, 50);

                            foreach (int indexChecked in checkedListBoxMSGTipoPagamento.CheckedIndices)
                            {
                                // The indexChecked variable contains the index of the item.
                                lbl1 = new  ceTe.DynamicPDF.PageElements.Label(checkedListBoxMSGTipoPagamento.Items[indexChecked].ToString(), 700, 25, 500, 50);
                            }
                            
                            ceTe.DynamicPDF.PageElements.Label lbl2 = new ceTe.DynamicPDF.PageElements.Label(T_Unicode(cab1), 10 + intMargem, 25, 600, 50); // Nome do condomínio
                            ceTe.DynamicPDF.PageElements.Label lbl3 = new ceTe.DynamicPDF.PageElements.Label(cab2, 10 + intMargem, 50, 750, 50);
                            ceTe.DynamicPDF.PageElements.Label lbl4 = new ceTe.DynamicPDF.PageElements.Label(cab3, 10 + intMargem, 90, 750, 50);
                            ceTe.DynamicPDF.PageElements.Label lbl5 = new ceTe.DynamicPDF.PageElements.Label(T_Unicode(cab4) + "/" + cab5, 1000 + intMargem, 90, 750, 50);
                            // lbl2.Font = ceTe.DynamicPDF.Font.CourierBold;
                            // lbl3.Font = ceTe.DynamicPDF.Font.TimesBold;
                            lbl1.FontSize = 27;
                            lbl2.FontSize = 26;
                            lbl3.FontSize = 22;
                            lbl4.FontSize = 22;
                            lbl5.FontSize = 22;
                            page.Elements.Add(lbl1);
                            page.Elements.Add(lbl2);
                            page.Elements.Add(lbl3);
                            page.Elements.Add(lbl4);
                            page.Elements.Add(lbl5);
                            // Parte de baixo
                            ceTe.DynamicPDF.PageElements.Label lbl21 = new ceTe.DynamicPDF.PageElements.Label(" ", 850, 35, 400, 50);

                            foreach (int indexChecked in checkedListBoxMSGTipoPagamento.CheckedIndices)
                            {
                                // The indexChecked variable contains the index of the item.
                                lbl21 = new ceTe.DynamicPDF.PageElements.Label(checkedListBoxMSGTipoPagamento.Items[indexChecked].ToString(), 700, 20 + intLinha, 500, 50);
                            }
                            
                            lbl21.FontSize = 27;
                            page.Elements.Add(lbl21);
                            ceTe.DynamicPDF.PageElements.Label lbl22 = new ceTe.DynamicPDF.PageElements.Label(T_Unicode(cab1), 10 + intMargem, 20 + intLinha, 600, 50); // Nome do condomínio
                            ceTe.DynamicPDF.PageElements.Label lbl23 = new ceTe.DynamicPDF.PageElements.Label(cab2, 10 + intMargem, 45 + intLinha, 750, 50);
                            ceTe.DynamicPDF.PageElements.Label lbl24 = new ceTe.DynamicPDF.PageElements.Label(cab3, 10 + intMargem, 85 + intLinha, 750, 50);
                            ceTe.DynamicPDF.PageElements.Label lbl25 = new ceTe.DynamicPDF.PageElements.Label(T_Unicode(cab4) + "/" + cab5, 1000 + intMargem, 85 + intLinha, 750, 50);
                            lbl22.FontSize = 26;
                            lbl23.FontSize = 22;
                            lbl24.FontSize = 22;
                            lbl25.FontSize = 22;
                            page.Elements.Add(lbl21);
                            page.Elements.Add(lbl22);
                            page.Elements.Add(lbl23);
                            page.Elements.Add(lbl24);
                            page.Elements.Add(lbl25);
                            // Parte comum
                            func1 = "Funcionário: " + line.Substring(11, 80);
                            func2 = "Cargo: " + line.Substring(111, 50);
                            func3 = "Departamento: ";
                            func4 = "Seção: ";
                            func5 = "Data admissão: " + line.Substring(161, 2) + "/" + line.Substring(163, 2) + "/" + line.Substring(165, 4);
                            func6 = "Data pagamento: " + line.Substring(296, 2) + "/" + line.Substring(298, 2) + "/" + line.Substring(300, 4);
                            // Parte de cima
                            ceTe.DynamicPDF.PageElements.Label lbl6 = new ceTe.DynamicPDF.PageElements.Label(func1, 10 + intMargem, 135, 600, 35);
                            ceTe.DynamicPDF.PageElements.Label lbl7 = new ceTe.DynamicPDF.PageElements.Label(func2, 720 + intMargem, 135, 300, 35);
                            ceTe.DynamicPDF.PageElements.Label lbl8 = new ceTe.DynamicPDF.PageElements.Label(func3, 10 + intMargem, 155, 300, 35);
                            ceTe.DynamicPDF.PageElements.Label lbl9 = new ceTe.DynamicPDF.PageElements.Label(func4, 720 + intMargem, 155, 300, 35);
                            ceTe.DynamicPDF.PageElements.Label lbl10 = new ceTe.DynamicPDF.PageElements.Label(func5, 10 + intMargem, 175, 300, 35);
                            ceTe.DynamicPDF.PageElements.Label lbl11 = new ceTe.DynamicPDF.PageElements.Label(func6, 720 + intMargem, 175, 300, 35);
                            lbl6.FontSize = 22;
                            lbl7.FontSize = 22;
                            lbl8.FontSize = 22;
                            lbl9.FontSize = 22;
                            lbl10.FontSize = 22;
                            lbl11.FontSize = 22;
                            page.Elements.Add(lbl6);
                            page.Elements.Add(lbl7);
                            page.Elements.Add(lbl8);
                            page.Elements.Add(lbl9);
                            page.Elements.Add(lbl10);
                            page.Elements.Add(lbl11);
                            // Parte de baixo
                            ceTe.DynamicPDF.PageElements.Label lbl26 = new ceTe.DynamicPDF.PageElements.Label(func1, 10 + intMargem, 135 + intLinha, 600, 35);
                            ceTe.DynamicPDF.PageElements.Label lbl27 = new ceTe.DynamicPDF.PageElements.Label(func2, 720 + intMargem, 135 + intLinha, 300, 35);
                            ceTe.DynamicPDF.PageElements.Label lbl28 = new ceTe.DynamicPDF.PageElements.Label(func3, 10 + intMargem, 155 + intLinha, 300, 35);
                            ceTe.DynamicPDF.PageElements.Label lbl29 = new ceTe.DynamicPDF.PageElements.Label(func4, 720 + intMargem, 155 + intLinha, 300, 35);
                            ceTe.DynamicPDF.PageElements.Label lbl210 = new ceTe.DynamicPDF.PageElements.Label(func5, 10 + intMargem, 175 + intLinha, 300, 35);
                            ceTe.DynamicPDF.PageElements.Label lbl211 = new ceTe.DynamicPDF.PageElements.Label(func6, 720 + intMargem, 175 + intLinha, 300, 35);
                            lbl26.FontSize = 22;
                            lbl27.FontSize = 22;
                            lbl28.FontSize = 22;
                            lbl29.FontSize = 22;
                            lbl210.FontSize = 22;
                            lbl211.FontSize = 22;
                            page.Elements.Add(lbl26);
                            page.Elements.Add(lbl27);
                            page.Elements.Add(lbl28);
                            page.Elements.Add(lbl29);
                            page.Elements.Add(lbl210);
                            page.Elements.Add(lbl211);
                            // Dados bancários
                            nroBanco = line.Substring(190-1, 3);
                            nomBanco = line.Substring(193-1, 50);
                            ageBanco = line.Substring(243-1, 15);
                            ccBanco = line.Substring(258-1, 15);
                            this.cpfCli = line.Substring(273 - 1, 20);
                            ceTe.DynamicPDF.PageElements.Label lblnroBanco1 = new ceTe.DynamicPDF.PageElements.Label("Banco: " + nroBanco + " Nome: " + nomBanco + " Agência: " + ageBanco + " Conta: " + ccBanco, 30 + intMargem, 835, 800, 35);
                            ceTe.DynamicPDF.PageElements.Label lblnroBanco2 = new ceTe.DynamicPDF.PageElements.Label("Banco: " + nroBanco + " Nome: " + nomBanco + " Agência: " + ageBanco + " Conta: " + ccBanco, 30 + intMargem, 835 + intLinha, 800, 35);
                            ceTe.DynamicPDF.PageElements.Label lblcpf1 = new ceTe.DynamicPDF.PageElements.Label("CPF: " + cpfCli, 30 + intMargem, 850, 800, 35);
                            ceTe.DynamicPDF.PageElements.Label lblcpf2 = new ceTe.DynamicPDF.PageElements.Label("CPF: " + cpfCli, 30 + intMargem, 850 + intLinha, 800, 35);
                            lblnroBanco1.FontSize = 14;
                            lblnroBanco2.FontSize = 14;
                            lblcpf1.FontSize = 14;
                            lblcpf2.FontSize = 14;
                            page.Elements.Add(lblnroBanco1);
                            page.Elements.Add(lblnroBanco2);
                            page.Elements.Add(lblcpf1);
                            page.Elements.Add(lblcpf2);
                        }

                        if (line.Substring(0, 1) == "3")
                        {
                            String v1, v2, v3, v4, v5, v6;

                            v1 = line.Substring(1, 4);
                            v2 = line.Substring(5, 50);
                            v3 = line.Substring(51, 11);
                            v4 = line.Substring(64, 1);
                            v5 = line.Substring(65, 15);
                            v6 = " ";
                            if (v4 == "D")
                            {
                                if (v5.IndexOf("-",0) > 0)
                                {
                                    v5 = v5.Substring(v5.IndexOf("-", 0) + 1);
                                    v5 = "-" + v5;
                                } else
                                {
                                    v5 = v5.Substring(v5.IndexOf("-", 0) + 1);
                                }
                            }
                            v6 = Convert.ToString(Convert.ToDouble(v5) / 100);

                            // Parte de cima
 
                            ceTe.DynamicPDF.PageElements.Label lblv1 = new ceTe.DynamicPDF.PageElements.Label(v1, 2 + intMargem, (Linha * 25) + 165, 400, 35);
                            ceTe.DynamicPDF.PageElements.Label lblv2 = new ceTe.DynamicPDF.PageElements.Label(T_Verba(Convert.ToInt16(v1)), 70 + intMargem, (Linha * 25) + 165, 300, 35);
                            ceTe.DynamicPDF.PageElements.Label lblv3 = new ceTe.DynamicPDF.PageElements.Label(v3, 712 + intMargem, (Linha * 25) + 165, 400, 35);
                            ceTe.DynamicPDF.PageElements.Label lblv4;
                            
                            //OpenTypeFont openTypeFont = new OpenTypeFont("\\Windows\\Fonts\\times.ttf");
                            
                            //lblv2.Font = openTypeFont;

                            if (v4 == "P")
                            {
                                lblv4 = new ceTe.DynamicPDF.PageElements.Label(format_value(v6), 805 + intMargem, (Linha * 25) + 165, 700, 35);
                                venc = venc + Convert.ToDouble(v6);
                            }
                            else
                            {
                                lblv4 = new ceTe.DynamicPDF.PageElements.Label(format_value(v6), 972 + intMargem, (Linha * 25) + 165, 800, 35);
                                desc = desc + Convert.ToDouble(v6);
                            }
                            // Ajustar a fonte
                            lblv1.FontSize = 20;
                            lblv2.FontSize = 20;
                            lblv3.FontSize = 20;
                            lblv4.FontSize = 20;
                            // Adicionar na página
                            page.Elements.Add(lblv1);
                            page.Elements.Add(lblv2);
                            page.Elements.Add(lblv3);
                            page.Elements.Add(lblv4);
                            // Parte de baixo
                            ceTe.DynamicPDF.PageElements.Label lblv21 = new ceTe.DynamicPDF.PageElements.Label(v1, 2 + intMargem, (Linha * 25) + 165 + intLinha, 400, 35);
                            ceTe.DynamicPDF.PageElements.Label lblv22 = new ceTe.DynamicPDF.PageElements.Label(T_Verba(Convert.ToInt16(v1)), 70 + intMargem, (Linha * 25) + 165 + intLinha, 300, 35);
                            ceTe.DynamicPDF.PageElements.Label lblv23 = new ceTe.DynamicPDF.PageElements.Label(v3, 712 + intMargem, (Linha * 25) + 165 + intLinha, 400, 35);
                            ceTe.DynamicPDF.PageElements.Label lblv24;
                            if (v4 == "P")
                            {
                                lblv24 = new ceTe.DynamicPDF.PageElements.Label(format_value(v6), 805 + intMargem, (Linha * 25) + 165 + intLinha, 700, 35);
                            }
                            else
                            {
                                lblv24 = new ceTe.DynamicPDF.PageElements.Label(format_value(v6), 980 + intMargem, (Linha * 25) + 165 + intLinha, 800, 35);
                            }
                            // Ajustar o tamanho da fonte
                            lblv21.FontSize = 20;
                            lblv22.FontSize = 20;
                            lblv23.FontSize = 20;
                            lblv24.FontSize = 20;
                            // Adicionar na página
                            page.Elements.Add(lblv21);
                            page.Elements.Add(lblv22);
                            page.Elements.Add(lblv23);
                            page.Elements.Add(lblv24);

                            // Retangulo

                        }

                        if (line.Substring(0, 1) == "4")
                        {
                            double vlr_liquido = venc - desc;
                            // Parte de cima
                            ceTe.DynamicPDF.PageElements.Label lbltotvenc = new ceTe.DynamicPDF.PageElements.Label(format_value(Convert.ToString(venc)), 720 + intMargem, 820, 800, 35);
                            ceTe.DynamicPDF.PageElements.Label lbltotdesc = new ceTe.DynamicPDF.PageElements.Label(format_value(Convert.ToString(desc)), 900 + intMargem, 820, 800, 35);
                            ceTe.DynamicPDF.PageElements.Label lblmsg1 = new ceTe.DynamicPDF.PageElements.Label("Valor Líquido", 825 + intMargem, 850, 800, 35);
                            ceTe.DynamicPDF.PageElements.Label lbltotliq = new ceTe.DynamicPDF.PageElements.Label(format_value(vlr_liquido.ToString("0.00")), 870 + intMargem, 850, 800, 35);

                            lbltotvenc.Font = ceTe.DynamicPDF.Font.CourierBold;
                            lbltotdesc.Font = ceTe.DynamicPDF.Font.CourierBold;
                            lbltotliq.Font = ceTe.DynamicPDF.Font.CourierBold;
                            lblmsg1.FontSize = 20;

                            ceTe.DynamicPDF.PageElements.Label lblb1 = new ceTe.DynamicPDF.PageElements.Label("Salário Base", 35 + intMargem, 870, 800, 35);
                            ceTe.DynamicPDF.PageElements.Label lblb2 = new ceTe.DynamicPDF.PageElements.Label("Salário Contr. INSS", 180 + intMargem, 870, 800, 35);
                            ceTe.DynamicPDF.PageElements.Label lblb3 = new ceTe.DynamicPDF.PageElements.Label("Base de Calculo FGTS", 380 + intMargem, 870, 800, 35);
                            ceTe.DynamicPDF.PageElements.Label lblb4 = new ceTe.DynamicPDF.PageElements.Label("FGTS do Mês", 580 + intMargem, 870, 800, 35);
                            ceTe.DynamicPDF.PageElements.Label lblb5 = new ceTe.DynamicPDF.PageElements.Label("Base de Calculo IRRF", 745 + intMargem, 870, 800, 35);
                            ceTe.DynamicPDF.PageElements.Label lblb6 = new ceTe.DynamicPDF.PageElements.Label("Dependentes IRRF", 945 + intMargem, 870, 800, 35);

                            // Valores 

                            ceTe.DynamicPDF.PageElements.Label lblv1 = new ceTe.DynamicPDF.PageElements.Label(format_value(Convert.ToString(Convert.ToDouble(line.Substring(46, 15)) / 100)), 35 + intMargem, 885, 800, 35);
                            ceTe.DynamicPDF.PageElements.Label lblv2 = new ceTe.DynamicPDF.PageElements.Label(format_value(Convert.ToString(Convert.ToDouble(line.Substring(61, 15)) / 100)), 180 + intMargem, 885, 800, 35);
                            ceTe.DynamicPDF.PageElements.Label lblv3 = new ceTe.DynamicPDF.PageElements.Label(format_value(Convert.ToString(Convert.ToDouble(line.Substring(76, 15)) / 100)), 380 + intMargem, 885, 800, 35);
                            ceTe.DynamicPDF.PageElements.Label lblv4 = new ceTe.DynamicPDF.PageElements.Label(format_value(Convert.ToString(Convert.ToDouble(line.Substring(91, 15)) / 100)), 580 + intMargem, 885, 800, 35);
                            ceTe.DynamicPDF.PageElements.Label lblv5 = new ceTe.DynamicPDF.PageElements.Label(format_value(Convert.ToString(Convert.ToDouble(line.Substring(106, 15)) / 100)), 780 + intMargem, 885, 800, 35);
                            // Desativado, Número de Dependentesndo IRRF ceTe.DynamicPDF.PageElements.Label lblv6 = new ceTe.DynamicPDF.PageElements.Label(format_value(line.Substring(77,15)), 900, 645, 700, 35);


                            // Aumentar a fonte
                            lblb1.FontSize = 18;
                            lblb2.FontSize = 18;
                            lblb3.FontSize = 18;
                            lblb4.FontSize = 18;
                            lblb5.FontSize = 18;
                            lblb6.FontSize = 18;

                            lblv1.FontSize = 18;
                            lblv2.FontSize = 18;
                            lblv3.FontSize = 18;
                            lblv4.FontSize = 18;
                            lblv5.FontSize = 18;

                            // Totais Vencimento, Descontos e Liquido
                            lbltotvenc.FontSize = 20;
                            lbltotdesc.FontSize = 20;
                            lbltotliq.FontSize = 20;

                            page.Elements.Add(lbltotvenc);
                            page.Elements.Add(lbltotdesc);
                            page.Elements.Add(lblmsg1);
                            page.Elements.Add(lbltotliq);

                            page.Elements.Add(lblb1);
                            page.Elements.Add(lblb2);
                            page.Elements.Add(lblb3);
                            page.Elements.Add(lblb4);
                            page.Elements.Add(lblb5);
                            page.Elements.Add(lblb6);

                            page.Elements.Add(lblv1);
                            page.Elements.Add(lblv2);
                            page.Elements.Add(lblv3);
                            page.Elements.Add(lblv4);
                            page.Elements.Add(lblv5);

                            // page.Elements.Add(lblv6);

                            // Parte de baixo
                            
                            ceTe.DynamicPDF.PageElements.Label lbltotvenc2 = new ceTe.DynamicPDF.PageElements.Label(format_value(Convert.ToString(venc)),720 + intMargem, 820 + intLinha, 800, 35);
                            ceTe.DynamicPDF.PageElements.Label lbltotdesc2 = new ceTe.DynamicPDF.PageElements.Label(format_value(Convert.ToString(desc)), 900 + intMargem, 820 + intLinha, 800, 35);
                            ceTe.DynamicPDF.PageElements.Label lblmsg21 = new ceTe.DynamicPDF.PageElements.Label("Valor Líquido", 825 + intMargem, 850 + intLinha, 800, 35);
                            ceTe.DynamicPDF.PageElements.Label lbltotliq2 = new ceTe.DynamicPDF.PageElements.Label(format_value(  vlr_liquido.ToString("0.00")), 870 + intMargem, 850 + intLinha, 800, 35);

                            lbltotvenc2.Font = ceTe.DynamicPDF.Font.CourierBold;
                            lbltotdesc2.Font = ceTe.DynamicPDF.Font.CourierBold;
                            lbltotliq2.Font = ceTe.DynamicPDF.Font.CourierBold;
                            lblmsg21.FontSize = 20;

                            ceTe.DynamicPDF.PageElements.Label lblb21 = new ceTe.DynamicPDF.PageElements.Label("Salário Base", 35 + intMargem, 870 + intLinha, 800, 35);
                            ceTe.DynamicPDF.PageElements.Label lblb22 = new ceTe.DynamicPDF.PageElements.Label("Salário Contr. INSS", 180 + intMargem, 870 + intLinha, 800, 35);
                            ceTe.DynamicPDF.PageElements.Label lblb23 = new ceTe.DynamicPDF.PageElements.Label("Base de Calculo FGTS", 380 + intMargem, 870 + intLinha, 800, 35);
                            ceTe.DynamicPDF.PageElements.Label lblb24 = new ceTe.DynamicPDF.PageElements.Label("FGTS do Mês", 580 + intMargem, 870 + intLinha, 800, 35);
                            ceTe.DynamicPDF.PageElements.Label lblb25 = new ceTe.DynamicPDF.PageElements.Label("Base de Calculo IRRF", 745 + intMargem, 870 + intLinha, 800, 35);
                            ceTe.DynamicPDF.PageElements.Label lblb26 = new ceTe.DynamicPDF.PageElements.Label("Dependentes IRRF", 945 + intMargem, 870 + intLinha, 800, 35);

                            // Valores 

                            ceTe.DynamicPDF.PageElements.Label lblv21 = new ceTe.DynamicPDF.PageElements.Label(format_value(Convert.ToString(Convert.ToDouble(line.Substring(46, 15)) / 100)), 35 + intMargem, 885 + intLinha, 800, 35);
                            ceTe.DynamicPDF.PageElements.Label lblv22 = new ceTe.DynamicPDF.PageElements.Label(format_value(Convert.ToString(Convert.ToDouble(line.Substring(61, 15)) / 100)), 180 + intMargem, 885 + intLinha, 800, 35);
                            ceTe.DynamicPDF.PageElements.Label lblv23 = new ceTe.DynamicPDF.PageElements.Label(format_value(Convert.ToString(Convert.ToDouble(line.Substring(76, 15)) / 100)), 380 + intMargem, 885 + intLinha, 800, 35);
                            ceTe.DynamicPDF.PageElements.Label lblv24 = new ceTe.DynamicPDF.PageElements.Label(format_value(Convert.ToString(Convert.ToDouble(line.Substring(91, 15)) / 100)), 580 + intMargem, 885 + intLinha, 800, 35);
                            ceTe.DynamicPDF.PageElements.Label lblv25 = new ceTe.DynamicPDF.PageElements.Label(format_value(Convert.ToString(Convert.ToDouble(line.Substring(106, 15)) / 100)), 780 + intMargem, 885 + intLinha, 800, 35);
                            // Desativado ceTe.DynamicPDF.PageElements.Label lblv26 = new ceTe.DynamicPDF.PageElements.Label(format_value(Convert.ToString(Convert.ToDouble(line.Substring(77, 15))/100)), 900, 645 + intLinha, 700, 35);

                            // Aumentar a fonte
                            lblb21.FontSize = 18;
                            lblb22.FontSize = 18;
                            lblb23.FontSize = 18;
                            lblb24.FontSize = 18;
                            lblb25.FontSize = 18;
                            lblb26.FontSize = 18;

                            lblv21.FontSize = 18;
                            lblv22.FontSize = 18;
                            lblv23.FontSize = 18;
                            lblv24.FontSize = 18;
                            lblv25.FontSize = 18;

                            // Totais Vencimento, Descontos e Liquido
                            lbltotvenc2.FontSize = 20;
                            lbltotdesc2.FontSize = 20;
                            lbltotliq2.FontSize = 20;

                            page.Elements.Add(lbltotvenc2);
                            page.Elements.Add(lbltotdesc2);
                            page.Elements.Add(lblmsg21);
                            page.Elements.Add(lbltotliq2);

                            page.Elements.Add(lblb21);
                            page.Elements.Add(lblb22);
                            page.Elements.Add(lblb23);
                            page.Elements.Add(lblb24);
                            page.Elements.Add(lblb25);
                            page.Elements.Add(lblb26);

                            page.Elements.Add(lblv21);
                            page.Elements.Add(lblv22);
                            page.Elements.Add(lblv23);
                            page.Elements.Add(lblv24);
                            page.Elements.Add(lblv25);

                        }
                        if (line.Substring(0, 1) == "5")
                        {

                            String strOBS;
                            bPagina = true;

                            strOBS = line.Substring(1, line.Length - 1);
                            
                            // Parte de cima
                            ceTe.DynamicPDF.PageElements.Label lblOBS = new ceTe.DynamicPDF.PageElements.Label(strOBS, 10 + intMargem, 820, 800, 65);

                            // Parte de baixo
                            ceTe.DynamicPDF.PageElements.Label lblOBS2 = new ceTe.DynamicPDF.PageElements.Label(strOBS, 10 + intMargem, 820 + intLinha, 800, 65);

                            page.Elements.Add(lblOBS);
                            page.Elements.Add(lblOBS2);

                            // Add page to document
                            document.Pages.Add(page);
                            Linha = 1;

                            CriarPaginaDoVerso();
                        }
                    }
                }
                catch (System.Exception eee)
                {
                    MessageBox.Show(eee.Message.ToString());
                }
                finally
                {
                    sr.Dispose();
                }
            }
            catch (System.Exception ee)
            {
                // Let the user know what went wrong.
                Console.WriteLine("The file could not be read:");
                Console.WriteLine(ee.Message.ToString());
            }


            PDFname = PDFname.Substring(0, PDFname.Length - 4);

            // Outputs

            if (!bPagina) {
                document.Pages.Add(page);
            }

            document.Draw(PDFname + ".pdf");

            MessageBox.Show("Gerei o arquivo " + PDFname + ".pdf");
            
        }

        private void btnCriar_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileBrowserDialog1 = new OpenFileDialog();
            if (fileBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                MessageBox.Show("Arquivo selecionado: " + fileBrowserDialog1.FileName);
            }
            else
            {
                MessageBox.Show("ERRO");
            }
            CriarPDF(fileBrowserDialog1.FileName);
        }

        private class CheckedItensColletion
        {
        }
    }
}