using System;
using System.Drawing;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace exercicio1
{
    class Program
    {
        static void Main(string[] args)
        {
            Document exercicio = new Document();

            Section secao1 = exercicio.AddSection();

            #region 1.1
                Paragraph titulo = secao1.AddParagraph();
                titulo.Format.HorizontalAlignment = HorizontalAlignment.Center;
            #endregion

            #region 1.2
                Paragraph p1_2 = secao1.AddParagraph();
                ParagraphStyle estilo_p1_2 = new ParagraphStyle(exercicio);
                
                estilo_p1_2.Name = "p1_2";
                estilo_p1_2.CharacterFormat.TextColor = Color.Blue;
                estilo_p1_2.CharacterFormat.Bold = true;

                exercicio.Styles.Add(estilo_p1_2);

                p1_2.ApplyStyle(estilo_p1_2.Name);
            #endregion

            #region 1.3
                Table tabela = secao1.AddTable(true);
                String[] cabecalho = {"Nome", "Descrição", "Nome do Vendedor", "Preço"};

                String[][] dados = {
                    new String[]{"Marmita 1", "A primeira marmita da tabela", "Thiago", "R$10,00"},
                    new String[]{"Marmita 2", "A segunda marmita da tabela", "Thiago", "R$10,00"},
                    new String[]{"Marmita 3", "A terceira marmita da tabela", "Thiago", "R$10,00"}
                };

                tabela.ResetCells(dados.Length + 1, cabecalho.Length);

                TableRow linha1 = tabela.Rows[0];
                linha1.IsHeader = true;

                linha1.RowFormat.BackColor = Color.LightGreen;

                for (int i = 0; i < cabecalho.Length; i++)
                {
                    Paragraph p = linha1.Cells[i].AddParagraph();
                    p.Format.HorizontalAlignment = HorizontalAlignment.Center;

                    TextRange tr = p.AppendText(cabecalho[i]);
                    tr.CharacterFormat.Bold = true;
                }

                for (int r = 0; r < dados.Length; r++)
                {
                    TableRow linhaDados = tabela.Rows[r+1];

                    for (int c = 0; c < dados[r].Length; c++)
                    {
                        Paragraph p2 = linhaDados.Cells[c].AddParagraph();
                        TextRange tr2 = p2.AppendText(dados[r][c]);

                        p2.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    }
                }
            #endregion

            Section secao2 = exercicio.AddSection();

            #region 2.1
                Paragraph imagem = secao2.AddParagraph();
                DocPicture imagemPinguins = imagem.AppendPicture(Image.FromFile(@"./saida/img/pinguins.jpeg"));
                imagem.AppendText("Pinguins Consversando de boa...");
                imagem.Format.HorizontalAlignment = HorizontalAlignment.Center;
            #endregion

            #region 2.3
                exercicio.SaveToFile(@"./saida\exercicio.docx", FileFormat.PDF);
            #endregion

            #region 3.1
                exercicio.SaveToFile(@"./saida\exercicio.html", FileFormat.Html);
            #endregion

                // #region 3.2
                //     string paragrafo = "Este é o paragrafo estilizado!";
                    
                //     Paragraph pDesafio = secao2.AddParagraph();

                //     for (int t = 0; t < paragrafo.Length / 3; t++)
                //     {
                //         pDesafio.AppendText(paragrafo[t]);
                //     }
                // #endregion

        }
    }
}
