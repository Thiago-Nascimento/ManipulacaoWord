using System;
using System.Drawing;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace exemplo1
{
    class Program
    {
        static void Main(string[] args)
        {
            #region CriacaoDoDocumento
                // Criando documento com nome "exemplo"
                Document exemplo = new Document();                
            #endregion

            #region CriaçãoDeSecao
                // Adiciona uma secao com nome "secaoCapa", a secao pode ser entendida como uma pagina
                Section secaoCapa = exemplo.AddSection();
            #endregion

            #region CriarParagrafo
                // Cria paragrafo "titulo" e o adiciona na "secaoCapa"
                Paragraph titulo = secaoCapa.AddParagraph();
            #endregion

            #region AdicionadoTextoNoParagrafo
                titulo.AppendText("Titulo do Documento\n\a");
            #endregion

            #region FormatandoParagrafo
                // Formatando o paragrafo titulo, com alinhamento central
                titulo.Format.HorizontalAlignment = HorizontalAlignment.Center;

                // Cria um estilo de formatação e adiciona ao documento
                ParagraphStyle estilo1 = new ParagraphStyle(exemplo);

                // Define o nome do estilo
                estilo1.Name = "Cor do Título";

                // Define a cor do titulo
                estilo1.CharacterFormat.TextColor = Color.Red;

                // Define Bold no estilo1
                estilo1.CharacterFormat.Bold = true;

                // Adiciona o estilo ao documento
                exemplo.Styles.Add(estilo1);

                // Aplica o estilo ao paragrafo
                titulo.ApplyStyle(estilo1.Name);
            #endregion

            #region Tabulação
                Paragraph textoDaCapa = secaoCapa.AddParagraph();
                textoDaCapa.AppendText("\tParagrafo Tabulado\n\a");

                Paragraph outroTexto = secaoCapa.AddParagraph();
                outroTexto.AppendText("\tTexto Adicionadao a outroTexto " + "E aparecem na mesma página");
            #endregion

            #region InsercaoDeImagens
                Paragraph imagemDaCapa = secaoCapa.AddParagraph();
                imagemDaCapa.AppendText("\n\n\tAgora vamos inserir uma imagem\n\n");
                imagemDaCapa.Format.HorizontalAlignment = HorizontalAlignment.Center;

                // Adiciona uma imagem "imagemExemplo" no paragrafo "imagemDaCapa"
                DocPicture imagemExemplo = imagemDaCapa.AppendPicture(Image.FromFile(@"./saida/img/logo_csharp.png"));

                // Largura e altura
                imagemExemplo.Width = 300;
                imagemExemplo.Height = 300;                
            #endregion

            #region AdicionarNovaSecao
                Section secaoCorpo = exemplo.AddSection();

                Paragraph paragrafoCorpo1 = secaoCorpo.AddParagraph();
                paragrafoCorpo1.AppendText("\tParagrafo criado na secao corpo" + "\tComo foi criada uma nova seção, perceba que este texto aparece em uma nova página");                
            #endregion

            #region Tabela
                Table tabela = secaoCorpo.AddTable(true);

                String[] cabecalho = {"Item", "Descrição", "Quantidade", "Preço Unitário", "Preço Total"};

                // Cria uma tabela com dados já inseridos
                String[][] dados = {
                    new String[]{"Cenoura", "Vegetal muito consumido", "1", "R$4,00", "R$4,00"},
                    new String[]{"Batata", "Vegetal muito consumido", "2", "R$5,00", "R$10,00"},
                    new String[]{"Alface", "Vegetal utilizado desde 500 a.C.", "1", "R$1,50", "R$1,50"},
                    new String[]{"Tomate", "Tomate é uma fruta", "2", "R$6,00", "R$12,00"}, 
                };

                // Adiciona as celulas na tabela
                tabela.ResetCells(dados.Length + 1, cabecalho.Length);

                // Adiciona uma linha na posição 0 do vetor de linhas
                // E define que esta linha é um cabecalho
                TableRow linha1 = tabela.Rows[0];
                linha1.IsHeader = true;

                linha1.Height = 23;

                linha1.RowFormat.BackColor = Color.AliceBlue;

                for (int i = 0; i < cabecalho.Length; i++)
                {
                    Paragraph p = linha1.Cells[i].AddParagraph();
                    linha1.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    p.Format.HorizontalAlignment = HorizontalAlignment.Center;

                    // Formatação dos dados do cabecalho
                    TextRange tr = p.AppendText(cabecalho[i]);
                    tr.CharacterFormat.FontName = "Calibri";
                    tr.CharacterFormat.FontSize = 14;
                    tr.CharacterFormat.TextColor = Color.Teal;
                    tr.CharacterFormat.Bold = true;
                }

                // Adiciona as linhas do corpo da tabela
                for (int r = 0; r < dados.Length; r++)
                {
                    TableRow linhaDados = tabela.Rows[r+1];
                    linhaDados.Height = 20;

                    for (int c = 0; c < dados[r].Length; c++)
                    {
                        linhaDados.Cells[c].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                        
                        Paragraph p2 = linhaDados.Cells[c].AddParagraph();
                        TextRange tr2 = p2.AppendText(dados[r][c]);

                        p2.Format.HorizontalAlignment = HorizontalAlignment.Center;
                        tr2.CharacterFormat.FontName = "Calibri";
                        tr2.CharacterFormat.FontSize = 12;
                        tr2.CharacterFormat.TextColor = Color.Brown;
                    }
                }
            #endregion

            #region SalvarArquivo
                exemplo.SaveToFile(@"./saida\exemploArquivoWord.docx", FileFormat.Docx);
            #endregion
        }
    }
}
