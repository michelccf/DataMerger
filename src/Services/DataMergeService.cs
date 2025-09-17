using ClosedXML.Excel;
using DataMerger.DTOs;
using DataMerger.Interfaces;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataMerger.Services
{
    public class DataMergeService : IDataMergeService
    {
        public DataMergeService() {}

        public async Task StartProcess()
        {
            try
            {
                List<LinhaGenerica> plan1 = LerPlanilha("DIRF.xlsx");
                Console.WriteLine("DIRF.xlsx lida.");

                List<LinhaGenerica> plan2 = LerPlanilha("COPARTICIPAÇAO.xlsx");
                Console.WriteLine("OPARTICIPAÇAO.xlsx lida.");

                List<Empresa> emp = LerPlanilhaCnpj("CNPJS.xlsx");
                Console.WriteLine("CNPJS.xlsx lida.");

                plan1 = PreencherDadosAdicionais(plan1);

                List<PlanilhaFinal> result = MontarResult(plan1, plan2, emp);

                // Definição de ordem das colunas na planilha final.
                string[] colunasDesejadas = new[] { "Matricula", "Titular", "CPF Titular", "Data Nascimento Titular", "Nome do Beneficiario", "CPF Beneficiario", "Data de Nascimento Beneficiario", "Valor", "Valor total", "Subfatura", "Data de referencia", "Nome do Prestador", "Cnpj prestador", "Valor reemboso anos anteriores" };
                Console.WriteLine("Definindo Cabeçalho da planilha final.");

                Console.WriteLine("Gerando planilha final.");
                GerarPlanilhaFinal(result, colunasDesejadas, $"Planilha Coparticipacao {DateTime.Now.Month}.{DateTime.Now.Year}.xlsx");
                Console.WriteLine("Planilha final gerada com sucesso.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro no processo, mensagem: {ex.Message}{Environment.NewLine}StackTrace:{ex.StackTrace}{Environment.NewLine}");
                Console.ReadLine();
            }

        }

        private List<LinhaGenerica> PreencherDadosAdicionais(List<LinhaGenerica> plan1)
        {
            foreach (var x in plan1.GroupBy(_ => _.Nome))
            {
                var p = plan1.FirstOrDefault(_ => _.Nome == x.Key);
                p.QtdDependentes = x.Count();
                if (p.QtdDependentes > 1)
                {
                    p.Total = x.SelectMany(_ => _.Colunas.Where(_ => _.Key == "Valor Participação")).Sum(_ => Convert.ToDecimal(_.Value)).ToString().Replace(".", ",");
                }
            }

            return plan1;
        }

        private List<PlanilhaFinal> MontarResult(List<LinhaGenerica> Dirf, List<LinhaGenerica> Cop, List<Empresa> emp)
        {
            List<PlanilhaFinal> result = new List<PlanilhaFinal>();
            foreach (var d in Dirf)
            {
                var cop = Cop.Where(_ => _.Nome == d.Nome)?.FirstOrDefault();
                PlanilhaFinal x = new PlanilhaFinal();
                x.Matricula = cop?.Colunas.Where(_ => _.Key == "MATRICULA ESPECIAL")?.FirstOrDefault().Value ?? string.Empty; // Onde pego?
                x.Titular = d.Nome;
                x.CPF_Titular = d.Colunas.Where(_ => _.Key == "CPF Titular").FirstOrDefault().Value;
                x.Data_Nascimento_Titular = string.Empty; // Onde pego?
                x.Nome_Do_Beneficiario = d.Colunas.Where(_ => _.Key == "Nome Dependente").FirstOrDefault().Value;

                if (x.Nome_Do_Beneficiario == string.Empty)
                    x.Nome_Do_Beneficiario = x.Titular;

                x.CPF_Beneficiario = d.Colunas.Where(_ => _.Key == "CPF Dependente").FirstOrDefault().Value;

                if (x.CPF_Beneficiario == string.Empty)
                    x.CPF_Beneficiario = x.CPF_Titular;

                x.Data_De_Nascimento_Beneficiario = d.Colunas.Where(_ => _.Key == "Data Nascimento Dependente").FirstOrDefault().Value;

                if (string.IsNullOrEmpty(x.Data_De_Nascimento_Beneficiario))
                    x.Data_De_Nascimento_Beneficiario = x.Data_Nascimento_Titular;

                x.Valor = d.Colunas.Where(_ => _.Key == "Valor Participação").FirstOrDefault().Value;
                x.Valor_Total = Convert.ToDecimal(d.Total) > 0 ? d.Total : "";
                x.Subfatura = cop?.Colunas.Where(_ => _.Key == "NUMERO DA SUBFATURA").FirstOrDefault().Value;

                x.Data_De_referencia = $"01/{(DateTime.Now.Month -2).ToString().PadLeft(2, '0')}/{DateTime.Now.Year}"; 

                x.NomeEmpresa = emp.Where(_ => _.Subfatura == Convert.ToInt32(x.Subfatura)).FirstOrDefault().Emp;
                x.Cnpj_Prestador = emp.Where(_ => _.Subfatura == Convert.ToInt32(x.Subfatura)).FirstOrDefault().Cnpj; // Onde pego?
                x.Valor_reemboso_anos_anteriores = string.Empty; // Onde pego?
                x.QtdDependentes = d.QtdDependentes;

                if (x.Matricula == null)
                    x.Matricula = "";
                result.Add(x);
            }
            return result;
        }

        private List<LinhaGenerica> LerPlanilha(string caminho)
        {
            var lista = new List<LinhaGenerica>();

            using XLWorkbook wb = new XLWorkbook(caminho);
            IXLWorksheet ws = wb.Worksheet(1); // primeira aba
            List<IXLRangeRow> rows = ws.RangeUsed().RowsUsed().ToList();

            if (!rows.Any()) return lista;

            

            int skip = 0;
            int columnNome = 0;

            if (caminho == "DIRF.xlsx")
            {
                skip = 3;
                columnNome = 3;
            }
            if (caminho == "COPARTICIPAÇAO.xlsx")
            {
                skip = 3;
                columnNome = 5;
            }
            // Pegar o cabeçalho
            List<string> header = rows[skip - 1].Cells().Where(_ => _.GetString() != string.Empty).Select(c => c.GetString()).ToList();


            // percorrer linhas
            foreach (IXLRangeRow row in rows.Skip(skip))
            {
                string nome = string.Empty;
                //Pega o nome do cliente refferênte à linha em questão
                nome = row.Cell(columnNome).GetString();
                if (string.IsNullOrWhiteSpace(nome)) continue;

                LinhaGenerica linha = new LinhaGenerica { Nome = nome };

                for (int i = 0; i < header.Count; i++)
                {
                    //Atribrui o valor na linha referente à coluna
                   
                    linha.Colunas[header[i]] = row.Cell(i + 1).GetString();

                }

                lista.Add(linha);
            }

            return lista;
        }

        private List<Empresa> LerPlanilhaCnpj(string caminho)
        {
            var lista = new List<Empresa>();

            using XLWorkbook wb = new XLWorkbook(caminho);
            IXLWorksheet ws = wb.Worksheet(1); // primeira aba
            List<IXLRangeRow> rows = ws.RangeUsed().RowsUsed().ToList();

            if (!rows.Any()) return lista;

            // Pegar o cabeçalho
            List<string> header = rows.First().Cells().Select(c => c.GetString()).ToList();

            // percorrer linhas
            foreach (IXLRangeRow row in rows.Skip(1))
            {
                Empresa linha = new Empresa();

                for (int i = 0; i < header.Count; i++)
                {
                    switch(i)
                    {
                        case 0: 
                            linha.Subfatura = Convert.ToInt32(row.Cell(i + 1).GetString());
                        break;
                        case 1:
                            linha.Cnpj = row.Cell(i + 1).GetString();
                        break;
                        case 2:
                            linha.Emp = row.Cell(i + 1).GetString();
                        break;
                    }
                }
                lista.Add(linha);
            }
            return lista;
        }

        private void GerarPlanilhaFinal(List<PlanilhaFinal> dados, string[] colunasDesejadas, string caminhoSaida)
        {
            using XLWorkbook wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Consolidado");

            // Cabeçalho
            for (int i = 0; i < colunasDesejadas.Length; i++)
                ws.Cell(1, i + 1).Value = colunasDesejadas[i];

            int row = 2;

            dados = dados.OrderBy(_ => _.Subfatura).ThenBy(_ => _.Titular).ToList();

            foreach (var linha in dados)
            {
                for (int i = 0; i < colunasDesejadas.Length; i++)
                {
                    switch (i)
                    {
                        case 0:
                            ws.Cell(row, i + 1).Value = linha.Matricula;
                            break;
                        case 1:
                            ws.Cell(row, i + 1).Value = linha.Titular;
                            break;
                        case 2:
                            ws.Cell(row, i + 1).Value = linha.CPF_Titular;
                            break;
                        case 3:
                            ws.Cell(row, i + 1).Value = linha.Data_Nascimento_Titular;
                            break;
                        case 4:
                            ws.Cell(row, i + 1).Value = linha.Nome_Do_Beneficiario;
                            break;
                        case 5:
                            ws.Cell(row, i + 1).Value = linha.CPF_Beneficiario;
                            break;
                        case 6:

                            if (!string.IsNullOrEmpty(linha.Data_De_Nascimento_Beneficiario) && !linha.Data_De_Nascimento_Beneficiario.Contains("/"))
                            {
                                DateTime baseDate = new DateTime(1900, 1, 1);
                                DateTime dt = baseDate.AddDays(Convert.ToDouble(linha.Data_De_Nascimento_Beneficiario) - 2);

                                ws.Cell(row, i + 1).Value = $"{dt.Day.ToString().PadLeft(2, '0')}/{dt.Month.ToString().PadLeft(2, '0')}/{dt.Year}";
                            }
                            else
                                ws.Cell(row, i + 1).Value = linha.Data_De_Nascimento_Beneficiario;

                            break;
                        case 7:
                            ws.Cell(row, i + 1).Value = linha.Valor;
                            break;
                        case 8:
                            if (linha.Valor_Total != string.Empty)
                            {
                                var ex = ws.Range(row, i + 1, (row + linha.QtdDependentes - 1), i + 1);
                                ex.Merge();
                                ex.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                ex.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                                ws.Cell(row, i + 1).Value = linha.Valor_Total;
                            }
                                break;
                        case 9:
                            ws.Cell(row, i + 1).Value = linha.Subfatura;
                            break;
                        case 10:
                            ws.Cell(row, i + 1).Value = linha.Data_De_referencia;
                            break;
                        case 11:
                            ws.Cell(row, i + 1).Value = linha.NomeEmpresa;
                            break;
                        case 12:
                            ws.Cell(row, i + 1).Value = linha.Cnpj_Prestador;
                            break;
                        case 13:
                            ws.Cell(row, i + 1).Value = linha.Nome_Do_Prestador;
                            break;
                        case 14:
                            ws.Cell(row, i + 1).Value = linha.Valor_reemboso_anos_anteriores;
                            break;
                    }
                    
                }
                row++;
            }

            ws = AplicarEstilos(ws);

            ws.Columns().AdjustToContents();
            wb.SaveAs(caminhoSaida);
        }

        private IXLWorksheet AplicarEstilos(IXLWorksheet ws)
        {
            int ultimaLinha = ws.LastRowUsed().RowNumber();
            int ultimaColuna = ws.LastColumnUsed().ColumnNumber();

            for (int r = 1; r <= ultimaLinha; r++)
            {
                for (int c = 1; c <= ultimaColuna; c++)
                {

                    var cell = ws.Cell(r, c);

                    if (r == 1)
                    {

                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        cell.Style.Font.Bold = true;
                        cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#dcdcdc");
                    }                  

                    cell.Style.Font.FontName = "Calibri";
                    cell.Style.Font.FontSize = 11;

                    cell.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                    cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    cell.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                    cell.Style.Border.RightBorder = XLBorderStyleValues.Thin;

                    cell.Style.Border.TopBorderColor = XLColor.Black;
                    cell.Style.Border.BottomBorderColor = XLColor.Black;
                    cell.Style.Border.LeftBorderColor = XLColor.Black;
                    cell.Style.Border.RightBorderColor = XLColor.Black;
                }
            }
            return ws;
        }
    }
}
