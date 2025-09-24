using ClosedXML.Excel;
using DataMerger.DTOs;
using DataMerger.Interfaces;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Numerics;
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
                Console.WriteLine($"Digite:{Environment.NewLine}1 - Coparticipação{Environment.NewLine}2 - Saúde{Environment.NewLine}3 - Dental{Environment.NewLine}");
                var key = Console.ReadKey(intercept: true);

                if (key.KeyChar == '1')
                {
                    List<LinhaGenerica> plan1 = LerPlanilha("DIRF.xlsx");
                    Console.WriteLine("DIRF.xlsx lida.");

                    List<LinhaGenerica> plan2 = LerPlanilha("COPARTICIPAÇAO.xlsx");
                    Console.WriteLine("COPARTICIPAÇAO.xlsx lida.");

                    List<LinhaGenerica> plan3 = LerPlanilha("SAUDE.xlsx");
                    Console.WriteLine("SAUDE.xlsx lida.");

                    List<Empresa> emp = LerPlanilhaCnpj("CNPJS.xlsx");
                    Console.WriteLine("CNPJS.xlsx lida.");

                    RemoverValoresZerados(plan1);

                    plan1 = PreencherDadosAdicionais(plan1);

                    List<PlanilhaFinal> result = MontarResult(plan1, plan2, emp, plan3);
                    Console.WriteLine("Definindo Cabeçalho da planilha final.");
                    string[] colunasDesejadas = new[] { "Matricula", "Titular", "CPF Titular", "Data Nascimento Titular", "Nome do Beneficiario", "CPF Beneficiario", "Data de Nascimento Beneficiario", "Valor", "Valor total", "Subfatura", "Data de referencia", "Nome do Prestador", "Cnpj prestador", "Valor reemboso anos anteriores" };

                    Console.WriteLine("Gerando planilha final.");
                    GerarPlanilhaFinal(result, colunasDesejadas, $"Planilha Coparticipacao {DateTime.Now.Month}.{DateTime.Now.Year}.xlsx");
                    Console.WriteLine("Planilha final gerada com sucesso.");
                }

                if (key.KeyChar == '2') 
                {

                    List<LinhaGenerica> plan1 = LerPlanilha("DIRF.xlsx");
                    Console.WriteLine("DIRF.xlsx lida.");

                    List<LinhaGenerica> plan4 = LerPlanilha("DENTAL.xlsx");
                    Console.WriteLine("DENTAL.xlsx lida.");

                    List<LinhaGenerica> plan3 = LerPlanilha("SAUDE.xlsx");
                    Console.WriteLine("SAUDE.xlsx lida.");

                    List<Empresa> emp = LerPlanilhaCnpj("CNPJS.xlsx");
                    Console.WriteLine("CNPJS.xlsx lida.");

                    List<PlanilhaFinalSaude> resultSaude = MontarResultSaude(plan3, plan4, emp, plan1);

                    string[] colunasDesejadasSaude = new[] { "DATA DE LANCAMENTO", "EMPRESA", "SUB FATURA", "CPF DO BENEFICIARIO", "MATRICULA ESPECIAL", "NUMERO DO CERTIFICADO", "NOME SEGURADO/DEPENDENTE", "DATA DE NASCIMENTO", "IDADE", "CODIGO DO SEXO", "ESTADO CIVIL", "COD. GRAU PARENT.DEP.", "TITULAR OU DEPENDENTE", "CODIGO DO PLANO", "VALOR DO LANCAMENTO", "DATA INICIO VIGENCIA", "DATA DE CANCELAMENTO", "TIPO DE LANÇAMENTO", "DATA TRANSFERENCIA DE SUBFATURA", "CARGO / OCUPACAO" };
                    Console.WriteLine("Definindo Cabeçalho da planilha final.");

                    Console.WriteLine("Gerando planilha final Saude.");
                    GerarPlanilhaFinalSaude(resultSaude, colunasDesejadasSaude, $"SAUDE - FATURA TECNICA {DateTime.Now.Month}.{DateTime.Now.Year}.xlsx");
                    Console.WriteLine("Planilha final Saude gerada com sucesso.");
                }

                if (key.KeyChar == '3')
                {
                    List<LinhaGenerica> plan1 = LerPlanilha("DIRF.xlsx");
                    Console.WriteLine("DIRF.xlsx lida.");

                    List<LinhaGenerica> plan4 = LerPlanilha("DENTAL.xlsx");
                    Console.WriteLine("DENTAL.xlsx lida.");

                    List<LinhaGenerica> plan3 = LerPlanilha("SAUDE.xlsx");
                    Console.WriteLine("SAUDE.xlsx lida.");

                    List<Empresa> emp = LerPlanilhaCnpj("CNPJS.xlsx");
                    Console.WriteLine("CNPJS.xlsx lida.");

                    List<PlanilhaFinalSaude> resultDental = MontarResultDental(plan3, plan4, emp, plan1);

                    string[] colunasDesejadasSaude = new[] { "DATA DE LANCAMENTO", "EMPRESA", "SUB FATURA", "CPF DO BENEFICIARIO", "MATRICULA ESPECIAL", "NUMERO DO CERTIFICADO", "NOME SEGURADO/DEPENDENTE", "DATA DE NASCIMENTO", "IDADE", "CODIGO DO SEXO", "ESTADO CIVIL", "COD. GRAU PARENT.DEP.", "TITULAR OU DEPENDENTE", "CODIGO DO PLANO", "VALOR DO LANCAMENTO", "DATA INICIO VIGENCIA", "DATA DE CANCELAMENTO", "TIPO DE LANÇAMENTO", "DATA TRANSFERENCIA DE SUBFATURA", "CARGO / OCUPACAO" };
                    Console.WriteLine("Definindo Cabeçalho da planilha final.");

                    Console.WriteLine("Gerando planilha final Dental.");
                    GerarPlanilhaFinalSaude(resultDental, colunasDesejadasSaude, $"DENTAL - FATURA TECNICA {DateTime.Now.Month}.{DateTime.Now.Year}.xlsx");
                    Console.WriteLine("Planilha final Dental gerada com sucesso.");
                }

                else if (key.KeyChar != '2' && key.KeyChar != '1')
                {
                    Console.WriteLine("Tecla digitada não corresponde as opções 1 e 2.");
                    StartProcess();
                }

                Console.WriteLine($"{Environment.NewLine}{Environment.NewLine}");
                Console.WriteLine("Deseja gera mais alguma planilha?");
                Console.WriteLine($"Digite:{Environment.NewLine}1 - Sim{Environment.NewLine}2 - Não");
                var key2 = Console.ReadKey(intercept: true);

                if (key2.KeyChar == '1')
                {
                    Console.WriteLine($"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}");
                    StartProcess();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro no processo, mensagem: {ex.Message}{Environment.NewLine}StackTrace:{ex.StackTrace}{Environment.NewLine}");
                Console.ReadLine();
            }

        }

        private void GerarPlanilhaFinalSaude(List<PlanilhaFinalSaude> resultSaude, string[] colunasDesejadas, string caminhoSaida)
        {
            using XLWorkbook wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Consolidado");

            // Cabeçalho
            for (int i = 0; i < colunasDesejadas.Length; i++)
                ws.Cell(1, i + 1).Value = colunasDesejadas[i];

            int row = 2;

            foreach (var linha in resultSaude)
            {
                for (int i = 0; i < colunasDesejadas.Length; i++)
                {
                    switch (i)
                    {
                        case 0:
                            ws.Cell(row, i + 1).Value = linha.DataLancamento;
                            break;
                        case 1:
                            ws.Cell(row, i + 1).Value = linha.Empresa;
                            break;
                        case 2:
                            ws.Cell(row, i + 1).Value = linha.Subfatura;
                            break;
                        case 3:
                            ws.Cell(row, i + 1).Value = linha.CpfBeneficiario;
                            break;
                        case 4:
                            ws.Cell(row, i + 1).Value = linha.Matricula;
                            break;
                        case 5:
                            ws.Cell(row, i + 1).Value = linha.NumeroCertificado;
                            break;
                        case 6:
                                ws.Cell(row, i + 1).Value = linha.NomeDependente;
                            break;
                        case 7:
                            ws.Cell(row, i + 1).Value = linha.DataNascimento;
                            break;
                        case 8:
                            ws.Cell(row, i + 1).Value = linha.Idade;
                            break;
                        case 9:
                            ws.Cell(row, i + 1).Value = linha.Sexo;
                            break;
                        case 10:
                            ws.Cell(row, i + 1).Value = linha.EstadoCivil;
                            break;
                        case 11:
                            ws.Cell(row, i + 1).Value = linha.GrauParentesco;
                            break;
                        case 12:
                            ws.Cell(row, i + 1).Value = linha.TitularDependente;
                            break;
                        case 13:
                            ws.Cell(row, i + 1).Value = linha.CodigoPlano;
                            break;
                        case 14:
                            ws.Cell(row, i + 1).Value = linha.ValorLancamento;
                            break;
                        case 15:
                            ws.Cell(row, i + 1).Value = linha.DataVigencia;
                            break;
                        case 16:
                            ws.Cell(row, i + 1).Value = linha.DataCancelamento;
                            break;
                        case 17:
                            ws.Cell(row, i + 1).Value = linha.TipoLancamento;
                            break;
                        case 18:
                            ws.Cell(row, i + 1).Value = linha.DataTransferenciaSubfatura;
                            break;
                        case 19:
                            ws.Cell(row, i + 1).Value = linha.Cargo;
                            break;
                    }

                }
                row++;
            }

            ws = AplicarEstilos(ws);

            ws.Columns().AdjustToContents();
            wb.SaveAs(caminhoSaida);
        }

        private List<PlanilhaFinalSaude> MontarResultDental(List<LinhaGenerica> Saude, List<LinhaGenerica> Dental, List<Empresa> Empresas, List<LinhaGenerica> Dirf)
        {
            List<PlanilhaFinalSaude> resultSaude = new List<PlanilhaFinalSaude>();

            //Saude.RemoveAll(linha => !Dirf.Any(linha2 =>
            //linha.Colunas.TryGetValue("NOME SEGURADO/DEPENDENTE", out var nomeSeguradoDependente) &&
            //linha2.Colunas.TryGetValue("Nome Dependente", out var nomeDependente) &&
            //string.Equals(nomeSeguradoDependente, nomeDependente, StringComparison.OrdinalIgnoreCase) ||
            //linha.Colunas.TryGetValue("NOME SEGURADO/DEPENDENTE", out var nomeTitular) &&
            //linha2.Colunas.TryGetValue("Nome Segurado Titular", out var Titular) &&
            //string.Equals(nomeTitular, Titular, StringComparison.OrdinalIgnoreCase)));

            Dental.RemoveAll(_ => _.Colunas["TIPO DO REGISTRO"] != "3" || _.Colunas["DATA DE NASCIMENTO"] == "00/00/0000");

            foreach (var d in Dental)
            {
                var dirf = Dirf.Where(_ => _.Nome == d.Nome && _.Colunas.Where(_ => _.Key == "Nome Segurado Titular").FirstOrDefault().Value == d.Nome).FirstOrDefault();
                PlanilhaFinalSaude x = new PlanilhaFinalSaude();
                x.DataLancamento = d.Colunas["DATA DE LANCAMENTO"];
                x.Subfatura = Empresas.Where(_ => _.Subfatura == Convert.ToInt32(d.Colunas.Where(_ => _.Key == "NUMERO DA SUBFATURA").FirstOrDefault().Value)).FirstOrDefault().Subfatura;
                x.Empresa = Empresas?.Where(_ => _.Subfatura == x.Subfatura).FirstOrDefault().Emp;
                x.CpfBeneficiario = dirf?.Colunas["CPF Titular"];
                x.Matricula = d.Colunas["MATRICULA ESPECIAL"];
                x.NumeroCertificado = d.Colunas["NUMERO DO CERTIFICADO"];
                x.NomeDependente = d.Nome;
                x.DataNascimento = d.Colunas["DATA DE NASCIMENTO"];
                x.Idade = CalcularIdade(Convert.ToDateTime(x.DataNascimento)).ToString();
                x.Sexo = ConverterSexo(Convert.ToInt32(d.Colunas["CODIGO DO SEXO"]));
                x.EstadoCivil = ConverterEstadoCivil(Convert.ToInt32(d.Colunas["ESTADO CIVIL"]));
                x.GrauParentesco = ConverterGrauParentesco(Convert.ToInt32(d.Colunas["COD. GRAU PARENT.DEP."]));
                x.TitularDependente = ConverterTirular(Convert.ToInt32(d.Colunas["COD. GRAU PARENT.DEP."]));
                x.CodigoPlano = d.Colunas["CODIGO DO PLANO"];
                x.DataVigencia = d.Colunas["DATA INICIO VIGENCIA"];
                x.DataCancelamento = string.Empty;
                x.TipoLancamento = d.Colunas["TIPO DE LANÇAMENTO"];
                x.DataTransferenciaSubfatura = string.Empty;
                x.Cargo = d.Colunas["CARGO / OCUPACAO"];
                x.ValorLancamento = d.Colunas["VALOR DO LANCAMENTO"];
                resultSaude.Add(x);
            }

            return resultSaude;
        }

        private List<PlanilhaFinalSaude> MontarResultSaude(List<LinhaGenerica> Saude, List<LinhaGenerica> Dental, List<Empresa> Empresas, List<LinhaGenerica> Dirf)
        {
            List<PlanilhaFinalSaude> resultSaude = new List<PlanilhaFinalSaude>();

            //Saude.RemoveAll(linha => !Dirf.Any(linha2 =>
            //linha.Colunas.TryGetValue("NOME SEGURADO/DEPENDENTE", out var nomeSeguradoDependente) &&
            //linha2.Colunas.TryGetValue("Nome Dependente", out var nomeDependente) &&
            //string.Equals(nomeSeguradoDependente, nomeDependente, StringComparison.OrdinalIgnoreCase) ||
            //linha.Colunas.TryGetValue("NOME SEGURADO/DEPENDENTE", out var nomeTitular) &&
            //linha2.Colunas.TryGetValue("Nome Segurado Titular", out var Titular) &&
            //string.Equals(nomeTitular, Titular, StringComparison.OrdinalIgnoreCase)));

            Saude.RemoveAll(_ => _.Colunas["TIPO DO REGISTRO"] != "3" || _.Colunas["DATA DE NASCIMENTO"] == "00/00/0000");

            foreach (var s in Saude)
            {


                var d = Dirf.Where(_ => _.Nome == s.Nome && _.Colunas.Where(_ => _.Key == "Nome Segurado Titular").FirstOrDefault().Value == s.Nome).FirstOrDefault();
                PlanilhaFinalSaude x = new PlanilhaFinalSaude();
                x.DataLancamento = s.Colunas["DATA DE LANCAMENTO"];
                x.Subfatura = Empresas.Where(_ => _.Subfatura == Convert.ToInt32(s.Colunas.Where(_ => _.Key == "NUMERO DA SUBFATURA").FirstOrDefault().Value)).FirstOrDefault().Subfatura;
                x.Empresa = Empresas?.Where(_ => _.Subfatura == x.Subfatura).FirstOrDefault().Emp;
                x.CpfBeneficiario = d?.Colunas["CPF Titular"];
                x.Matricula = s.Colunas["MATRICULA ESPECIAL"];
                x.NumeroCertificado = s.Colunas["NUMERO DO CERTIFICADO"];
                x.NomeDependente = s.Nome;
                x.DataNascimento = s.Colunas["DATA DE NASCIMENTO"];
                x.Idade = CalcularIdade(Convert.ToDateTime(x.DataNascimento)).ToString();
                x.Sexo = ConverterSexo(Convert.ToInt32(s.Colunas["CODIGO DO SEXO"]));
                x.EstadoCivil = ConverterEstadoCivil(Convert.ToInt32(s.Colunas["ESTADO CIVIL"]));
                x.GrauParentesco = ConverterGrauParentesco(Convert.ToInt32(s.Colunas["COD. GRAU PARENT.DEP."]));
                x.TitularDependente = ConverterTirular(Convert.ToInt32(s.Colunas["COD. GRAU PARENT.DEP."]));
                x.CodigoPlano = s.Colunas["CODIGO DO PLANO"];
                x.DataVigencia = s.Colunas["DATA INICIO VIGENCIA"];
                x.DataCancelamento = string.Empty;
                x.TipoLancamento = s.Colunas["TIPO DE LANÇAMENTO"];
                x.DataTransferenciaSubfatura = string.Empty;
                x.Cargo = s.Colunas["CARGO / OCUPACAO"];
                x.ValorLancamento = s.Colunas["VALOR DO LANCAMENTO"];
                resultSaude.Add(x);
            }

            return resultSaude;
        }

        public int CalcularIdade(DateTime dataNascimento)
        {
            var idade = DateTime.Now.Year - dataNascimento.Year;

            // Ajusta a idade se a data de hoje ainda não atingiu o aniversário neste ano
            if (DateTime.Now < dataNascimento.AddYears(idade))
            {
                idade--;
            }

            return idade;
        }

        private string ConverterTirular(int Codigo)
        {
            return Codigo == 0 ? "TITULAR" : "DEPENDENTE";
        }

        private string ConverterGrauParentesco(int Codigo)
        {
            switch (Codigo)
            {
                case 0:
                    return "TITULAR";
                case 1:
                    return "CONJUGE";
                case 2:
                    return "FILHO(a)";
                case 3:
                    return "OUTROS";
                default:
                    return "TITULAR";
            }
        }

        private string ConverterEstadoCivil(int Codigo)
        {
            switch (Codigo)
            {
                case 1:
                    return "SOLTEIRO";
                case 2:
                    return "CASADO";
                case 3:
                    return "VIUVO";
                case 4:
                    return "SEPARADO/DIVORCIADO";
                default:
                    return "SOLTEIRO";
            }
        }

        private string ConverterSexo(int Codigo)
        {
            return Codigo == 1 ? "MASCULINO" : "FEMININO";
        }

        private void RemoverValoresZerados(List<LinhaGenerica> plan1)
        {

            plan1.RemoveAll(linha => linha.Colunas.TryGetValue("Valor Participação", out var valor) && valor == "0");
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

        private List<PlanilhaFinal> MontarResult(List<LinhaGenerica> Dirf, List<LinhaGenerica> Cop, List<Empresa> emp, List<LinhaGenerica> Saude)
        {
            List<PlanilhaFinal> result = new List<PlanilhaFinal>();

            foreach (var d in Dirf)
            {
                var cop = Cop.Where(_ => _.Nome == d.Nome)?.FirstOrDefault();
                PlanilhaFinal x = new PlanilhaFinal();
                x.Matricula = cop?.Colunas.Where(_ => _.Key == "MATRICULA ESPECIAL")?.FirstOrDefault().Value ?? string.Empty; // Onde pego?
                x.Titular = d.Nome;
                x.CPF_Titular = d.Colunas.Where(_ => _.Key == "CPF Titular").FirstOrDefault().Value;
                x.Nome_Do_Beneficiario = d.Colunas.Where(_ => _.Key == "Nome Dependente").FirstOrDefault().Value;

                if (x.Nome_Do_Beneficiario == string.Empty)
                    x.Nome_Do_Beneficiario = x.Titular;

                x.Data_Nascimento_Titular = Saude?.Where(_ => _.Colunas["NOME SEGURADO/DEPENDENTE"] == x.Titular || _.Colunas["NOME SEGURADO/DEPENDENTE"] == x.Nome_Do_Beneficiario)?.FirstOrDefault()?.Colunas["DATA DE NASCIMENTO"];

                x.CPF_Beneficiario = d.Colunas.Where(_ => _.Key == "CPF Dependente").FirstOrDefault().Value;

                if (x.CPF_Beneficiario == string.Empty)
                    x.CPF_Beneficiario = x.CPF_Titular;

                x.Data_De_Nascimento_Beneficiario = d.Colunas.Where(_ => _.Key == "Data Nascimento Dependente").FirstOrDefault().Value;

                if (string.IsNullOrEmpty(x.Data_De_Nascimento_Beneficiario))
                    x.Data_De_Nascimento_Beneficiario = x.Data_Nascimento_Titular;

                x.Valor = d.Colunas.Where(_ => _.Key == "Valor Participação").FirstOrDefault().Value;

                if (x.Valor == "0")
                {
                    x.Valor = string.Empty;
                }

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

            bool isDirf = caminho == "DIRF.xlsx";
            bool isCopart = caminho == "COPARTICIPAÇAO.xlsx";
            bool isSaude = caminho == "SAUDE.xlsx";
            bool isDental = caminho == "DENTAL.xlsx";

            int skip = 0;
            int columnNome = 0;

            if (isDirf)
            {
                skip = 3;
                columnNome = 3;
            }
            if (isCopart)
            {
                skip = 3;
                columnNome = 5;
            }
            if (isSaude)
            {
                skip = 5;
                columnNome = 5;

            }
            if (isDental)
            {
                skip = 5;
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
                                ws.Cell(row, i + 1).Value = GerarData(linha.Data_De_Nascimento_Beneficiario); 
                            
                            else
                                ws.Cell(row, i + 1).Value = TratarData(linha.Data_De_Nascimento_Beneficiario);
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

        private XLCellValue TratarData(string data_De_Nascimento_Beneficiario)
        {
            if (!string.IsNullOrEmpty(data_De_Nascimento_Beneficiario) && data_De_Nascimento_Beneficiario.Contains('/'))
                return data_De_Nascimento_Beneficiario.Substring(0, 10);
            else
                return data_De_Nascimento_Beneficiario;

        }

        private XLCellValue GerarData(string Data)
        {
            DateTime baseDate = new DateTime(1900, 1, 1);
            DateTime dt = baseDate.AddDays(Convert.ToDouble(Data) - 2);
            return $"{dt.Day.ToString().PadLeft(2, '0')}/{dt.Month.ToString().PadLeft(2, '0')}/{dt.Year}";
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
