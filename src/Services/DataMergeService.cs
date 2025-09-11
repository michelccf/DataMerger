using ClosedXML.Excel;
using DataMerger.DTOs;
using DataMerger.Interfaces;
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
        public DataMergeService() 
        {
            
        }

        public async Task StartProcess()
        {
            List<LinhaGenerica> plan1 = LerPlanilha("planilha1.xlsx");
            Console.WriteLine("Primeira planinha lida.");

            List<LinhaGenerica> plan2 = LerPlanilha("planilha2.xlsx");
            Console.WriteLine("Segunda planinha lida.");

            List<LinhaGenerica> plan3 = LerPlanilha("planilha3.xlsx");
            Console.WriteLine("Terceira planinha lida.");

            List<LinhaGenerica> dadosMesclados = MesclarPlanilhas(new List<List<LinhaGenerica>> { plan1, plan2, plan3 });
            Console.WriteLine("Dados mesclados com Sucesso.");

            // Definição de ordem das colunas na planilha final.
            string[] colunasDesejadas = new[] { "Nome", "Email", "Cpf", "Endereço", "Empresa" };
            Console.WriteLine("Definindo Cabeçalho da planilha final.");

            Console.WriteLine("Gerando planilha final.");
            GerarPlanilhaFinal(dadosMesclados, colunasDesejadas, "planilha4.xlsx");
            Console.WriteLine("Planilha final gerada com sucesso.");
        }

        private List<LinhaGenerica> LerPlanilha(string caminho)
        {
            var lista = new List<LinhaGenerica>();

            using XLWorkbook wb = new XLWorkbook(caminho);
            IXLWorksheet ws = wb.Worksheet(1); // primeira aba
            List<IXLRangeRow> rows = ws.RangeUsed().RowsUsed().ToList();

            if (!rows.Any()) return lista;

            // pegar o cabeçalho
            List<string> header = rows.First().Cells().Select(c => c.GetString()).ToList();

            // percorrer linhas
            foreach (IXLRangeRow row in rows.Skip(1))
            {
                //Pega o nome do cliente refferênte à linha em questão
                string nome = row.Cell(1).GetString();
                if (string.IsNullOrWhiteSpace(nome)) continue;

                var linha = new LinhaGenerica { Nome = nome };

                for (int i = 0; i < header.Count; i++)
                {
                    //Atribrui o valor na linha referente à coluna)
                    linha.Colunas[header[i]] = row.Cell(i + 1).GetString();
                }

                lista.Add(linha);
            }

            return lista;
        }


        private List<LinhaGenerica> MesclarPlanilhas(List<List<LinhaGenerica>> planilhas)
        {
            var dict = new Dictionary<string, LinhaGenerica>(StringComparer.OrdinalIgnoreCase);

            foreach (var plan in planilhas)
            {
                foreach (var linha in plan)
                {
                    if (!dict.TryGetValue(linha.Nome, out var existente))
                    {
                        existente = new LinhaGenerica { Nome = linha.Nome };
                        dict[linha.Nome] = existente;
                    }

                    // adiciona todas as colunas, sem duplicar
                    foreach (var kv in linha.Colunas)
                    {
                        if (!existente.Colunas.ContainsKey(kv.Key))
                            existente.Colunas[kv.Key] = kv.Value;
                    }
                }
            }

            return dict.Values.ToList();
        }

        private void GerarPlanilhaFinal(List<LinhaGenerica> dados, string[] colunasDesejadas, string caminhoSaida)
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Consolidado");

            // Cabeçalho
            for (int i = 0; i < colunasDesejadas.Length; i++)
                ws.Cell(1, i + 1).Value = colunasDesejadas[i];

            int row = 2;
            foreach (var linha in dados)
            {
                for (int i = 0; i < colunasDesejadas.Length; i++)
                {
                    string coluna = colunasDesejadas[i];
                    ws.Cell(row, i + 1).Value = linha.Colunas.TryGetValue(coluna, out var valor) ? valor : "";
                }
                row++;
            }

            ws.Columns().AdjustToContents();
            wb.SaveAs(caminhoSaida);
        }

    }
}
