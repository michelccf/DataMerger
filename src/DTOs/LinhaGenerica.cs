using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataMerger.DTOs
{
    public class LinhaGenerica
    {
        public string Nome { get; set; } = string.Empty; // chave para cruzar os dados
        public string Total { get; set; }
        public int QtdDependentes { get; set; }
        public Dictionary<string, string> Colunas { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    }
}
