using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataMerger.DTOs
{
    public class PlanilhaFinal
    {
        public string Matricula { get; set; }
        public string Titular { get; set; }
        public string CPF_Titular { get; set; }
        public string Data_Nascimento_Titular { get; set; }
        public string Nome_Do_Beneficiario { get; set; }
        public string CPF_Beneficiario { get; set; }
        public string Data_De_Nascimento_Beneficiario { get; set; }
        public string Valor { get; set; }
        public string Valor_Total { get; set; }
        public string Subfatura { get; set; }
        public string Data_De_referencia { get; set; }
        public string NomeEmpresa { get; set; }
        public string Cnpj_Prestador { get; set; }
        public string Nome_Do_Prestador { get; set; }
        public string Valor_reemboso_anos_anteriores { get; set; }
    }
}
