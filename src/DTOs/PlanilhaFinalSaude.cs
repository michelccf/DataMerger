using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataMerger.DTOs
{
    public class PlanilhaFinalSaude
    {
        public string DataLancamento { get; set; }
        public string Empresa { get; set; }
        public int Subfatura { get; set; }
        public string CpfBeneficiario { get; set; }
        public string Matricula { get; set; }
        public string NumeroCertificado { get; set; }
        public string NomeDependente { get; set; }
        public string DataNascimento { get; set; }
        public string Idade { get; set; }
        public string Sexo { get; set; }
        public string EstadoCivil { get; set; }
        public string GrauParentesco { get; set; }
        public string TitularDependente { get; set; }
        public string CodigoPlano { get; set; }
        public string ValorLancamento { get; set; }
        public string DataVigencia { get; set; }
        public string DataCancelamento { get; set; }
        public string TipoLancamento { get; set; }
        public string DataTransferenciaSubfatura { get; set; }
        public string Cargo { get; set; }
    }
}
