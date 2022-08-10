using System.ComponentModel.DataAnnotations;
using Newtonsoft.Json;

namespace TesteDiretrizesDAPI.Models
{
    public class DirDoctosDir
    {
        [Key]
        public decimal idDiretriz { get; set; }

        public string nomeDiretriz { get; set; }

        public short tpDoctoDiretriz { get; set; }

        public DateTime dtRegistro { get; set; }
        public DateTime dtFim { get; set; }

        public string txtDocto { get; set; }

        public string docto1 { get; set; }

        public string nomeDocto1 { get; set; }

        public string tipodocto1 { get; set; }

        public string docto2 { get; set; }

        public string nomeDocto2 { get; set; }

        public string tipodocto2 { get; set; }

        public short idEmpresa { get; set; }

        public short idDepto { get; set; }
        public short idProjeto { get; set; }
        public short tpDiretriz { get; set; }
        //public string upsize_ts { get; set; }

       
    }


}
