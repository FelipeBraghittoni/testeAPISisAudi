using System.ComponentModel.DataAnnotations;

namespace TesteDiretrizesDAPI.Models
{
    public class DirDoctosDir
    {
        [Key]
        public int idDiretriz { get; set; }

        public string nomeDiretriz { get; set; }

        public string tipodocto1 { get; set; }

        public string tipodocto2 { get; set; }
        public DateTime dtRegistro { get; set; }

        public string tpDoctoDiretriz { get; set; }

        public DateTime dtFim { get; set; }

        public string nomeDocto1 { get; set; }

        public string nomeDocto2 { get; set; }

        public int idEmpresa { get; set; }

        public int idDepto { get; set; }

        public int idProjeto { get; set; }

        public int tpDiretriz { get; set; }
    }
}
