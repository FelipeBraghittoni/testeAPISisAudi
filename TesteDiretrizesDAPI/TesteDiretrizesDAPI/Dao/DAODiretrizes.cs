using Microsoft.Data.SqlClient;
using TesteDiretrizesDAPI.Models;
using System.Data;
//using System.Text.Json;
//using System.Text.Json.Serialization;
using Newtonsoft.Json;

namespace TesteDiretrizesDAPI.Dao
{
    public class DAODiretrizes

    {
        string conexao = @"Data Source=DESKTOP-HV9333S;Initial Catalog=Auditeste;Integrated Security=False;User ID=felipe.rozzi;Password=Felipe1999#;Encrypt=False;TrustServerCertificate=False"; //string local
        //string conexao = "Server=(DESKTOP-HV9333S)\v11.0;Integrated Security = true;";
        //string conexao = @"Data Source=DESKTOP-PUGVQGJ,1433\sqlexpress;Initial Catalog=Auditeste;User Id=sa;Password=sisaudi@2022;Trusted_Connection=True;"; //string dev
        //DBConectaSQLExpress = "Provider=SQLOLEDB;Data Source=DESKTOP-PUGVQGJ,1433\sqlexpress;Initial Catalog=Auditeste;User Id=sa;Password=sisaudi@2022;"

        public List<DirDoctosDir> GetDiretrizes()
        {

            List<DirDoctosDir> listDiretrizes = new List<DirDoctosDir>();
            using (SqlConnection conn = new SqlConnection(conexao))
            {

                conn.Open();
                //using (SqlCommand cmd = new SqlCommand("select * from dirdoctosdir where idDiretriz = '7202201' ", conn))
                #region
                using (SqlCommand cmd = new SqlCommand("select * from dirdoctosdir ", conn))
                {

                    cmd.CommandType = CommandType.Text;
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader != null)
                        {
                            while (reader.Read())
                            {
                                var dir = new DirDoctosDir();
                                dir.idDiretriz = Convert.ToDecimal(reader["idDiretriz"]);
                                dir.nomeDiretriz = reader["nomeDiretriz"].ToString();
                                dir.tpDoctoDiretriz = Convert.ToInt16(reader["tpDoctoDiretriz"]);
                                dir.dtRegistro = Convert.ToDateTime(reader["dtRegistro"]);
                                dir.dtFim = Convert.ToDateTime(reader["dtFim"]);
                                dir.txtDocto = reader["txtDocto"].ToString();
                                dir.docto1 = reader["docto1"].ToString();
                                dir.nomeDocto1 = reader["nomeDocto1"].ToString();
                                dir.tipodocto1 = reader["tipodocto1"].ToString();
                                dir.docto2 = reader["docto2"].ToString();
                                dir.nomeDocto2 = reader["nomeDocto2"].ToString();
                                dir.tipodocto2 = reader["tipodocto1"].ToString();
                                dir.idEmpresa = Convert.ToInt16(reader["idEmpresa"]);
                                dir.idDepto = Convert.ToInt16(reader["idDepto"]);
                                dir.idProjeto = Convert.ToInt16(reader["idProjeto"]);
                                dir.tpDiretriz = Convert.ToInt16(reader["tpDiretriz"]);

                                listDiretrizes.Add(dir);
                            }

                        }
                    }
                }
                #endregion// list dirdoctosdir

            }


            return listDiretrizes;
        }

        /*public List<DirDoctosDir> PostDiretrizes()
        {
            List<DirDoctosDir> listDiretrizes = new List<DirDoctosDir>();
            using (SqlConnection conn = new SqlConnection(conexao))
            {
                conn.Open();
                //using (SqlCommand cmd = new SqlCommand("select * from dirdoctosdir where idDiretriz = '7202201' ", conn))
                #region
                using (SqlCommand cmd = new SqlCommand("insert into TesteDiretrizes ", conn))
                {

                    cmd.CommandType = CommandType.Text;
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader != null)
                        {
                            while (reader.Read())
                            {
                                var dir = new DirDoctosDir();
                                dir.idDiretriz = Convert.ToDecimal(reader["idDiretriz"]);
                                dir.nomeDiretriz = reader["nomeDiretriz"].ToString();
                                dir.tpDoctoDiretriz = Convert.ToInt16(reader["tpDoctoDiretriz"]);
                                dir.dtRegistro = Convert.ToDateTime(reader["dtRegistro"]);
                                dir.dtFim = Convert.ToDateTime(reader["dtFim"]);
                                dir.txtDocto = reader["txtDocto"].ToString();
                                dir.docto1 = reader["docto1"].ToString();
                                dir.nomeDocto1 = reader["nomeDocto1"].ToString();
                                dir.tipodocto1 = reader["tipodocto1"].ToString();
                                dir.docto2 = reader["docto2"].ToString();
                                dir.nomeDocto2 = reader["nomeDocto2"].ToString();
                                dir.tipodocto2 = reader["tipodocto1"].ToString();
                                dir.idEmpresa = Convert.ToInt16(reader["idEmpresa"]);
                                dir.idDepto = Convert.ToInt16(reader["idDepto"]);
                                dir.idProjeto = Convert.ToInt16(reader["idProjeto"]);
                                dir.tpDiretriz = Convert.ToInt16(reader["tpDiretriz"]);

                                listDiretrizes.Add(dir);
                            }

                        }
                    }
                }
                #endregion// list dirdoctosdir
            }
        }*/
    }
}
