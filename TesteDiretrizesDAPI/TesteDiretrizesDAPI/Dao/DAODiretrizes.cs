using Microsoft.Data.SqlClient;
using TesteDiretrizesDAPI.Models;
using System.Data;

namespace TesteDiretrizesDAPI.Dao
{
    public class DAODiretrizes
    {
        string conexao = @"Data Source=DESKTOP-HV9333S;Initial Catalog=Auditeste;Integrated Security=True";
        //string conexao = "Server=(DESKTOP-HV9333S)\v11.0;Integrated Security = true;";
        //string conexao = @"Data Source=DESKTOP-PUGVQGJ,1433\sqlexpress;Initial Catalog=Auditeste;User Id=sa;Password=sisaudi@2022;Trusted_Connection=True;";
        //DBConectaSQLExpress = "Provider=SQLOLEDB;Data Source=DESKTOP-PUGVQGJ,1433\sqlexpress;Initial Catalog=Auditeste;User Id=sa;Password=sisaudi@2022;"

        public List<DirDoctosDir> GetDiretrizes()
        {

            List<DirDoctosDir> listDiretrizes = new List<DirDoctosDir>();
            using (SqlConnection conn = new SqlConnection(conexao))
            {
            
                    conn.Open();             
                using (SqlCommand cmd = new SqlCommand("Select * From DirDoctosDir", conn))
                {
                    
                    //cmd.CommandType = CommandType.Text;
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            while (reader.Read())
                            {
                                var dir = new DirDoctosDir();
                                dir.nomeDiretriz = reader["nomeDiretriz"].ToString();
                                listDiretrizes.Add(dir);
                            }
                        }
                    }
                }
            }
            return listDiretrizes;
        }
    }
}
