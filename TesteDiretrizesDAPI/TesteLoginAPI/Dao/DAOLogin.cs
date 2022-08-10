using Microsoft.Data.SqlClient;
using TesteDiretrizesDAPI.Models;
using System.Data;

namespace TesteDiretrizesDAPI.Dao
{
    public class DAOLogin
    {
        string conexao = @"Data Source=DESKTOP-HV9333S;Initial Catalog=Auditeste;Integrated Security=False;User ID=felipe.rozzi;Password=Felipe1999#;Encrypt=False;TrustServerCertificate=False"; //string local

        public List<LoginUsuarios> GetAllUsuarios()
        {
            List<LoginUsuarios> ListLogins = new List<LoginUsuarios>();
            using (SqlConnection conn = new SqlConnection(conexao))
            {

                conn.Open();
                using (SqlCommand cmd = new SqlCommand("select * from FuncionariosUsuarios", conn))
                {

                    cmd.CommandType = CommandType.Text;
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader != null)
                        {
                            while (reader.Read())
                            {
                                var login = new LoginUsuarios();
                                login.id = Convert.ToInt32(reader["ID"]);
                                login.usuario = reader["nome"].ToString();
                                login.password = reader["senha"].ToString();

                                ListLogins.Add(login);
                            }
                        }
                    }
                }
            }
            return ListLogins;
        }
        
    }
}
