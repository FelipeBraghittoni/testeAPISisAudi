using TesteDiretrizesDAPI.Dao;
using TesteDiretrizesDAPI.Models;

namespace TesteDiretrizesDAPI.Repository
{
    public class LoginUsuariosRepository
    {
        private readonly DAOLogin _daoLogin;


        public LoginUsuariosRepository()
        {
            _daoLogin = new DAOLogin();
        }

        public List<LoginUsuarios> GetLogin()
        {
            return _daoLogin.GetLogin();
           
        }
    }
}
