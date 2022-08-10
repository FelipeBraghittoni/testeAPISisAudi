using TesteDiretrizesDAPI.Dao;
using TesteDiretrizesDAPI.Models;

namespace TesteDiretrizesDAPI.Repository
{
    public class DirDoctosDirRepository
    {
        private readonly DAODiretrizes _daoDiretrizes;
        
        
        public DirDoctosDirRepository()
        {
            _daoDiretrizes = new DAODiretrizes();
        }

        public List<DirDoctosDir> GetDirDoctos
        {
            get
            {
                return _daoDiretrizes.GetDiretrizes();
            }
        }
    }
}
