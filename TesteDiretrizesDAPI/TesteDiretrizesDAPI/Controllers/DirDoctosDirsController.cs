using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using TesteDiretrizesDAPI.Models;
using TesteDiretrizesDAPI.Repository;
using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft.AspNetCore.Cors;

namespace TesteDiretrizesDAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class DirDoctosDirsController : ControllerBase
    {
        private readonly DirDoctosContext _context;

        //public DirDoctosDirsController(DirDoctosContext context)
        //{
        //   _context = context;
        //}

        private readonly DirDoctosDirRepository _dirdoctosdirRepository;
        //
        public DirDoctosDirsController()
        {
            _dirdoctosdirRepository = new DirDoctosDirRepository();
        }

        
        // GET: api/DirDoctosDirs
        [HttpGet]
        public async Task<ActionResult<IEnumerable<DirDoctosDir>>> GetDirDoctos()
        {
            
            var dir = new DiretrizesD.DirDoctosDir();

            var db = dir.dbConecta(0,1, "SELECT * FROM DirDoctosDir WHERE tpDiretriz = 1 ORDER BY dtRegistro");
            if(db == 0)
            {
                List<Models.DirDoctosDir> ListDirDoctosDir = new List<Models.DirDoctosDir>();
             
                var RC = dir.leSeq(1);
                while (RC == 0)
                {
                   Models.DirDoctosDir diretrizes = new Models.DirDoctosDir();

                    diretrizes.idDiretriz = dir.idDiretriz;
                    diretrizes.nomeDiretriz = dir.nomeDiretriz;
                    diretrizes.tipodocto1 = dir.tipodocto1;
                    diretrizes.tipodocto2 = dir.tipodocto2;
                    diretrizes.dtRegistro = dir.dtRegistro;
                    diretrizes.dtFim = dir.dtFim;
                    diretrizes.tpDoctoDiretriz = dir.tpDoctoDiretriz;
                    diretrizes.nomeDocto1 = dir.nomeDocto1;
                    diretrizes.nomeDocto2 = dir.nomeDocto2;
                    diretrizes.idEmpresa = dir.idEmpresa;
                    diretrizes.idProjeto = dir.idProjeto;
                    diretrizes.tpDiretriz = dir.tpDiretriz;

                    ListDirDoctosDir.Add(diretrizes);
                   

                    RC = dir.leSeq(0);
                    
                }

                return Ok(new { status = true, response = RC, ListDirDoctosDir });
                }
            else
            {
                return StatusCode(StatusCodes.Status500InternalServerError, "Erro com o banco de dados");
            }

            
        }
           

            // GET: api/DirDoctosDirs/5
            [HttpGet("{headerId}")]
        public async Task<IActionResult> GetDirDoctosDir(int headerId )
        {
            var dirDoctosDir = new DiretrizesD.DirDoctosDir();

            var RCdb = dirDoctosDir.dbConecta(0, 0, "");
            if(RCdb == 0)
            {
                var localiza = dirDoctosDir.localiza(headerId, 1);

                if(localiza == 0)
                {
                    Models.DirDoctosDir Listdiretrizes = new Models.DirDoctosDir();

                    Listdiretrizes.idDiretriz = headerId; //dirDoctosDir.idDiretriz;
                    Listdiretrizes.nomeDiretriz = dirDoctosDir.nomeDiretriz;
                    Listdiretrizes.tipodocto1 = dirDoctosDir.tipodocto1;
                    Listdiretrizes.tipodocto2 = dirDoctosDir.tipodocto2;
                    Listdiretrizes.dtRegistro = dirDoctosDir.dtRegistro;
                    Listdiretrizes.dtFim = dirDoctosDir.dtFim;
                    Listdiretrizes.tpDoctoDiretriz = dirDoctosDir.tpDoctoDiretriz;
                    Listdiretrizes.nomeDocto1 = dirDoctosDir.nomeDocto1;
                    Listdiretrizes.nomeDocto2 = dirDoctosDir.nomeDocto2;
                    Listdiretrizes.idEmpresa = dirDoctosDir.idEmpresa;
                    Listdiretrizes.idProjeto = dirDoctosDir.idProjeto;
                    Listdiretrizes.tpDiretriz = dirDoctosDir.tpDiretriz;


                    return Ok(new { status = true, response = "Sucesso ao carregar alteração", Listdiretrizes });

                }
                else
                {
                    return StatusCode(StatusCodes.Status500InternalServerError, "erro interno");
                }
                

            }else
                {
                    return StatusCode(StatusCodes.Status500InternalServerError, "erro interno");
                }
        }


        // PUT: api/DirDoctosDirs/5
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPut("{id}")]
        public async Task<IActionResult> PutDirDoctosDir(int id, DiretrizesD.DirDoctosDir dirDoctosDir)
        {
            var RCdb = dirDoctosDir.dbConecta(0, 0, "");
            if (RCdb == 0)
            {
                var localiza = dirDoctosDir.localiza(id, 1);

                if (localiza == 0)
                {
                    Models.DirDoctosDir Listdiretrizes = new Models.DirDoctosDir();

                    Listdiretrizes.idDiretriz = dirDoctosDir.idDiretriz;
                    Listdiretrizes.nomeDiretriz = dirDoctosDir.nomeDiretriz;
                    Listdiretrizes.tipodocto1 = dirDoctosDir.tipodocto1;
                    Listdiretrizes.tipodocto2 = dirDoctosDir.tipodocto2;
                    Listdiretrizes.dtRegistro = dirDoctosDir.dtRegistro;
                    Listdiretrizes.dtFim = dirDoctosDir.dtFim;
                    Listdiretrizes.tpDoctoDiretriz = dirDoctosDir.tpDoctoDiretriz;
                    Listdiretrizes.nomeDocto1 = dirDoctosDir.nomeDocto1;
                    Listdiretrizes.nomeDocto2 = dirDoctosDir.nomeDocto2;
                    Listdiretrizes.idEmpresa = dirDoctosDir.idEmpresa;
                    Listdiretrizes.idProjeto = dirDoctosDir.idProjeto;
                    Listdiretrizes.tpDiretriz = dirDoctosDir.tpDiretriz;


                    return Ok(new { status = true, response = "Sucesso ao carregar alteração", Listdiretrizes });

                }
                else
                {
                    return StatusCode(StatusCodes.Status500InternalServerError, "erro interno");
                }


            }
            else
            {
                return StatusCode(StatusCodes.Status500InternalServerError, "erro interno");
            }



        }

        //POST: api/DirDoctosDirs
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPost]
        public async Task<ActionResult<DirDoctosDir>> PostDirDoctosDir(DiretrizesD.DirDoctosDir dirDoctosDir)
        {
            
             var RCdb =  dirDoctosDir.dbConecta(0, 0, "");
            if (RCdb == 0) { 
            
                var RCinclui = dirDoctosDir.inclui(false, dirDoctosDir.tipodocto1, false, dirDoctosDir.tipodocto2, "asd");
                if (RCinclui == 0)
                {
                    return StatusCode(StatusCodes.Status201Created, "Dados gravado com sucesso!");
                }
                else
                {
                    return StatusCode(StatusCodes.Status500InternalServerError, "Erro interno ao gravar dados!");
                }
            }
            else
            {
                return StatusCode(StatusCodes.Status500InternalServerError, "Erro com conexão de banco de dados!");
            }
 
        }

        // DELETE: api/DirDoctosDirs/5
        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteDirDoctosDir(int id, DiretrizesD.DirDoctosDir dirDoctosDir)
        {
            var RCdb = dirDoctosDir.dbConecta(0, 0, "");

            if(RCdb == 0)
            {
                var RCElimina = dirDoctosDir.elimina(id);
                return StatusCode(StatusCodes.Status200OK, "Sucesso ao deletar diretriz!");
            }
            else
            {
                return StatusCode(StatusCodes.Status500InternalServerError, "Erro ao deletar Diretriz!");
            }
        }

        private bool DirDoctosDirExists(int id)
        {
            return (_context.DirDoctos?.Any(e => e.idDiretriz == id)).GetValueOrDefault();
        }

        /*//GET api/values/5
        [HttpGet("{id}")]
        public string Get(int id)
        {
            return $"TESTE: #{id}";
        }*/

    }

    

}
