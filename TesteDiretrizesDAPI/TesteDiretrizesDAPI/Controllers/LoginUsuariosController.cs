using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using TesteDiretrizesDAPI.Repository;
using TesteDiretrizesDAPI.Models;
using SegurancaD;

namespace TesteDiretrizesDAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class LoginUsuariosController : ControllerBase
    {
        private readonly LoginUsuariosContext _context;
        private readonly LoginUsuariosRepository _loginUsuariosRepository;

        public LoginUsuariosController()
        {
            _loginUsuariosRepository = new LoginUsuariosRepository();
        }

        // GET: api/LoginUsuarios
        [HttpGet]
        public async Task<ActionResult<IEnumerable<LoginUsuarios>>> GetLogin()
        {
            var login = new LoginUsuarios();
            var result = _loginUsuariosRepository.GetLogin();
            return Ok(new { status = true, response = result });
        }

        // GET: api/LoginUsuarios/5
        [HttpGet("{id}")]
        public async Task<ActionResult<LoginUsuarios>> GetLoginUsuarios(decimal id)
        {
            if (_context.Login == null)
            {
                return NotFound();
            }
            var loginUsuarios = await _context.Login.FindAsync(id);

            if (loginUsuarios == null)
            {
                return NotFound();
            }

            return loginUsuarios;
        }

        // PUT: api/LoginUsuarios/5
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPut("{id}")]
        public async Task<IActionResult> PutLoginUsuarios(decimal id, LoginUsuarios loginUsuarios)
        {
            if (id != loginUsuarios.id)
            {
                return BadRequest();
            }

            _context.Entry(loginUsuarios).State = EntityState.Modified;

            try
            {
                await _context.SaveChangesAsync();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!LoginUsuariosExists(id))
                {
                    return NotFound();
                }
                else
                {
                    throw;
                }
            }

            return NoContent();
        }

        // POST: api/LoginUsuarios
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPost]
        public async Task<ActionResult<LoginUsuarios>> PostLoginUsuarios(SegurancaD.segUsuario segUsuario)
        {
           
                
                

                //var valor2 = new SegurancaD.cdSeguranca1();
                //var number2 = valor2.GravaDADOSsys("Provider=SQLOLEDB;Data Source=192.168.1.119,1433;sqlexpress;Initial Catalog=Auditeste;User Id=sa;Password=sisaudi@2022;");
                //"Provider=SQLOLEDB;Data Source=DESKTOP-HV9333S,1433;sqlexpress;Initial Catalog=Auditeste;User Id=felipe.rozzi;Password=Felipe1999#;");
                //& vbLf & "Provider=SQLOLEDB;Data Source=152.249.241.111,8282\sqlexpress;Initial Catalog=appHoras;User Id=sa;Password=@udiHor@sDevto;");

                //var db = new Utilidades.DBUtils();
                //var dbConecta = db.dbConecta(0, 0, "");

                
                var valor = new SegurancaD.segUsuario();
                var number = valor.trataSeguranca(10, segUsuario.Usuario, segUsuario.senha);

                return StatusCode(StatusCodes.Status202Accepted, "Conexão com banco realizada com sucesso");
                
               


              


        }

        // DELETE: api/LoginUsuarios/5
        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteLoginUsuarios(decimal id)
        {
            if (_context.Login == null)
            {
                return NotFound();
            }
            var loginUsuarios = await _context.Login.FindAsync(id);
            if (loginUsuarios == null)
            {
                return NotFound();
            }

            _context.Login.Remove(loginUsuarios);
            await _context.SaveChangesAsync();

            return NoContent();
        }

        private bool LoginUsuariosExists(decimal id)
        {
            return (_context.Login?.Any(e => e.id == id)).GetValueOrDefault();
        }
    }
}
