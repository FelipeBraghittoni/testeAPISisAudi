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
            var result = _dirdoctosdirRepository.GetDirDoctos;
            return Ok(new {status = true, response = result});


            //      if (_context.DirDoctos == null)
            //      {
            //          return NotFound();
            //      }
            //        return await _context.DirDoctos.ToListAsync();
            //    }
        }
            List<DirDoctosDir>dirdoctosdir= new List<DirDoctosDir>();
            // GET: api/DirDoctosDirs/5
            [HttpGet("{id}")]
        public async Task<ActionResult<DirDoctosDir>> GetDirDoctosDir(int id)
        {
          if (_context.DirDoctos == null)
          {
              return NotFound();
          }
            var dirDoctosDir = await _context.DirDoctos.FindAsync(id);

            if (dirDoctosDir == null)
            {
                return NotFound();
            }

            return dirDoctosDir;
        }

        // PUT: api/DirDoctosDirs/5
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPut("{id}")]
        public async Task<IActionResult> PutDirDoctosDir(int id, DirDoctosDir dirDoctosDir)
        {
            if (id != dirDoctosDir.idDiretriz)
            {
                return BadRequest();
            }

            _context.Entry(dirDoctosDir).State = EntityState.Modified;

            try
            {
                await _context.SaveChangesAsync();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!DirDoctosDirExists(id))
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

        //POST: api/DirDoctosDirs
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPost]
       /* public async Task<ActionResult<DirDoctosDir>> PostDirDoctosDir(DiretrizesD.DirDoctosDir dirDoctosDir)
        {
            try
            {
                if ('a' == 'b')
                {
                    await _context.SaveChangesAsync();
                }

                var dir = new DiretrizesD.DirDoctosDir();
                var number = dir.inclui (true, "aaa", true,"bbb","ccc");
                Console.WriteLine("TESTEEEEEEEEEE " + number.ToString());
                if (number == 1)
                {
                    return Ok(new { status = true, response = number.ToString() });
                }
                return Ok(new { status = true, response = number.ToString() });
            }
            catch (DbUpdateException)
            {
                if (LoginUsuariosExists(loginUsuarios.id))
                {
                    return Conflict();
                }
                else
                {
                    throw;
                }
            }
        }*/

        // DELETE: api/DirDoctosDirs/5
        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteDirDoctosDir(int id)
        {
            if (_context.DirDoctos == null)
            {
                return NotFound();
            }
            var dirDoctosDir = await _context.DirDoctos.FindAsync(id);
            if (dirDoctosDir == null)
            {
                return NotFound();
            }

            _context.DirDoctos.Remove(dirDoctosDir);
            await _context.SaveChangesAsync();

            return NoContent();
        }

        private bool DirDoctosDirExists(int id)
        {
            return (_context.DirDoctos?.Any(e => e.idDiretriz == id)).GetValueOrDefault();
        }

    }
    
}
