using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web.Http.Description;
using LeagueSportsManager;
using LeagueSportsManager.Areas.Admin.Models;

namespace LeagueSportsManager.Areas.Admin.Controllers
{
    public class AdminsController : ApiController
    {
        private LeagueSportsManager db = new LeagueSportsManager();

        // GET: api/Admins
        public IQueryable<Models.AdminModel> GetAdmins()
        {
            return db.AdminModels;
        }

        // GET: api/Admins/5
        [ResponseType(typeof(Models.AdminModel))]
        public IHttpActionResult GetAdmin(int id)
        {
            Models.AdminModel admin = db.AdminModels.Find(id);
            if (admin == null)
            {
                return NotFound();
            }

            return Ok(admin);
        }

        // PUT: api/Admins/5
        [ResponseType(typeof(void))]
        public IHttpActionResult PutAdmin(int id, Models.AdminModel admin)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (id != admin.AdminId)
            {
                return BadRequest();
            }

            db.Entry(admin).State = EntityState.Modified;

            try
            {
                db.SaveChanges();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!AdminExists(id))
                {
                    return NotFound();
                }
                else
                {
                    throw;
                }
            }

            return StatusCode(HttpStatusCode.NoContent);
        }

        // POST: api/Admins
        [ResponseType(typeof(Models.AdminModel))]
        public IHttpActionResult PostAdmin(Models.AdminModel admin)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            db.AdminModels.Add(admin);
            db.SaveChanges();

            return CreatedAtRoute("DefaultApi", new { id = admin.AdminId }, admin);
        }

        // DELETE: api/Admins/5
        [ResponseType(typeof(Models.AdminModel))]
        public IHttpActionResult DeleteAdmin(int id)
        {
            Models.AdminModel admin = db.AdminModels.Find(id);
            if (admin == null)
            {
                return NotFound();
            }

            db.AdminModels.Remove(admin);
            db.SaveChanges();

            return Ok(admin);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        private bool AdminExists(int id)
        {
            return db.AdminModels.Count(e => e.AdminId == id) > 0;
        }
    }
}