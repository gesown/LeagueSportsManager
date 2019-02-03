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
using LeagueSportsManager.Areas.Role.Models;

namespace LeagueSportsManager.Areas.Role.Controllers
{
    public class RoleModelsController : ApiController
    {
        private LeagueSportsManager db = new LeagueSportsManager();

        // GET: api/RoleModels
        public IQueryable<RoleModel> GetRoleModels()
        {
            return db.RoleModels;
        }

        // GET: api/RoleModels/5
        [ResponseType(typeof(RoleModel))]
        public IHttpActionResult GetRoleModel(int id)
        {
            RoleModel roleModel = db.RoleModels.Find(id);
            if (roleModel == null)
            {
                return NotFound();
            }

            return Ok(roleModel);
        }

        // PUT: api/RoleModels/5
        [ResponseType(typeof(void))]
        public IHttpActionResult PutRoleModel(int id, RoleModel roleModel)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (id != roleModel.RoleId)
            {
                return BadRequest();
            }

            db.Entry(roleModel).State = EntityState.Modified;

            try
            {
                db.SaveChanges();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!RoleModelExists(id))
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

        // POST: api/RoleModels
        [ResponseType(typeof(RoleModel))]
        public IHttpActionResult PostRoleModel(RoleModel roleModel)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            db.RoleModels.Add(roleModel);
            db.SaveChanges();

            return CreatedAtRoute("DefaultApi", new { id = roleModel.RoleId }, roleModel);
        }

        // DELETE: api/RoleModels/5
        [ResponseType(typeof(RoleModel))]
        public IHttpActionResult DeleteRoleModel(int id)
        {
            RoleModel roleModel = db.RoleModels.Find(id);
            if (roleModel == null)
            {
                return NotFound();
            }

            db.RoleModels.Remove(roleModel);
            db.SaveChanges();

            return Ok(roleModel);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        private bool RoleModelExists(int id)
        {
            return db.RoleModels.Count(e => e.RoleId == id) > 0;
        }
    }
}