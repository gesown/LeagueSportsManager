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
using LeagueSportsManager.Areas.Register.Models;

namespace LeagueSportsManager.Areas.Register.Controllers
{
    public class RegisterModelsController : ApiController
    {
        private LeagueSportsManager db = new LeagueSportsManager();

        // GET: api/RegisterModels
        public IQueryable<RegisterModel> GetRegisterModels()
        {
            return db.RegisterModels;
        }

        // GET: api/RegisterModels/5
        [ResponseType(typeof(RegisterModel))]
        public IHttpActionResult GetRegisterModel(int id)
        {
            RegisterModel registerModel = db.RegisterModels.Find(id);
            if (registerModel == null)
            {
                return NotFound();
            }

            return Ok(registerModel);
        }

        // PUT: api/RegisterModels/5
        [ResponseType(typeof(void))]
        public IHttpActionResult PutRegisterModel(int id, RegisterModel registerModel)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (id != registerModel.RegisterId)
            {
                return BadRequest();
            }

            db.Entry(registerModel).State = EntityState.Modified;

            try
            {
                db.SaveChanges();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!RegisterModelExists(id))
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

        // POST: api/RegisterModels
        [ResponseType(typeof(RegisterModel))]
        public IHttpActionResult PostRegisterModel(RegisterModel registerModel)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            db.RegisterModels.Add(registerModel);
            db.SaveChanges();

            return CreatedAtRoute("DefaultApi", new { id = registerModel.RegisterId }, registerModel);
        }

        // DELETE: api/RegisterModels/5
        [ResponseType(typeof(RegisterModel))]
        public IHttpActionResult DeleteRegisterModel(int id)
        {
            RegisterModel registerModel = db.RegisterModels.Find(id);
            if (registerModel == null)
            {
                return NotFound();
            }

            db.RegisterModels.Remove(registerModel);
            db.SaveChanges();

            return Ok(registerModel);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        private bool RegisterModelExists(int id)
        {
            return db.RegisterModels.Count(e => e.RegisterId == id) > 0;
        }
    }
}