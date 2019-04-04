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
using LeagueSportsManager.Areas.Support.Models;

namespace LeagueSportsManager.Areas.Support.Controllers
{
    public class SupportModelsController : ApiController
    {
        private LeagueSportsManager db = new LeagueSportsManager();

        // GET: api/SupportModels
        public IQueryable<SupportModel> GetSupportModels()
        {
            return db.SupportModels;
        }

        // GET: api/SupportModels/5
        [ResponseType(typeof(SupportModel))]
        public IHttpActionResult GetSupportModel(int id)
        {
            SupportModel supportModel = db.SupportModels.Find(id);
            if (supportModel == null)
            {
                return NotFound();
            }

            return Ok(supportModel);
        }

        // PUT: api/SupportModels/5
        [ResponseType(typeof(void))]
        public IHttpActionResult PutSupportModel(int id, SupportModel supportModel)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (id != supportModel.SupportId)
            {
                return BadRequest();
            }

            db.Entry(supportModel).State = EntityState.Modified;

            try
            {
                db.SaveChanges();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!SupportModelExists(id))
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

        // POST: api/SupportModels
        [ResponseType(typeof(SupportModel))]
        public IHttpActionResult PostSupportModel(SupportModel supportModel)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            db.SupportModels.Add(supportModel);
            db.SaveChanges();

            return CreatedAtRoute("DefaultApi", new { id = supportModel.SupportId }, supportModel);
        }

        // DELETE: api/SupportModels/5
        [ResponseType(typeof(SupportModel))]
        public IHttpActionResult DeleteSupportModel(int id)
        {
            SupportModel supportModel = db.SupportModels.Find(id);
            if (supportModel == null)
            {
                return NotFound();
            }

            db.SupportModels.Remove(supportModel);
            db.SaveChanges();

            return Ok(supportModel);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        private bool SupportModelExists(int id)
        {
            return db.SupportModels.Count(e => e.SupportId == id) > 0;
        }
    }
}