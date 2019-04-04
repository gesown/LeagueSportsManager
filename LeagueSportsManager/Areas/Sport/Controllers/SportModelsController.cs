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
using LeagueSportsManager.Areas.Sport.Models;

namespace LeagueSportsManager.Areas.Sport.Controllers
{
    public class SportModelsController : ApiController
    {
        private LeagueSportsManager db = new LeagueSportsManager();

        // GET: api/SportModels
        public IQueryable<SportModel> GetSportModels()
        {
            return db.SportModels;
        }

        // GET: api/SportModels/5
        [ResponseType(typeof(SportModel))]
        public IHttpActionResult GetSportModel(int id)
        {
            SportModel sportModel = db.SportModels.Find(id);
            if (sportModel == null)
            {
                return NotFound();
            }

            return Ok(sportModel);
        }

        // PUT: api/SportModels/5
        [ResponseType(typeof(void))]
        public IHttpActionResult PutSportModel(int id, SportModel sportModel)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (id != sportModel.SportId)
            {
                return BadRequest();
            }

            db.Entry(sportModel).State = EntityState.Modified;

            try
            {
                db.SaveChanges();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!SportModelExists(id))
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

        // POST: api/SportModels
        [ResponseType(typeof(SportModel))]
        public IHttpActionResult PostSportModel(SportModel sportModel)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            db.SportModels.Add(sportModel);
            db.SaveChanges();

            return CreatedAtRoute("DefaultApi", new { id = sportModel.SportId }, sportModel);
        }

        // DELETE: api/SportModels/5
        [ResponseType(typeof(SportModel))]
        public IHttpActionResult DeleteSportModel(int id)
        {
            SportModel sportModel = db.SportModels.Find(id);
            if (sportModel == null)
            {
                return NotFound();
            }

            db.SportModels.Remove(sportModel);
            db.SaveChanges();

            return Ok(sportModel);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        private bool SportModelExists(int id)
        {
            return db.SportModels.Count(e => e.SportId == id) > 0;
        }
    }
}