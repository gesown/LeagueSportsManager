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
using LeagueSportsManager.Areas.Competition.Models;

namespace LeagueSportsManager.Areas.Competition.Controllers
{
    public class CompetitionModelsController : ApiController
    {
        private LeagueSportsManager db = new LeagueSportsManager();

        // GET: api/CompetitionModels
        public IQueryable<CompetitionModel> GetCompetitionModels()
        {
            return db.CompetitionModels;
        }

        // GET: api/CompetitionModels/5
        [ResponseType(typeof(CompetitionModel))]
        public IHttpActionResult GetCompetitionModel(int id)
        {
            CompetitionModel competitionModel = db.CompetitionModels.Find(id);
            if (competitionModel == null)
            {
                return NotFound();
            }

            return Ok(competitionModel);
        }

        // PUT: api/CompetitionModels/5
        [ResponseType(typeof(void))]
        public IHttpActionResult PutCompetitionModel(int id, CompetitionModel competitionModel)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (id != competitionModel.CompetitionId)
            {
                return BadRequest();
            }

            db.Entry(competitionModel).State = EntityState.Modified;

            try
            {
                db.SaveChanges();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!CompetitionModelExists(id))
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

        // POST: api/CompetitionModels
        [ResponseType(typeof(CompetitionModel))]
        public IHttpActionResult PostCompetitionModel(CompetitionModel competitionModel)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            db.CompetitionModels.Add(competitionModel);
            db.SaveChanges();

            return CreatedAtRoute("DefaultApi", new { id = competitionModel.CompetitionId }, competitionModel);
        }

        // DELETE: api/CompetitionModels/5
        [ResponseType(typeof(CompetitionModel))]
        public IHttpActionResult DeleteCompetitionModel(int id)
        {
            CompetitionModel competitionModel = db.CompetitionModels.Find(id);
            if (competitionModel == null)
            {
                return NotFound();
            }

            db.CompetitionModels.Remove(competitionModel);
            db.SaveChanges();

            return Ok(competitionModel);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        private bool CompetitionModelExists(int id)
        {
            return db.CompetitionModels.Count(e => e.CompetitionId == id) > 0;
        }
    }
}