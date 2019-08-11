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
using LeagueSportsManager.Areas.League.Models;

namespace LeagueSportsManager.Areas.League.Controllers
{
    public class LeagueModelsController : ApiController
    {
        private LeagueSportsManager db = new LeagueSportsManager();

        // GET: api/LeagueModels
        public IQueryable<LeagueModel> GetLeagueModels()
        {
            return db.LeagueModels;
        }

        // GET: api/LeagueModels/5
        [ResponseType(typeof(LeagueModel))]
        public IHttpActionResult GetLeagueModel(int id)
        {
            LeagueModel LeagueModel = db.LeagueModels.Find(id);
            if (LeagueModel == null)
            {
                return NotFound();
            }

            return Ok(LeagueModel);
        }

        // PUT: api/LeagueModels/5
        [ResponseType(typeof(void))]
        public IHttpActionResult PutLeagueModel(int id, LeagueModel LeagueModel)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (id != LeagueModel.LeagueId)
            {
                return BadRequest();
            }

            db.Entry(LeagueModel).State = EntityState.Modified;

            try
            {
                db.SaveChanges();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!LeagueModelExists(id))
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

        // POST: api/LeagueModels
        [ResponseType(typeof(LeagueModel))]
        public IHttpActionResult PostLeagueModel(LeagueModel LeagueModel)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            db.LeagueModels.Add(LeagueModel);
            db.SaveChanges();

            return CreatedAtRoute("DefaultApi", new { id = LeagueModel.LeagueId }, LeagueModel);
        }

        // DELETE: api/LeagueModels/5
        [ResponseType(typeof(LeagueModel))]
        public IHttpActionResult DeleteLeagueModel(int id)
        {
            LeagueModel LeagueModel = db.LeagueModels.Find(id);
            if (LeagueModel == null)
            {
                return NotFound();
            }

            db.LeagueModels.Remove(LeagueModel);
            db.SaveChanges();

            return Ok(LeagueModel);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        private bool LeagueModelExists(int id)
        {
            return db.LeagueModels.Count(e => e.LeagueId == id) > 0;
        }
    }
}