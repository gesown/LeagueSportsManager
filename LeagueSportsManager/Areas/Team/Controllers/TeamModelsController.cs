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
using LeagueSportsManager.Areas.Team.Models;

namespace LeagueSportsManager.Areas.Team.Controllers
{
    public class TeamModelsController : ApiController
    {
        private LeagueSportsManager db = new LeagueSportsManager();

        // GET: api/TeamModels
        public IQueryable<TeamModel> GetTeamModels()
        {
            return db.TeamModels;
        }

        // GET: api/TeamModels/5
        [ResponseType(typeof(TeamModel))]
        public IHttpActionResult GetTeamModel(int id)
        {
            TeamModel TeamModel = db.TeamModels.Find(id);
            if (TeamModel == null)
            {
                return NotFound();
            }

            return Ok(TeamModel);
        }

        // PUT: api/TeamModels/5
        [ResponseType(typeof(void))]
        public IHttpActionResult PutTeamModel(int id, TeamModel TeamModel)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (id != TeamModel.TeamId)
            {
                return BadRequest();
            }

            db.Entry(TeamModel).State = EntityState.Modified;

            try
            {
                db.SaveChanges();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!TeamModelExists(id))
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

        // POST: api/TeamModels
        [ResponseType(typeof(TeamModel))]
        public IHttpActionResult PostTeamModel(TeamModel TeamModel)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            db.TeamModels.Add(TeamModel);
            db.SaveChanges();

            return CreatedAtRoute("DefaultApi", new { id = TeamModel.TeamId }, TeamModel);
        }

        // DELETE: api/TeamModels/5
        [ResponseType(typeof(TeamModel))]
        public IHttpActionResult DeleteTeamModel(int id)
        {
            TeamModel TeamModel = db.TeamModels.Find(id);
            if (TeamModel == null)
            {
                return NotFound();
            }

            db.TeamModels.Remove(TeamModel);
            db.SaveChanges();

            return Ok(TeamModel);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        private bool TeamModelExists(int id)
        {
            return db.TeamModels.Count(e => e.TeamId == id) > 0;
        }
    }
}