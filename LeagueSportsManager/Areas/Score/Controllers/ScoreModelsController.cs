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
using LeagueSportsManager.Areas.Score.Models;

namespace LeagueSportsManager.Areas.Score.Controllers
{
    public class ScoreModelsController : ApiController
    {
        private LeagueSportsManager db = new LeagueSportsManager();

        // GET: api/ScoreModels
        public IQueryable<ScoreModel> GetScoreModels()
        {
            return db.ScoreModels;
        }

        // GET: api/ScoreModels/5
        [ResponseType(typeof(ScoreModel))]
        public IHttpActionResult GetScoreModel(int id)
        {
            ScoreModel scoreModel = db.ScoreModels.Find(id);
            if (scoreModel == null)
            {
                return NotFound();
            }

            return Ok(scoreModel);
        }

        // PUT: api/ScoreModels/5
        [ResponseType(typeof(void))]
        public IHttpActionResult PutScoreModel(int id, ScoreModel scoreModel)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (id != scoreModel.ScoreId)
            {
                return BadRequest();
            }

            db.Entry(scoreModel).State = EntityState.Modified;

            try
            {
                db.SaveChanges();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!ScoreModelExists(id))
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

        // POST: api/ScoreModels
        [ResponseType(typeof(ScoreModel))]
        public IHttpActionResult PostScoreModel(ScoreModel scoreModel)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            db.ScoreModels.Add(scoreModel);
            db.SaveChanges();

            return CreatedAtRoute("DefaultApi", new { id = scoreModel.ScoreId }, scoreModel);
        }

        // DELETE: api/ScoreModels/5
        [ResponseType(typeof(ScoreModel))]
        public IHttpActionResult DeleteScoreModel(int id)
        {
            ScoreModel scoreModel = db.ScoreModels.Find(id);
            if (scoreModel == null)
            {
                return NotFound();
            }

            db.ScoreModels.Remove(scoreModel);
            db.SaveChanges();

            return Ok(scoreModel);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        private bool ScoreModelExists(int id)
        {
            return db.ScoreModels.Count(e => e.ScoreId == id) > 0;
        }
    }
}