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
using LeagueSportsManager.Areas.Ranking.Models;

namespace LeagueSportsManager.Areas.Ranking.Controllers
{
    public class RankingModelsController : ApiController
    {
        private LeagueSportsManager db = new LeagueSportsManager();

        // GET: api/RankingModels
        public IQueryable<RankingModel> GetRankingModels()
        {
            return db.RankingModels;
        }

        // GET: api/RankingModels/5
        [ResponseType(typeof(RankingModel))]
        public IHttpActionResult GetRankingModel(int id)
        {
            RankingModel rankingModel = db.RankingModels.Find(id);
            if (rankingModel == null)
            {
                return NotFound();
            }

            return Ok(rankingModel);
        }

        // PUT: api/RankingModels/5
        [ResponseType(typeof(void))]
        public IHttpActionResult PutRankingModel(int id, RankingModel rankingModel)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (id != rankingModel.RankingId)
            {
                return BadRequest();
            }

            db.Entry(rankingModel).State = EntityState.Modified;

            try
            {
                db.SaveChanges();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!RankingModelExists(id))
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

        // POST: api/RankingModels
        [ResponseType(typeof(RankingModel))]
        public IHttpActionResult PostRankingModel(RankingModel rankingModel)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            db.RankingModels.Add(rankingModel);
            db.SaveChanges();

            return CreatedAtRoute("DefaultApi", new { id = rankingModel.RankingId }, rankingModel);
        }

        // DELETE: api/RankingModels/5
        [ResponseType(typeof(RankingModel))]
        public IHttpActionResult DeleteRankingModel(int id)
        {
            RankingModel rankingModel = db.RankingModels.Find(id);
            if (rankingModel == null)
            {
                return NotFound();
            }

            db.RankingModels.Remove(rankingModel);
            db.SaveChanges();

            return Ok(rankingModel);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        private bool RankingModelExists(int id)
        {
            return db.RankingModels.Count(e => e.RankingId == id) > 0;
        }
    }
}