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
using LeagueSportsManager.Areas.Result.Models;

namespace LeagueSportsManager.Areas.Result.Controllers
{
    public class ResultModelsController : ApiController
    {
        private LeagueSportsManager db = new LeagueSportsManager();

        // GET: api/ResultModels
        public IQueryable<ResultModel> GetResultModels()
        {
            return db.ResultModels;
        }

        // GET: api/ResultModels/5
        [ResponseType(typeof(ResultModel))]
        public IHttpActionResult GetResultModel(int id)
        {
            ResultModel resultModel = db.ResultModels.Find(id);
            if (resultModel == null)
            {
                return NotFound();
            }

            return Ok(resultModel);
        }

        // PUT: api/ResultModels/5
        [ResponseType(typeof(void))]
        public IHttpActionResult PutResultModel(int id, ResultModel resultModel)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (id != resultModel.ResultId)
            {
                return BadRequest();
            }

            db.Entry(resultModel).State = EntityState.Modified;

            try
            {
                db.SaveChanges();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!ResultModelExists(id))
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

        // POST: api/ResultModels
        [ResponseType(typeof(ResultModel))]
        public IHttpActionResult PostResultModel(ResultModel resultModel)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            db.ResultModels.Add(resultModel);
            db.SaveChanges();

            return CreatedAtRoute("DefaultApi", new { id = resultModel.ResultId }, resultModel);
        }

        // DELETE: api/ResultModels/5
        [ResponseType(typeof(ResultModel))]
        public IHttpActionResult DeleteResultModel(int id)
        {
            ResultModel resultModel = db.ResultModels.Find(id);
            if (resultModel == null)
            {
                return NotFound();
            }

            db.ResultModels.Remove(resultModel);
            db.SaveChanges();

            return Ok(resultModel);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        private bool ResultModelExists(int id)
        {
            return db.ResultModels.Count(e => e.ResultId == id) > 0;
        }
    }
}