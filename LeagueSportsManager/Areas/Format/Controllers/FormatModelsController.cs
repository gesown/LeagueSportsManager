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
using LeagueSportsManager.Areas.Format.Models;

namespace LeagueSportsManager.Areas.Format.Controllers
{
    public class FormatModelsController : ApiController
    {
        private LeagueSportsManager db = new LeagueSportsManager();

        // GET: api/FormatModels
        public IQueryable<FormatModel> GetFormatModels()
        {
            return db.FormatModels;
        }

        // GET: api/FormatModels/5
        [ResponseType(typeof(FormatModel))]
        public IHttpActionResult GetFormatModel(int id)
        {
            FormatModel formatModel = db.FormatModels.Find(id);
            if (formatModel == null)
            {
                return NotFound();
            }

            return Ok(formatModel);
        }

        // PUT: api/FormatModels/5
        [ResponseType(typeof(void))]
        public IHttpActionResult PutFormatModel(int id, FormatModel formatModel)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (id != formatModel.FormatId)
            {
                return BadRequest();
            }

            db.Entry(formatModel).State = EntityState.Modified;

            try
            {
                db.SaveChanges();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!FormatModelExists(id))
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

        // POST: api/FormatModels
        [ResponseType(typeof(FormatModel))]
        public IHttpActionResult PostFormatModel(FormatModel formatModel)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            db.FormatModels.Add(formatModel);
            db.SaveChanges();

            return CreatedAtRoute("DefaultApi", new { id = formatModel.FormatId }, formatModel);
        }

        // DELETE: api/FormatModels/5
        [ResponseType(typeof(FormatModel))]
        public IHttpActionResult DeleteFormatModel(int id)
        {
            FormatModel formatModel = db.FormatModels.Find(id);
            if (formatModel == null)
            {
                return NotFound();
            }

            db.FormatModels.Remove(formatModel);
            db.SaveChanges();

            return Ok(formatModel);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        private bool FormatModelExists(int id)
        {
            return db.FormatModels.Count(e => e.FormatId == id) > 0;
        }
    }
}