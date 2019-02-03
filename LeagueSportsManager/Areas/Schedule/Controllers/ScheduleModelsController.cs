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
using LeagueSportsManager.Areas.Schedule.Models;

namespace LeagueSportsManager.Areas.Schedule.Controllers
{
    public class ScheduleModelsController : ApiController
    {
        private LeagueSportsManager db = new LeagueSportsManager();

        // GET: api/ScheduleModels
        public IQueryable<ScheduleModel> GetScheduleModels()
        {
            return db.ScheduleModels;
        }

        // GET: api/ScheduleModels/5
        [ResponseType(typeof(ScheduleModel))]
        public IHttpActionResult GetScheduleModel(int id)
        {
            ScheduleModel scheduleModel = db.ScheduleModels.Find(id);
            if (scheduleModel == null)
            {
                return NotFound();
            }

            return Ok(scheduleModel);
        }

        // PUT: api/ScheduleModels/5
        [ResponseType(typeof(void))]
        public IHttpActionResult PutScheduleModel(int id, ScheduleModel scheduleModel)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (id != scheduleModel.ScheduleId)
            {
                return BadRequest();
            }

            db.Entry(scheduleModel).State = EntityState.Modified;

            try
            {
                db.SaveChanges();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!ScheduleModelExists(id))
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

        // POST: api/ScheduleModels
        [ResponseType(typeof(ScheduleModel))]
        public IHttpActionResult PostScheduleModel(ScheduleModel scheduleModel)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            db.ScheduleModels.Add(scheduleModel);
            db.SaveChanges();

            return CreatedAtRoute("DefaultApi", new { id = scheduleModel.ScheduleId }, scheduleModel);
        }

        // DELETE: api/ScheduleModels/5
        [ResponseType(typeof(ScheduleModel))]
        public IHttpActionResult DeleteScheduleModel(int id)
        {
            ScheduleModel scheduleModel = db.ScheduleModels.Find(id);
            if (scheduleModel == null)
            {
                return NotFound();
            }

            db.ScheduleModels.Remove(scheduleModel);
            db.SaveChanges();

            return Ok(scheduleModel);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        private bool ScheduleModelExists(int id)
        {
            return db.ScheduleModels.Count(e => e.ScheduleId == id) > 0;
        }
    }
}