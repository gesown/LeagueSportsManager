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
using LeagueSportsManager.Areas.Event.Models;

namespace LeagueSportsManager.Areas.Event.Controllers
{
    public class EventModelsController : ApiController
    {
        private LeagueSportsManager db = new LeagueSportsManager();

        // GET: api/EventModels
        public IQueryable<EventModel> GetEventModels()
        {
            return db.EventModels;
        }

        // GET: api/EventModels/5
        [ResponseType(typeof(EventModel))]
        public IHttpActionResult GetEventModel(int id)
        {
            EventModel EventModel = db.EventModels.Find(id);
            if (EventModel == null)
            {
                return NotFound();
            }

            return Ok(EventModel);
        }

        // PUT: api/EventModels/5
        [ResponseType(typeof(void))]
        public IHttpActionResult PutEventModel(int id, EventModel EventModel)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (id != EventModel.EventId)
            {
                return BadRequest();
            }

            db.Entry(EventModel).State = EntityState.Modified;

            try
            {
                db.SaveChanges();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!EventModelExists(id))
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

        // POST: api/EventModels
        [ResponseType(typeof(EventModel))]
        public IHttpActionResult PostEventModel(EventModel EventModel)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            db.EventModels.Add(EventModel);
            db.SaveChanges();

            return CreatedAtRoute("DefaultApi", new { id = EventModel.EventId }, EventModel);
        }

        // DELETE: api/EventModels/5
        [ResponseType(typeof(EventModel))]
        public IHttpActionResult DeleteEventModel(int id)
        {
            EventModel EventModel = db.EventModels.Find(id);
            if (EventModel == null)
            {
                return NotFound();
            }

            db.EventModels.Remove(EventModel);
            db.SaveChanges();

            return Ok(EventModel);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        private bool EventModelExists(int id)
        {
            return db.EventModels.Count(e => e.EventId == id) > 0;
        }
    }
}