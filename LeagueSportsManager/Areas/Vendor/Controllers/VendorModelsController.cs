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
using LeagueSportsManager.Areas.Vendor.Models;

namespace LeagueSportsManager.Areas.Vendor.Controllers
{
    public class VendorModelsController : ApiController
    {
        private LeagueSportsManager db = new LeagueSportsManager();

        // GET: api/VendorModels
        public IQueryable<VendorModel> GetVendorModels()
        {
            return db.VendorModels;
        }

        // GET: api/VendorModels/5
        [ResponseType(typeof(VendorModel))]
        public IHttpActionResult GetVendorModel(int id)
        {
            VendorModel VendorModel = db.VendorModels.Find(id);
            if (VendorModel == null)
            {
                return NotFound();
            }

            return Ok(VendorModel);
        }

        // PUT: api/VendorModels/5
        [ResponseType(typeof(void))]
        public IHttpActionResult PutVendorModel(int id, VendorModel VendorModel)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (id != VendorModel.VendorId)
            {
                return BadRequest();
            }

            db.Entry(VendorModel).State = EntityState.Modified;

            try
            {
                db.SaveChanges();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!VendorModelExists(id))
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

        // POST: api/VendorModels
        [ResponseType(typeof(VendorModel))]
        public IHttpActionResult PostVendorModel(VendorModel VendorModel)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            db.VendorModels.Add(VendorModel);
            db.SaveChanges();

            return CreatedAtRoute("DefaultApi", new { id = VendorModel.VendorId }, VendorModel);
        }

        // DELETE: api/VendorModels/5
        [ResponseType(typeof(VendorModel))]
        public IHttpActionResult DeleteVendorModel(int id)
        {
            VendorModel VendorModel = db.VendorModels.Find(id);
            if (VendorModel == null)
            {
                return NotFound();
            }

            db.VendorModels.Remove(VendorModel);
            db.SaveChanges();

            return Ok(VendorModel);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        private bool VendorModelExists(int id)
        {
            return db.VendorModels.Count(e => e.VendorId == id) > 0;
        }
    }
}