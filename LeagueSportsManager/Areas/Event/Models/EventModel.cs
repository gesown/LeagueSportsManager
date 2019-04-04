using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace LeagueSportsManager.Areas.Event.Models
{
    [Table("Event")]
    public class EventModel
    {
        [Key]
        public  int EventId { get; set; }
        public string Name { get; set; }
        public virtual int SportId { get; set; }
        public virtual int ScheduleId { get; set; }
        public virtual int FormatId { get; set; }
        public virtual int CompetitionId { get; set; }
        public IDictionary<int,int> RoleUser { get; set; } 
        public IList<int> SupportId { get; set; }  
    }
}