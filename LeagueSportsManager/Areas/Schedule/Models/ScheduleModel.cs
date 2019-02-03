using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace LeagueSportsManager.Areas.Schedule.Models
{
    [Table("Schedule")]
    public class ScheduleModel
    {
        [Key]
        public int ScheduleId { get; set; }
        public string ScheduleType { get; set; }
    }
}