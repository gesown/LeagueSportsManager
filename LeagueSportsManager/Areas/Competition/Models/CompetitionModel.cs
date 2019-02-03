using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace LeagueSportsManager.Areas.Competition.Models
{
    [Table("Competition")]
    public class CompetitionModel
    {
        [Key]
        public int CompetitionId { get; set; }
        public string Name { get; set; }
    }
}
