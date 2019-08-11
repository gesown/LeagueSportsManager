using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace LeagueSportsManager.Areas.Team.Models
{
    [Table("Team")]
    public class TeamModel
    {
        [Key]
        public int TeamId { get; set; }
        public string Name { get; set; }
    }
}