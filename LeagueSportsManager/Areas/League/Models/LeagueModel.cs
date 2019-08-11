using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace LeagueSportsManager.Areas.League.Models
{
    [Table("League")]
    public class LeagueModel
    {
        [Key]
        public int LeagueId { get; set; }

        public string Name { get; set; }

        public string Description { get; set; }
    }
}