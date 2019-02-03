using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace LeagueSportsManager.Areas.Ranking.Models
{
    [Table("Ranking")]
    public class RankingModel
    {
        [Key]
        public int RankingId { get; set; }
        public string Name { get; set; }
    }
}