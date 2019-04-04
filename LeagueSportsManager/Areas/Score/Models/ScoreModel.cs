using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace LeagueSportsManager.Areas.Score.Models
{
    [Table("Score")]
    public class ScoreModel
    {
        [Key]
        public int ScoreId { get; set; }
        public string Name { get; set; }
    }
}