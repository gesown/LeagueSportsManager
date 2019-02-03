using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace LeagueSportsManager.Areas.Result.Models
{
    [Table("Result")]
    public class ResultModel
    {
        [Key]
        public int ResultId { get; set; }
        public string Name { get; set; }
    }
}