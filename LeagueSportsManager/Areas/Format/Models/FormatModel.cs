using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace LeagueSportsManager.Areas.Format.Models
{
    [Table("Format")]
    public class FormatModel
    {
        [Key]
        public int FormatId { get; set; }
        public string Name { get; set; }
    }
}