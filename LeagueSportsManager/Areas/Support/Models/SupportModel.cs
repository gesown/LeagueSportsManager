using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace LeagueSportsManager.Areas.Support.Models
{
    [Table("Support")]
    public class SupportModel
    {
        [Key]
        public int SupportId { get; set; }
        public string Name { get; set; }
        public KeyValuePair<AspNetUser,object> Content { get; set; }
    }
}