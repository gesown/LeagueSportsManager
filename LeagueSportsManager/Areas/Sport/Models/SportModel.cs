using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using LeagueSportsManager.Areas.Admin.Models;

namespace LeagueSportsManager.Areas.Sport.Models
{
    [Table("Sport")]
    public class SportModel
    {
        [Key]
        public int SportId { get; set; }
        public string Name { get; set; }
        public int AdminId { get; set; }
        public  IList<string> UserNames { get; set; } 
    }
}