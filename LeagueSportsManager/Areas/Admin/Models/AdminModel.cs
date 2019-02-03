using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace LeagueSportsManager.Areas.Admin.Models
{
    [Table("Admin")]
    public class AdminModel
    {
        [Key]
        public int AdminId { get; set; }

        public string UserName { get; set; }
       public int AdminTypeId { get; set; }
   }
    [Table("AdminType")]
    public class AdminTypeModel
    {
        [Key]
        public int AdminTypeId { get; set; }
        public string Name { get; set; }
    }
}