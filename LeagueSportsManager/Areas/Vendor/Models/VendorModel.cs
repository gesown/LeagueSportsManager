using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace LeagueSportsManager.Areas.Vendor.Models
{
    [Table("Vendor")]
    public class VendorModel
    {
        [Key]
        public int VendorId { get; set; }

        public string Name { get; set; }
    }
}