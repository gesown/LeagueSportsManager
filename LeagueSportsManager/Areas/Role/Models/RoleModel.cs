using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;
using LeagueSportsManager.Areas.Admin.Models;

namespace LeagueSportsManager.Areas.Role.Models
{
    [Table("Role")]
    public class RoleModel
    {
        public RoleModel ()
        {
        }


        [Key]
        public int RoleId { get; set; }
        public string Name { get; set; }
        public int RoleTypeId { get; set; }
    }
    [Table("RoleType")]
    public class RoleTypeModel
    {
        [Key]
        public int RoleTypeId { get; set; }
        public string Name { get; set; }
        public int RoleId { get; set; }
    }
}