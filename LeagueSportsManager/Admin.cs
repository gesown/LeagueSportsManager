namespace LeagueSportsManager
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Admin")]
    public partial class Admin
    {
        public int AdminId { get; set; }

        public string UserName { get; set; }

        public int AdminTypeId { get; set; }

        public virtual AdminType AdminType { get; set; }
    }
}
