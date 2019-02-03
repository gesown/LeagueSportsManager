namespace LeagueSportsManager
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Sport")]
    public partial class Sport
    {
        public int SportId { get; set; }

        public string Name { get; set; }

        public int AdminId { get; set; }

        public int? RegisterModel_RegisterId { get; set; }

        public virtual Register Register { get; set; }
    }
}
