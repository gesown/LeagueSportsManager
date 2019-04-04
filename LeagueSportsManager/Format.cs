namespace LeagueSportsManager
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Format")]
    public partial class Format
    {
        public int FormatId { get; set; }

        public string Name { get; set; }
    }
}
