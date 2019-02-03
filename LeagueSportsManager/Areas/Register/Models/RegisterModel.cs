using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Security.AccessControl;
using LeagueSportsManager.Areas.Sport.Models;

namespace LeagueSportsManager.Areas.Register.Models
{
    [Table("Register")]
    public class RegisterModel
    {
        [Key]
        public int RegisterId { get; set; }
        public ContactView Contact { get; set; }  
        public IList<SportModel> Sports { get; set; }
        public LMStatus Status { get; set; }
    }

    public enum LMStatus
    {
        Active,
        New,
        InActive,
        Approved,
        NotApproved
    }
    [Table("Contact")]
    public class ContactView
    {
        [Key]
        public int ContactId { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string ASpNetUserId { get; set; }
    }
}