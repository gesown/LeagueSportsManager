namespace LeagueSportsManager.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class detailregister : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.Contact",
                c => new
                    {
                        ContactId = c.Int(nullable: false, identity: true),
                        FirstName = c.String(),
                        LastName = c.String(),
                        ASpNetUserId = c.String(),
                    })
                .PrimaryKey(t => t.ContactId);
            
            AddColumn("dbo.Register", "Status", c => c.Int(nullable: false));
            AddColumn("dbo.Register", "Contact_ContactId", c => c.Int());
            AddColumn("dbo.Sport", "RegisterModel_RegisterId", c => c.Int());
            CreateIndex("dbo.Register", "Contact_ContactId");
            CreateIndex("dbo.Sport", "RegisterModel_RegisterId");
            AddForeignKey("dbo.Register", "Contact_ContactId", "dbo.Contact", "ContactId");
            AddForeignKey("dbo.Sport", "RegisterModel_RegisterId", "dbo.Register", "RegisterId");
            DropColumn("dbo.Register", "Name");
        }
        
        public override void Down()
        {
            AddColumn("dbo.Register", "Name", c => c.String());
            DropForeignKey("dbo.Sport", "RegisterModel_RegisterId", "dbo.Register");
            DropForeignKey("dbo.Register", "Contact_ContactId", "dbo.Contact");
            DropIndex("dbo.Sport", new[] { "RegisterModel_RegisterId" });
            DropIndex("dbo.Register", new[] { "Contact_ContactId" });
            DropColumn("dbo.Sport", "RegisterModel_RegisterId");
            DropColumn("dbo.Register", "Contact_ContactId");
            DropColumn("dbo.Register", "Status");
            DropTable("dbo.Contact");
        }
    }
}
