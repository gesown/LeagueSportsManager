using System.Web.Mvc;

namespace LeagueSportsManager.Areas.Format
{
    public class FormatAreaRegistration : AreaRegistration 
    {
        public override string AreaName 
        {
            get 
            {
                return "Format";
            }
        }

        public override void RegisterArea(AreaRegistrationContext context) 
        {
            context.MapRoute(
                "Format_default",
                "Format/{controller}/{action}/{id}",
                new { action = "Index", id = UrlParameter.Optional }
            );
        }
    }
}