using System.Web.Mvc;

namespace LeagueSportsManager.Areas.Sport
{
    public class SportAreaRegistration : AreaRegistration 
    {
        public override string AreaName 
        {
            get 
            {
                return "Sport";
            }
        }

        public override void RegisterArea(AreaRegistrationContext context) 
        {
            context.MapRoute(
                "Sport_default",
                "Sport/{controller}/{action}/{id}",
                new { action = "Index", id = UrlParameter.Optional }
            );
        }
    }
}