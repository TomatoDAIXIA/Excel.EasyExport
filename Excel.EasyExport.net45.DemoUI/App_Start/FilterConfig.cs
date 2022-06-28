using System.Web;
using System.Web.Mvc;
using Muzi.ExcelExport.net45.DemoUI;

namespace Muzi.ExcelExport.net45.DemoUI.App_Start
{
    public class FilterConfig
    {
        public static void RegisterGlobalFilters(GlobalFilterCollection filters)
        {
            filters.Add(new HandleErrorAttribute());
        }
    }
}
