using System;
using System.Reflection;

namespace Muzi.ExcelExport.net45.DemoUI.Areas.HelpPage.ModelDescriptions
{
    public interface IModelDocumentationProvider
    {
        string GetDocumentation(MemberInfo member);

        string GetDocumentation(Type type);
    }
}