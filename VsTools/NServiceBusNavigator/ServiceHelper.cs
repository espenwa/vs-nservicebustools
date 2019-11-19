using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.TextManager.Interop;

namespace NServiceBusNavigator
{
    public static class ServiceHelper
    {
        public static void FindPattern(string regexPattern, IAsyncServiceProvider serviceProvider)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            DTE dteObject = ThreadHelper.JoinableTaskFactory.Run(() =>
                serviceProvider.GetServiceAsync(typeof(DTE))
            ) as DTE;
            dteObject.Find.PatternSyntax = vsFindPatternSyntax.vsFindPatternSyntaxRegExpr;
            dteObject.Find.Action = vsFindAction.vsFindActionFindAll;
            dteObject.Find.Backwards = false;
            //dteObject.Find.ResultsLocation = vsFindResultsLocation.vsFindResultsNone;
            dteObject.Find.ResultsLocation = vsFindResultsLocation.vsFindResults1;
            dteObject.Find.MatchWholeWord = false;
            dteObject.Find.MatchInHiddenText = true;
            dteObject.Find.FindWhat = regexPattern;
            dteObject.Find.Execute();
        }

        public static string GetSelectedText(IAsyncServiceProvider serviceProvider)
        {
            object service = ThreadHelper.JoinableTaskFactory.Run(() =>
                serviceProvider.GetServiceAsync(typeof(SVsTextManager))
            );

            var textManager = service as IVsTextManager2;
            int result = textManager.GetActiveView2(1, null, (uint)_VIEWFRAMETYPE.vftCodeWindow, out IVsTextView view);
            view.GetSelectedText(out string selectedText);
            return selectedText;
        }
    }
}
