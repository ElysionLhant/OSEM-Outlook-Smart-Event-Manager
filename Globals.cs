using Microsoft.Office.Interop.Outlook;

namespace OSEMAddIn
{
    internal static class AddInContext
    {
        private static Application? _application;

        internal static Application? OutlookApplication => _application;

        internal static void Initialize(Application application)
        {
            _application = application;
        }

        internal static void Reset()
        {
            _application = null;
        }
    }
}
