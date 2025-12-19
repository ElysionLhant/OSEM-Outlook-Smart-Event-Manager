using System;
using OSEMAddIn.Models;

namespace OSEMAddIn.ViewModels
{
    internal sealed class EmailTemplateRequestEventArgs : EventArgs
    {
        public EmailTemplateRequestEventArgs(EmailTemplateType templateType)
        {
            TemplateType = templateType;
        }

        public EmailTemplateType TemplateType { get; }
    }
}
