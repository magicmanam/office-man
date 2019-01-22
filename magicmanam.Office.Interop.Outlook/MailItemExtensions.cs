using Microsoft.Office.Interop.Outlook;

namespace magicmanam.Office.Interop.Outlook
{
    public static class MailItemExtensions
    {
        public static void Send(this MailItem mail, bool modal = false)
        {
            mail.Display(modal);
            mail.Send();
        }
    }
}
