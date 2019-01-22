using Microsoft.Office.Interop.Outlook;

namespace magicmanam.Office.Interop.Outlook
{
    public static class ApplicationExtensions
    {
        public static MailItem CreateMailItem(this Application app)
        {
            return (MailItem) app.CreateItem(OlItemType.olMailItem);
        }

        public static AppointmentItem CreateAppointmentItem(this Application app)
        {
            return (AppointmentItem)app.CreateItem(OlItemType.olAppointmentItem);
        }

        public static TaskItem CreateTaskItem(this Application app)
        {
            return (TaskItem)app.CreateItem(OlItemType.olTaskItem);
        }

        public static ContactItem CreateContactItem(this Application app)
        {
            return (ContactItem)app.CreateItem(OlItemType.olContactItem);
        }
    }
}
