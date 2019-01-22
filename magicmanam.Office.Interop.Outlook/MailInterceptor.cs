using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace magicmanam.Office.Interop.Outlook
{
    public class MailInterceptor
    {
        private readonly Application _app;

        public delegate void NewEmailHandler(MailItem mail);

        public event NewEmailHandler NewMail;
        public Predicate<MailItem> Filter { get; set; }

        public MailInterceptor(Predicate<MailItem> filter = null)
        {
            this._app = new Application();
            this.Filter = filter;

            this._app.NewMailEx += onNewEmail;
        }

        private void onNewEmail(string EntryIDCollection)
        {
            foreach (var entryId in EntryIDCollection.Split(','))
            {
                var mail = this._app.Session.GetItemFromID(EntryIDCollection, Type.Missing) as MailItem;

                if (this.Filter == null || this.Filter(mail))
                {
                    this.NewMail?.Invoke(mail);
                }
            }
        }
    }
}
