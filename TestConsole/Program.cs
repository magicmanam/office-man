using magicmanam.Office.Interop.Outlook;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            var outlookInterceptor = new MailInterceptor(delegate (MailItem m) { return m.Subject == "SuBjEcT"; });

            outlookInterceptor.NewMail += OutlookInterceptor_NewMail;

            outlookInterceptor.Filter = Program.Sample;
            outlookInterceptor.Filter = delegate (MailItem m) { return string.Equals(m.Sender.Address, "magic@man.com", StringComparison.InvariantCultureIgnoreCase); };

            var app = new Application();
            var mail = app.CreateMailItem();
            mail.Subject = "Test";
            mail.To = "mag@ic.man";
            mail.Send();

            Console.ReadKey();
        }

        private static bool Sample(MailItem mail)
        {
            return true;
        }

        private static void OutlookInterceptor_NewMail(MailItem mail)
        {
            Console.WriteLine(mail.Subject);
        }
    }
}
