using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Outlook; 

namespace ConsoleOffice
{
    class Program
    {
        static void Main(string[] args)
        {
            Application outlookApplication = null;
            NameSpace outlookNamespace = null;
            MAPIFolder inboxFolder = null;
            Items mailItems = null;

            try
            {
                outlookApplication = new Application();
                outlookNamespace = outlookApplication.GetNamespace("MAPI");
                inboxFolder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                mailItems = inboxFolder.Items;
                DemoAppointmentsInRange();
                Console.WriteLine(mailItems.Count);

                foreach (MailItem item in mailItems)
                {
                    var stringBuilder = new StringBuilder();
                    stringBuilder.AppendLine("From: " + item.SenderEmailAddress);
                    stringBuilder.AppendLine("To: " + item.To);
                    stringBuilder.AppendLine("CC: " + item.CC);
                    stringBuilder.AppendLine("");
                    stringBuilder.AppendLine("Subject: " + item.Subject);
                    stringBuilder.AppendLine(item.Body);

                    Console.WriteLine(stringBuilder);
                    Marshal.ReleaseComObject(item);
                }
            } 
            catch (System.Exception e)
            {
                Console.WriteLine("{0} Exception caught: ", e);
            }
            finally
            {
                ReleaseComObject(mailItems);
                ReleaseComObject(inboxFolder);
                ReleaseComObject(outlookNamespace);
                ReleaseComObject(outlookApplication);
            }

            Console.WriteLine("OK");
            Console.ReadKey();
        }

        private static void NewMeeting()
        {
            var outlookApplication = new Application();
            AppointmentItem agendaMeeting = (AppointmentItem)outlookApplication.CreateItem(OlItemType.olAppointmentItem);
            if (agendaMeeting != null)
            {
                agendaMeeting.MeetingStatus = OlMeetingStatus.olMeeting;
                agendaMeeting.Location = "Out of Office";
                agendaMeeting.Subject = "OOO";
                agendaMeeting.Body = "I will be out of office.";
                agendaMeeting.Start = new DateTime(2021, 9, 3, 12, 0, 0);
                //agendaMeeting.Duration = 60;
                agendaMeeting.AllDayEvent = true;
                agendaMeeting.BusyStatus = OlBusyStatus.olFree;
                agendaMeeting.ReminderSet = false;
                agendaMeeting.Sensitivity = OlSensitivity.olPrivate;
                Recipient recipient = agendaMeeting.Recipients.Add("Rajanikant Singh");
                recipient.Type = (int)OlMeetingRecipientType.olRequired;
                ((_AppointmentItem)agendaMeeting).Send();
            }
        }

        private static void NewAppointment()
        {
            var outlookApplication = new Application();
            AppointmentItem newAppointment = (AppointmentItem)outlookApplication.CreateItem(OlItemType.olAppointmentItem);
            newAppointment.Start = DateTime.Now.AddHours(2);
            newAppointment.End = DateTime.Now.AddHours(3);
            newAppointment.Location = "ConferenceRoom #2345";
            newAppointment.Body = "We will discuss progress on the group project.";
            newAppointment.AllDayEvent = false;
            newAppointment.Subject = "Group Project";
            newAppointment.Recipients.Add("Rajanikant Singh");
            Recipients sentTo = newAppointment.Recipients;
            Recipient sentInvite = null;
            sentInvite = sentTo.Add("singh.rajanikant@yahoo.com");
            sentInvite.Type = (int)OlMeetingRecipientType.olRequired;
            sentInvite = sentTo.Add("Rajanikant Singh");
            sentInvite.Type = (int)OlMeetingRecipientType.olOptional;
            sentTo.ResolveAll();
            newAppointment.Save();
        }

        private static void DemoAppointmentsInRange()
        {
            var outlookApplication = new Application();
            Folder calFolder = outlookApplication.Session.GetDefaultFolder(OlDefaultFolders.olFolderCalendar) as Folder;
            DateTime start = DateTime.Now;
            DateTime end = start.AddDays(5);
            Items rangeAppts = GetAppointmentsInRange(calFolder, start, end);
            if (rangeAppts != null)
            {
                foreach (AppointmentItem appt in rangeAppts)
                {
                    Console.WriteLine("Subject: " + appt.Subject
                        + " Start: " + appt.Start.ToString("g"));
                }
            }
        }

        private static Items GetAppointmentsInRange(Folder folder, DateTime startTime, DateTime endTime)
        {
            string filter = "[Start] >= '"
                + startTime.ToString("g")
                + "' AND [End] <= '"
                + endTime.ToString("g") + "'";
            Console.WriteLine(filter);
            try
            {
                Items calItems = folder.Items;
                calItems.IncludeRecurrences = true;
                calItems.Sort("[Start]", Type.Missing);
                Items restrictItems = calItems.Restrict(filter);
                if (restrictItems.Count > 0)
                {
                    return restrictItems;
                }
                else
                {
                    return null;
                }
            }
            catch { return null; }
        }

        private static void ReleaseComObject(object obj)
        {
            if (obj != null)
            {
                Marshal.ReleaseComObject(obj);
                obj = null;
            }
        }
    }
}
