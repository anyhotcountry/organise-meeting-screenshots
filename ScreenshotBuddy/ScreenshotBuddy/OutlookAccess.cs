using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Outlook;

namespace ScreenshotBuddy
{
    public class OutlookAccess
    {
        private readonly Application app;
        private readonly MAPIFolder folder;

        public OutlookAccess()
        {
            app = new Application();
            folder = app.Session.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
        }

        public IEnumerable<string> GetEvents(DateTime time)
        {
            var restriction = $"[Start] < '{time:g}' AND [End] > '{time:g}'";
            var items = folder.Items;
            items.IncludeRecurrences = true;
            items.Sort("[Start]");
            var itemsInDateRange = items.Restrict(restriction);
            var finalItems = itemsInDateRange;
            finalItems.Sort("[Start]");
            foreach (AppointmentItem item in finalItems)
            {
                if (!item.AllDayEvent)
                {
                    yield return item.Subject;
                }
            }
        }
    }
}
