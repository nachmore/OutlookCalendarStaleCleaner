using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookCalendarStaleCleaner
{

  static class Program
  {

    private static class Stats
    {
      public static int Deleted = 0;
      public static int MarkedTenative = 0;
      public static int Ignored = 0;
    }

    private static void Usage()
    {
      Console.WriteLine(@"
Simple program to clean stale calendar invites sitting in your inbox. This tool will:

  1. Delete unanswered calendar invites if the meeting happened over 24 hours ago
     -> Note: meeting will remain on your calendar, just marked as tentative
  2. Delete calendar invites that have already been accepted (perhaps manually on your calendar)
  3. Automatically process meeting cancellations
      ");
    }

    public static async Task Main(string[] args)
    {
      if (args.Length > 0) 
      {
        Usage();
        return;
      }

      var folders = await OutlookHelper.GetInboxes(false);

      foreach (var folder in folders)
      {
        Console.WriteLine($"🔃 Processing Inbox Folder: {folder.FolderPath}");

        CleanStaleItems(folder);

        Console.WriteLine($"✅ Finished Processing Inbox Folder: {folder.FolderPath}\n");
      }

      Console.WriteLine("\n---------------\n");
      Console.WriteLine("🏁 Completed!");
      Console.WriteLine($"  🙈 Ignored         : {Stats.Ignored}");
      Console.WriteLine($"  ⛺ Marked Tentative: {Stats.MarkedTenative}");
      Console.WriteLine($"  ❌ Deleted         : {Stats.Deleted}");
    }

    private static void CleanStaleItems(Outlook.Folder folder)
    {
      string filter = "([MessageClass] = 'IPM.Schedule.Meeting.Request' OR [MessageClass] = 'IPM.Schedule.Meeting.Canceled')";
      var items = folder.Items.Restrict(filter);

      foreach (var item in items)
      {
        var meeting = (item is Outlook.MeetingItem ? (Outlook.MeetingItem)item : null);

        if (meeting != null)
        {
          var appointment = meeting.GetAssociatedAppointment(true);

          // appointment will be null when it has been deleted manually from the calendar but the
          // mail item is still there
          var processed = (appointment == null);

          if (processed)
          {
            Console.WriteLine($"❌ Processing already deleted meeting");
            Console.WriteLine($"  ✉️ {meeting.Subject}");
            Console.WriteLine($"  📧 From: {meeting.SenderEmailAddress}");
          }
          else
          {
            Console.WriteLine($"✉️ {appointment.Subject}");
            Console.WriteLine($"  📧 From: {appointment.Organizer}");
            Console.WriteLine($"  📆 Scheduled: {appointment.Start} -> {appointment.End}");
            Console.WriteLine($"  🗿 Response Status: {appointment.ResponseStatus}");
            Console.WriteLine($"  🚩 Meeting Status: {appointment.MeetingStatus}");

            processed = ProcessMeetingCancellation(appointment);

            if (!processed)
            {
              processed = ProcessMeetingRequest(appointment);
            }

            if (!processed)
            {
              Stats.Ignored++;
              Console.WriteLine(" -> 🙈 Ignored");
              continue;
            }
          }

          if (processed)
          {
            // delete the original item from the inbox (will remain on the calendar)
            meeting.Delete();

            Stats.Deleted++;

            Console.WriteLine(" -> ✅ Cleaned!");
          }
        }
      }
    }

    private static bool ProcessMeetingRequest(Outlook.AppointmentItem appointment)
    {
      // don't remove anything within the last day
      if (appointment.StartUTC < DateTime.UtcNow.Subtract(TimeSpan.FromDays(1)))
      {
        // don't set to tentative if we already accepted (or organized) the meeting,
        // for example directly via our calendar, or if the meeting is cancelled
        if (appointment.ResponseStatus != Outlook.OlResponseStatus.olResponseAccepted &&
            appointment.ResponseStatus != Outlook.OlResponseStatus.olResponseOrganized)
        {
          // record a Tentative response
          var response = appointment.Respond(Outlook.OlMeetingResponse.olMeetingTentative, true, Type.Missing);

          // randomly some responses will come back null
          if (response != null)
          {
            // delete the MeetingItem that is generated as a response so that no repsonse
            // is sent to the organizer
            response.Close(Outlook.OlInspectorClose.olDiscard);

            Stats.MarkedTenative++;
          }
        }

        return true;
      }

      return false;
    }

    private static bool ProcessMeetingCancellation(Outlook.AppointmentItem appointment)
    {
      if (appointment.MeetingStatus == Outlook.OlMeetingStatus.olMeetingCanceled ||
          appointment.MeetingStatus == Outlook.OlMeetingStatus.olMeetingReceivedAndCanceled)
      {
        // process all cancellations, regardless of time
        appointment.Delete();

        return true;
      }

      return false;
    }
  }
}
