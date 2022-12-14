using System.Diagnostics;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookCalendarStaleCleaner
{
  internal static class OutlookHelper
  {

    private const int DEFAULT_SEARCH_HOURS = 1;
    private const int OUTLOOK_LAUNCH_SLEEP_INTERVAL = 15000;

    public static async Task<List<Outlook.Folder>> GetInboxes(bool autoLaunchOutlook)
    {
      var folders = new List<Outlook.Folder>();

      if (!autoLaunchOutlook && !OutlookRunning())
      {
        return folders;
      }

      Outlook.Application outlook;
      Outlook.Stores stores;

      try
      {
        await EnsureOutlookIsRunningAsync();
        
        outlook = new Outlook.Application();
        stores = outlook.Session.Stores;
      }
      catch (Exception e)
      {
        Debug.WriteLine($"Exception initializing Outlook in GetInboxes.\n{e}");

        // this can generally be ignored (it's often a COM RETRYLATER when Outlook is stuck
        // booting up etc). Regardless, there is no remediation, so let's bail.
        //
        // Note: Yes, this is a double return, yes the for loop will just exit and it will return
        //       anyway, but who knows what will get added post the for at some point in the future
        //       and what chaose that will cause. This exception is exceedingly rare so would rather
        //       fail fast.
        return folders;
      }

      // foreach on COM objects can sometimes get into weird states when encountering
      // a corrupt pst,  where null objects repeat themselves, and a foreach goes into
      // an infinite loop, so prefer traditional for
      //
      // Note: These are one-based arrays
      //       See: https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook._stores.item?view=outlook-pia#Microsoft_Office_Interop_Outlook__Stores_Item_System_Object_
      for (int i = 1; i <= stores?.Count; i++)
        {
          Outlook.Store store = null;

          try
          {
            // this is in the try since sometimes COM will freak out and throw
            // IndexOutOfRangeException even though we're < Count (corrupt pst situation)
            store = stores[i];

            // ignore public folders (causes slow Exchange calls, and we don't have a use case
            // for interactions with those)
            if (store.ExchangeStoreType == Outlook.OlExchangeStoreType.olExchangePublicFolder)
              continue;

            var folder = (Outlook.Folder)store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            System.Diagnostics.Debug.WriteLine($"Found inbox: {folder.Name} in store {store.DisplayName}");

            folders.Add(folder);
          }
          catch (Exception e)
          {
            // Not every root folder has an Inbox(for example, Public folders), so this exception can be ignored
            // This also catches cases where store cannot be retrieved from a specific index (even though the API
            // says it is there. In those cases you'll see something like:
            //  System.Runtime.InteropServices.COMException (0x80040119): Outlook cannot start because a data file to
            //  send and receive messages cannot be found. Check your settings in this Microsoft Outlook profile. <snip>
            Debug.WriteLine($"Failed to get Inbo for {store?.DisplayName} type: {store?.ExchangeStoreType}:\n{e}");
          }
        }

      return folders;
    }

    internal static bool OutlookRunning()
    {
      var outlookProcs = Process.GetProcessesByName("outlook");

      return outlookProcs.Length > 0;
    }

    private static async Task EnsureOutlookIsRunningAsync()
    {
      if (!OutlookRunning())
      {
        var startInfo = new ProcessStartInfo("outlook.exe");
        startInfo.UseShellExecute = true;

        Process.Start(startInfo);

        await Task.Delay(OUTLOOK_LAUNCH_SLEEP_INTERVAL);

        if (!OutlookRunning())
        {
          throw new TimeoutException("Timed out waiting for Outlook to launch");
        }
      }
    }



  }
}
