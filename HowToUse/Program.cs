using HowToUse.Models;
using SharePoint;
using System;

namespace HowToUse
{
    class Program
    {
        static void Main()
        {
            try
            {
                using (var client = new SharePointService(@"https://snconceptohg.sharepoint.com/Testumgebung", "_USER_", "_PASSWORD_"))
                {
                    // Write
                    var entry = new TestListEntry
                    {
                        Title = "Herr",
                        Erstellt_Am = DateTime.Now,
                        Name = "Mustermann"
                    };

                    var id = client.AddItemToSharePointList(entry, "TestListe");

                    // Update
                    var uentry = new TestListEntry
                    {
                        ID = id,
                        Title = "Herr",
                        Erstellt_Am = DateTime.Now,
                        Name = "Meier"
                    };
                    client.UpdateExistingItemOnSharePointList(uentry, "TestListe");

                    // Read
                    var testList = client.ReadListFromSharePoint<TestListEntry>("TestListe");

                    foreach (var item in testList)
                    {
                        Console.WriteLine(item.ToString());
                    }

                    // Delete
                    client.DeleteEntryFromSharePointListById(id, "TestListe");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
            finally
            {
                Console.ReadLine();
            }
        }

    }
}
