using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;
using PipeDriveAPI;

namespace PipeDriveUpload
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("PIPEDRIVE IMPORT V6.1");
                Console.WriteLine("------------------------");
                Console.WriteLine();
                UploadToPipeDrive("8411", "newCustomer");
                //UploadToPipeDrive("M085212.xml", "newDeal");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
        }

        static void UploadToPipeDrive(string args1, string args2)
        {
            string Path = @"\\MCRAESERVER-PC\McRae Files\PipeDrive\XML\" + args1 + ".xml";

            if (args2 == "newDeal")
            {
                if (Deals.LoadNewDeal(Path, args2))
                    Deals.CreateDeal();
                else
                {
                    Deals.ReloadDeal();
                }
            }
            else if (args2 == "updateDeal")
            {
                Deals.LoadUpdateDeal(Path, args2);
                Deals.UpdateDeal();
            }
            else if (args2 == "newCustomer")
            {
                Deals.LoadNewCustomer(Path);
            }
        }
        
        static void TestingOrgs()
        {
            Deals.LoadNewOrg(@"\\MCRAESERVER-PC\McRae Files\PipeDrive\XML\1789.xml");
            Deals.CreateOrg();
        }
        static void TestingPersons()
        {
            string Path = @"\\MCRAESERVER-PC\McRae Files\PipeDrive\XML\1788.xml";
            Deals.LoadNewCustomer(Path);
            Deals.CreateCustomer();
        }
    }
}