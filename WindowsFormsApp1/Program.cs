using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using WindowsFormsApp1.Properties;

namespace WindowsFormsApp1
{

    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            List<Employee> emp = new List<Employee>();

            // Adding customer information to the list
            emp.Add(new Employee { Name = "Bhavin Nirmal", Email = "BhavinN@caryacalgary.ca", AccountNumber = "0032" });
            emp.Add(new Employee { Name = "Bhavin Nirmal 1", Email = "BhavinN@caryacalgary.ca", AccountNumber = "1147" });
            emp.Add(new Employee { Name = "Bhavin Nirmal 2", Email = "BhavinN@caryacalgary.ca", AccountNumber = "3545" });
            
            // Displaying customer information
            Console.WriteLine("List of Customers:");
            foreach (Employee e in emp)
            {
                Console.WriteLine($"Name: {e.Name}");
                Console.WriteLine($"Email: {e.Email}");
                Console.WriteLine($"Account Number: {e.AccountNumber}");
                Console.WriteLine();
            }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
