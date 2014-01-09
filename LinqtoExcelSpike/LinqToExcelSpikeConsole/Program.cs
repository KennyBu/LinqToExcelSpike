using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using LinqToExcel;

namespace LinqToExcelSpikeConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            const string excelPath = "c:/temp/excel/Employees.xlsx";

            var excel = new ExcelQueryFactory(excelPath);
            
            var employees = from e in excel.Worksheet<Employee>("Employee")
                                   select e;

            Console.WriteLine("All Employees");
            Console.WriteLine("");
            WriteInfo(employees);
            Console.WriteLine("");
            Console.WriteLine("");

            var managers = from e in excel.Worksheet<Employee>("Employee")
                            where e.IsManager
                            select e;

            Console.WriteLine("Managers");
            Console.WriteLine("");
            WriteInfo(managers);
            Console.WriteLine("");
            Console.WriteLine("");

            var regularEmployees = from e in excel.Worksheet<Employee>("Employee")
                           where e.IsManager == false
                           select e;

            Console.WriteLine("Regular Employees");
            Console.WriteLine("");
            WriteInfo(regularEmployees);
            Console.WriteLine("");
            Console.Read();
        }

        private static void WriteInfo(IEnumerable<Employee> employees)
        {
            foreach (var employee in employees)
            {
                Console.WriteLine("Id:{0} Name:{1} Email:{2} IsManager:{3} StartDate{4}",
                    employee.Id, employee.Name, employee.Email, employee.IsManager, employee.StartDate);
            }
        }
    }
}
