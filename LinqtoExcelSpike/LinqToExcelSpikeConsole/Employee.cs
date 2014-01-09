using System;

namespace LinqToExcelSpikeConsole
{
    public class Employee
    {
        public int Id { get; set; } 
        public string Name { get; set; } 
        public string Email { get; set; } 
        public bool IsManager { get; set; } 
        public DateTime StartDate { get; set; } 
    }
}