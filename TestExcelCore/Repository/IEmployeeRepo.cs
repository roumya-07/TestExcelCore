using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using TestExcelCore.Models;

namespace TestExcelCore.Repository
{
   public interface IEmployeeRepo
    {
    
        Task<List<EmployeeDetails>> Get();
        void AddEmployee(List<Employee> employees);
    }
}
