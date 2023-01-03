using Dapper;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using TestExcelCore.Models;

namespace TestExcelCore.Repository
{
    public class EmployeeRepo: IEmployeeRepo
    {
        private readonly DapperContext _context;
        public EmployeeRepo(DapperContext context)
        {
            _context = context;
        }

        public void AddEmployee(List<Employee> employees)
        {
            for (int i = 0; i < employees.Count; i++)
            {
                string conn = "Data Source=server6;Initial Catalog=CachDB17;User ID=user17;Password=csmpl#1234;Encrypt=false"; //ConfigurationManager.ConnectionStrings["dbconnection"].ConnectionString;
                SqlConnection con = new SqlConnection(conn);
                string query = "  Insert into [CachDB17].[dbo].[Employee] (Id,Name,Email,Mobile,Specialization) Values('" + Convert.ToInt32(employees[i].Id) + "','" + employees[i].Name.ToString() + "','" + employees[i].Email.ToString() + "','" + employees[i].Mobile.ToString() + "','" + employees[i].Specialization.ToString() + "')";
                con.Open();
                SqlCommand cmd = new SqlCommand(query, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
        }

        public async Task<List<EmployeeDetails>> Get()
        {
            List<EmployeeDetails> EmpList = null;
            try
            {
                using (var connection = _context.CreateConnection())
                {
                    DynamicParameters dynamicParam = new DynamicParameters();
                    dynamicParam.Add("@Action", "GetAll");
                    var result = await connection.QueryAsync<EmployeeDetails>("Proc_Employee", dynamicParam, commandType: CommandType.StoredProcedure);
                    return result.ToList();
                }
            }
            catch (Exception ex)
            {
                ex.ToString();
            }
            return EmpList;
        }
    }
    public class DapperContext
    {
        private readonly IConfiguration _configuration;
        private readonly string _connectionString;
        public DapperContext(IConfiguration configuration)
        {
            _configuration = configuration;
            _connectionString = _configuration.GetConnectionString("dbconnection");
        }
        public IDbConnection CreateConnection()
            => new SqlConnection(_connectionString);
    }
}
