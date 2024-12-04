using ClosedXML.Excel;
using System.DirectoryServices;
using System.Collections.Generic;
using DocumentFormat.OpenXml.ExtendedProperties;

string ldapPath = "LDAP://OU=Usuarios,OU=Sync_O365,DC=eletronuclear,DC=gov,DC=br";
string excelDirectory = @"E:\Temp\";

DirectoryEntry entry = new DirectoryEntry(ldapPath);
DirectorySearcher searcher = new DirectorySearcher(entry)
{
    Filter = "(objectClass=user)"
};

searcher.PropertiesToLoad.Add("name");
searcher.PropertiesToLoad.Add("department");
searcher.PropertiesToLoad.Add("sAMAccountName");
searcher.PropertiesToLoad.Add("company");
searcher.PropertiesToLoad.Add("employeeID");

var usersByDepartment = new Dictionary<string, List<User>>();

foreach (SearchResult result in searcher.FindAll())
{
    string userName = result.Properties["name"][0].ToString();
    string samAccountName = result.Properties["sAMAccountName"][0].ToString();
    string department = result.Properties["department"].Count > 0 ? result.Properties["department"][0].ToString() : "Sem UO";
    string company = result.Properties["company"].Count > 0 ? result.Properties["company"][0].ToString() : "Sem Empresa";
    string matriculaSap = result.Properties["employeeID"].Count > 0 ? result.Properties["employeeID"][0].ToString() : "Sem Matricula";

    if (!usersByDepartment.ContainsKey(department))
    {
        usersByDepartment[department] = new List<User>();
    }

    usersByDepartment[department].Add(new User { Name = userName, SamAccountName = samAccountName, Company = company, MatriculaSAP = matriculaSap });
}



foreach (var department in usersByDepartment)
{
    // Ordena os usuários por nome
    department.Value.Sort((x, y) => x.Name.CompareTo(y.Name));

    string filePath = System.IO.Path.Combine(excelDirectory, $"{department.Key}.xlsx");

    using (var workbook = new XLWorkbook())
    {
        var worksheet = workbook.Worksheets.Add("Sheet1");
        int row = 1;
        worksheet.Cell(row, 1).Value = "Name";
        worksheet.Cell(row, 2).Value = "sAMAccountName";
        worksheet.Cell(row, 3).Value = "Empresa";
        worksheet.Cell(row, 4).Value = "Matricula SAP";
        row++;
        foreach (var user in department.Value)
        {
            worksheet.Cell(row, 1).Value = user.Name;
            worksheet.Cell(row, 2).Value = user.SamAccountName;
            worksheet.Cell(row, 3).Value = user.Company;
            worksheet.Cell(row, 4).Value = user.MatriculaSAP;
            row++;
        }
        workbook.SaveAs(filePath);
    }

    Console.WriteLine($"Created: {filePath}");
}

Console.WriteLine("Process completed.");


public class User
{
    public string Name { get; set; }
    public string SamAccountName { get; set; }
    public string Company { get; set; }
    public string MatriculaSAP { get; set; }
}