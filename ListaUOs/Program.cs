
using ClosedXML.Excel;
using System.DirectoryServices;
using System.Collections.Generic;

string ldapPath = "LDAP://OU=Usuarios,OU=Sync_O365,DC=eletronuclear,DC=gov,DC=br";
string excelDirectory = @"E:\Temp\";

DirectoryEntry entry = new DirectoryEntry(ldapPath);
DirectorySearcher searcher = new DirectorySearcher(entry)
{
    Filter = "(objectClass=user)"
};

searcher.PropertiesToLoad.Add("name");
searcher.PropertiesToLoad.Add("department");

var usersByDepartment = new Dictionary<string, List<string>>();

foreach (SearchResult result in searcher.FindAll())
{
    string userName = result.Properties["name"][0].ToString();
    string department = result.Properties["department"].Count > 0 ? result.Properties["department"][0].ToString() : "Sem UO";

    if (!usersByDepartment.ContainsKey(department))
    {
        usersByDepartment[department] = new List<string>();
    }

    usersByDepartment[department].Add(userName);
}

foreach (var department in usersByDepartment)
{
    // Ordena os usuários por nome
    department.Value.Sort();

    string filePath = System.IO.Path.Combine(excelDirectory, $"{department.Key}.xlsx");

    using (var workbook = new XLWorkbook())
    {
        var worksheet = workbook.Worksheets.Add("Sheet1");
        int row = 1;
        foreach (var user in department.Value)
        {
            worksheet.Cell(row, 1).Value = user;
            row++;
        }
        workbook.SaveAs(filePath);
    }

    Console.WriteLine($"Created: {filePath}");
}
