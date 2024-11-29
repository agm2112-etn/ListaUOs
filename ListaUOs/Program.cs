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
searcher.PropertiesToLoad.Add("sAMAccountName");

var usersByDepartment = new Dictionary<string, List<User>>();

foreach (SearchResult result in searcher.FindAll())
{
    string userName = result.Properties["name"][0].ToString();
    string samAccountName = result.Properties["sAMAccountName"][0].ToString();
    string department = result.Properties["department"].Count > 0 ? result.Properties["department"][0].ToString() : "Sem UO";

    if (!usersByDepartment.ContainsKey(department))
    {
        usersByDepartment[department] = new List<User>();
    }

    usersByDepartment[department].Add(new User { Name = userName, SamAccountName = samAccountName });
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
        row++;
        foreach (var user in department.Value)
        {
            worksheet.Cell(row, 1).Value = user.Name;
            worksheet.Cell(row, 2).Value = user.SamAccountName;
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
}