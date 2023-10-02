using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConvertExcelToStatek.Services
{
    public interface IFileServices
    {
        ExcelWorksheet GetExcelWorkSheet(string pathFile);
        void SaveFile(string fullName, List<string> content);

        string[] GetPathFiles(string pathFolder);
    }
}
