using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConvertExcelToStatek.Services
{
    public class FileServices : IFileServices
    {
        private IMessageServices _messageServices;

        public FileServices(IMessageServices messageServices)
        {
            _messageServices = messageServices;
        }
        public ExcelWorksheet GetExcelWorkSheet(string pathFile)
        {
            _messageServices.SendMessage("Чтение Excel");

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var excelPackage = new ExcelPackage(pathFile);
            return excelPackage.Workbook.Worksheets[0];
        }
        public string[] GetPathFiles(string pathFolder)
        {
            _messageServices.SendMessage($"Получение файлов из директории: {pathFolder}");
            return Directory.GetFiles(pathFolder);
        }

        public void SaveFile(string fullName, List<string> content)
        {
            _messageServices.SendMessage("------------------------------------------------");
            _messageServices.SendMessage("Сохранение результатов");
            File.WriteAllLines(fullName, content, Encoding.GetEncoding(1251));
            _messageServices.SendMessage($"Файл сохранен: {fullName}");
        }
    }
}
