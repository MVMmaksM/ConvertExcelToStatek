using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace ConvertExcelToStatek.Services
{
    public class ConverterServices : IConverterServices
    {
        private IMessageServices _messageServices;
        public ConverterServices(IMessageServices messageServices)
        {
            _messageServices = messageServices;
        }
        public List<string>? ConvertingToStatek(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet is null)
                return null;

            var resultConverting = new List<string>();

            _messageServices.SendMessage("Начало конвертации");
            _messageServices.SendMessage("---------------------------------------------");

            for (int r = 1; r <= excelWorksheet.Dimension.End.Row; r++)
            {
                var strBuilderConverting = new StringBuilder();

                for (int c = 1; c <= excelWorksheet.Dimension.End.Column; c++)
                {
                    strBuilderConverting.Append($"{excelWorksheet.Cells[r, c].Value}");

                    if (c != excelWorksheet.Dimension.End.Column)
                        strBuilderConverting.Append('#');
                }

                resultConverting.Add(strBuilderConverting.ToString());
                _messageServices.SendMessage(strBuilderConverting.ToString());
            }

            return resultConverting;
        }
    }
}
