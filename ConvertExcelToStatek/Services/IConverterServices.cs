﻿using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConvertExcelToStatek.Services
{
    public interface IConverterServices
    {
        List<string>? ConvertingToStatek(ExcelWorksheet excelWorksheet);
    }
}
