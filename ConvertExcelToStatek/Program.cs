using ConvertExcelToStatek.Services;
using System.IO;

namespace ConvertExcelToStatek
{
    internal class Program
    {
        private static string inputFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "input");
        private static string outputFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "output");
        private static IMessageServices _messageServices = new MessageServices();
        private static IFileServices _fileServices = new FileServices(_messageServices);
        private static IConverterServices _converterServices = new ConverterServices(_messageServices);
        private static IFolderServices _folderServices = new FolderServices(_messageServices);
        static void Main(string[] args)
        {
            try
            {
                _messageServices.SendMessage("---------------Старт-----------------");

                _messageServices.SendMessage("Проверка существования директорий");
                _folderServices.ExistAndCreateFolder(inputFolderPath);
                _folderServices.ExistAndCreateFolder(outputFolderPath);

                var files = _fileServices.GetPathFiles(inputFolderPath);
                _messageServices.SendMessage($"Получено файлов: {files.Length}");

                if (files.Length > 0)
                {
                    foreach (var path in files)
                    {
                        var excelWorkSheet = _fileServices.GetExcelWorkSheet(path);
                        var resultConverting = _converterServices.ConvertingToStatek(excelWorkSheet);

                        if (resultConverting is not null)
                        {
                            var nameFileWithoutExt = Path.GetFileNameWithoutExtension(path);
                            var fileNameSave = Path.Combine(outputFolderPath, $"{nameFileWithoutExt}.txt");

                            _fileServices.SaveFile(fileNameSave, resultConverting);
                        }
                    }
                }

                _messageServices.SendMessage("\n---------------Завершено-------------");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                _messageServices.SendMessage(ex.Message);
            }
        }
    }
}