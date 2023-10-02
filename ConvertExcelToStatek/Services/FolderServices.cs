using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConvertExcelToStatek.Services
{
    public class FolderServices : IFolderServices
    {
        private IMessageServices _messageServices;
        public FolderServices(IMessageServices messageServices)
        {
            _messageServices = messageServices;
        }
        public void ExistAndCreateFolder(string pathFolder)
        {
            if (!Directory.Exists(pathFolder))
            {
                Directory.CreateDirectory(pathFolder);
                _messageServices.SendMessage($"Созадана директоия: {pathFolder}");
            }
        }
    }
}
