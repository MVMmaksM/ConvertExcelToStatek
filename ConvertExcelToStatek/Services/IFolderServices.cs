using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConvertExcelToStatek.Services
{
    public interface IFolderServices
    {
        void ExistAndCreateFolder(string pathFolder);
    }
}
