using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpecFlowTableCreator.interfaces
{
    public interface FileConnection
    {
        FileConnection ForFile(string filePath);
        DataSet GetAllTables();
    }
}
