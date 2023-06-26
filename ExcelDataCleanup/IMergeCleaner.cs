using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDataCleanup
{
    internal interface IMergeCleaner
    {

        /// <summary>
        /// Unmerges all merges in the specified worksheet and ensures little to no formatting is lost in the process
        /// </summary>
        /// <param name="worksheet">the worksheet that needs its merged cells unmerged</param>
        /// <returns>true if the unmerg was sucsessfull and false otherwise</returns>
        public bool Unmerge(ExcelWorksheet worksheet);
    }
}
