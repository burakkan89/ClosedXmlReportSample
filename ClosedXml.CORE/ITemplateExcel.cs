using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClosedXml.CORE
{
    public interface ITemplateExcel
    {
     public string CreateExcelWithTemplate(string TemplatePath, string TemlateName, string nameManager , dynamic[] model, string outputPath);
}
}
