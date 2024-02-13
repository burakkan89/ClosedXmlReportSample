using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClosedXml.CORE
{
    public interface ITemplateExcel
    {
     public string CreateExcelWithTemplate(string TemplatePath, string TemplateName, string nameManager , List<List<dynamic>> model, string outputPath);

     public string CreateExcelWithTemplateWithoutNm(string TemplatePath, string TemplateName, List<dynamic> model, string outputPath);
    }
}
