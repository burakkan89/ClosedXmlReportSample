using ClosedXML.Report;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClosedXml.CORE
{
    public class CreateExcel : ITemplateExcel
    {
        public string CreateExcelWithTemplate(string TemplatePath, string TemlateName, string nameManager, dynamic[] model, string outputPath)
        {
            var template = new XLTemplate(TemplatePath + TemlateName);

            template.AddVariable(nameManager, model);

            template.Generate();
            string outputFile = outputPath + Guid.NewGuid().ToString() + ".xlsx";

            template.SaveAs(outputFile);

            return outputFile;
        }
    }
}
