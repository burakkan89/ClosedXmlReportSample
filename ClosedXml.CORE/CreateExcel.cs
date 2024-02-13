using ClosedXML.Excel;
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
        public string CreateExcelWithTemplate(string TemplatePath, string TemlateName, string nameManager, List<List<dynamic>> model, string outputPath)
        {
            var workbook = new XLWorkbook();


            int value = 1;
            for (int i = 0; i < model.Count; i++)
            {

                var template = new XLTemplate(TemplatePath + TemlateName); // loading same template again

                template.AddVariable(nameManager, model[i]);
                template.Generate();
                
                template.Workbook.Worksheet(1).CopyTo(workbook, "Sheet_"+  value.ToString());
         
                value++;

            }
            string outputFile = outputPath + Guid.NewGuid().ToString() + ".xlsx";

            workbook.SaveAs(outputFile);

            return outputFile;
        }
    }
}
