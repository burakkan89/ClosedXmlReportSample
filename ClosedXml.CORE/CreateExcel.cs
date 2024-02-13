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
        private static string Extensions = ".xlsx";
        public string CreateExcelWithTemplate(string TemplatePath, string TemplateName, string nameManager, List<List<dynamic>> model, string outputPath)
        {
            var workbook = new XLWorkbook();


            int value = 1;
            for (int i = 0; i < model.Count; i++)
            {

                var template = new XLTemplate(TemplatePath + TemplateName + Extensions); // loading same template again

                template.AddVariable(nameManager, model[i]);
                template.Generate();
                
                template.Workbook.Worksheet(1).CopyTo(workbook, "Sheet_"+  value.ToString());
         
                value++;

            }
            string outputFile = outputPath + TemplateName + "_" + Guid.NewGuid().ToString() + ".xlsx";

            workbook.SaveAs(outputFile);

            return outputFile;
        }

        public string CreateExcelWithTemplateWithoutNm(string TemplatePath, string TemplateName, List<dynamic> model, string outputPath)
        {
            var workbook = new XLWorkbook();


            int value = 1;
            foreach (var item in model)
            {
                var template = new XLTemplate(TemplatePath + TemplateName + Extensions); // loading same template again

                template.AddVariable(item);
                template.Generate();

                template.Workbook.Worksheet(1).CopyTo(workbook, "Sheet_" + value.ToString());

                value++;

            }


            string outputFile = outputPath + TemplateName +"_"+ Guid.NewGuid().ToString() + ".xlsx";

            workbook.SaveAs(outputFile);

            return outputFile;
        }
    }
}
