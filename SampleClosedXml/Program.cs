// See https://aka.ms/new-console-template for more information
using ClosedXml.CORE;

Console.WriteLine("Hello, World!");

string TemplatePath = @".\Templates\";
string Outputpath = @".\Results\";

var path = TemplatePath + "mali_Template.xlsx";

var visitors = new List<dynamic>
                {
                    new { Name = "Ali Duru", Attendance = new List<dynamic> { } },
                    new { Name = "Hasan Şaş", Attendance = new List<dynamic> {
                        new { Month = "Mart", Visits = 2 },
                        new { Month = "Nisan", Visits = 3 },
                        new { Month = "Mayıs", Visits = 7 },
                    } },
                    new { Name = "Burak Kandemir", Attendance = new List<dynamic> {
                        new { Month = "Ocak", Visits = 5 },
                        new { Month = "Mayıs", Visits = 8 },
                        new { Month = "Eylül", Visits = 6 },
                    } },
                    new { Name = "Pascal Nouma", Attendance = new List<dynamic> {
                        new { Month = "Mart", Visits = 5 },
                        new { Month = "Ekim", Visits = 8 },
                        new { Month = "Eylül", Visits = 6 },
                    } },
                };
ITemplateExcel excel = new CreateExcel();
var result = excel.CreateExcelWithTemplate(TemplatePath, "mali_Template.xlsx", "maliDeneme", visitors.ToArray(), Outputpath);


Console.WriteLine(result);