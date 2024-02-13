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
var visitors2 = new List<dynamic>
                {
                    new { Name = "Veli Duru", Attendance = new List<dynamic> { } },
                    new { Name = "Kemal Şaş", Attendance = new List<dynamic> {
                        new { Month = "Mart", Visits = 2 },
                        new { Month = "Nisan", Visits = 3 },
                        new { Month = "Mayıs", Visits = 7 },
                    } },
                    new { Name = "Tarık Kandemir", Attendance = new List<dynamic> {
                        new { Month = "Ocak", Visits = 5 },
                        new { Month = "Mayıs", Visits = 8 },
                        new { Month = "Eylül", Visits = 6 },
                    } },
                    new { Name = "H Nouma", Attendance = new List<dynamic> {
                        new { Month = "Mart", Visits = 5 },
                        new { Month = "Ekim", Visits = 8 },
                        new { Month = "Eylül", Visits = 6 },
                    } },
                };

List<List<dynamic>> objectList = new List<List<dynamic>>();
objectList.Add(visitors);
objectList.Add(visitors2);

List<string> resultList = new List<string>();

ITemplateExcel excel = new CreateExcel();
resultList.Add(excel.CreateExcelWithTemplate(TemplatePath, "mali_Template", "maliDeneme", objectList, Outputpath));




var objectN = new
{
    Description = "Kişi Genel Bilgi",
    DynamicHeader = new[] { "Evlilik Durumu", "Askerlik" },
    Data = new List<dynamic> {
                        new { Name = "Naci Yılma<", Age =20,  Values = new[] {"Bekar","Yok"} ,Job ="Asker"  },
                        new { Name = "Kemalettin Cindoruk",Age =25,  Values = new[]{"Bekar","Yok"} ,Job ="Memur" },
                        new { Name = "Hasan Şaş",Age =30,  Values = new[]{"Bekar","Yok"},Job ="Futbolcu"  },
                        new { Name = "Altay",Age =35, Values = new[]{"Bekar","Yok" },Job ="Şarkıcı" } 
       }
};

var objectS = new
{
    Description = "Kişi Özel Bilgi",
    DynamicHeader = new[] { "Evlilik Durumu", "Askerlik" },
    Data = new List<dynamic> {
                        new { Name = "Ali Duru", Age =20,  Values = new[] {"Bekar","Yok"} ,Job ="Aşçı" },
                        new { Name = "Veli Cindoruk",Age =25,  Values = new[]{"Bekar","Yok"} ,Job ="Serbest Meslek" },
                        new { Name = "İlhan Mansız",Age =30,  Values = new[]{"Bekar","Yok"} ,Job ="Topçu" },
                        new { Name = "Şakira",Age =35, Values = new[]{"Bekar","Yok" } ,Job ="Popçu"}
       }
};

List<dynamic> dynamicList = new List<dynamic>();
dynamicList.Add(objectN);
dynamicList.Add(objectS);


resultList.Add(excel.CreateExcelWithTemplateWithoutNm(TemplatePath, "DynamicExcel", dynamicList, Outputpath));

foreach (var item in resultList)
{
    Console.WriteLine(item);
}


Console.ReadLine();