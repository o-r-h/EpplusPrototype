using OfficeOpenXml;
using OfficeOpenXml.Style.XmlAccess;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using WebAppExcelSample.Classes;


namespace WebAppExcelSample.Controllers
{
    public class HomeController : Controller
    {
     

        public HomeController()
        {

        }

               
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }


        
        public ActionResult ExcelExport()
        {
            try
            {
                var excelPackage = ExampleCreateExcelEPPLUS();
                Session["DownloadExcel_FileManager"] = excelPackage.GetAsByteArray();
                return Json("", JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {

                throw;
            }

        }

        public ActionResult Download()
        {
            if (Session["DownloadExcel_FileManager"] != null)
            {
                byte[] data = Session["DownloadExcel_FileManager"] as byte[];
                return File(data, "application/octet-stream", "FileManager.xlsx");
            }
            else
            {
                return new EmptyResult();
            }
        }



        private List<ExcelCellStyle> SetExcelCellStyles()
        {
            List<ExcelCellStyle> excelCellStyles = new List<ExcelCellStyle>();
            RGBCustom rGBCustom = new RGBCustom()
            {
                RColor = 190,
                BColor = 0,
                GColor = 0
            };

            ExcelCellStyle styleHeader = new ExcelCellStyle();
            styleHeader.ExcelStyleName = "HeaderCustom";
            styleHeader.FontSize = 16;
            styleHeader.FontName = "Arial";
            styleHeader.FontBold = true;
            styleHeader.BackGroundRGB = rGBCustom;
            styleHeader.FontRGB = rGBCustom;

            excelCellStyles.Add(styleHeader);
            return excelCellStyles;
        }

        public ExcelPackage ExampleCreateExcelEPPLUS()
        {
            
            ExcelPackage expack = new ExcelPackage();
            ExcelFile xls = new ExcelFile();
            ExcelWorkSheet sheet = new ExcelWorkSheet();
            List<ExcelCellStyle> excelCellStyles = new List<ExcelCellStyle>();

            sheet.Cells = new List<Cell>();
            xls.ExcelWorkSheets = new List<ExcelWorkSheet>();
            xls.ExcelFileName = "TestExcel";
            
            sheet.Name = "Sheet nr 1";
            excelCellStyles = SetExcelCellStyles();
            
            xls.ExcelWorkSheets.Add(sheet);
            xls.ExcelWorkSheets[0].ExcelCellStyles = excelCellStyles;

            List<Cell> cellList = new List<Cell>();
            Cell cell = new Cell { ColPos = 1, RowPos = 1, Value = "TEST", Style = sheet.ExcelCellStyles[0] };
            cellList.Add(cell);
            sheet.Cells.Add(cell);
           
            cellList = ExcelHelper.CreateCellTable<Example>(2, 2, GetAllexamples()); 

            foreach (var item in cellList)
            {
                sheet.Cells.Add(item);
            }

           
            foreach (var item in xls.ExcelWorkSheets)
            {
                expack.Workbook.Worksheets.Add(item.Name);
            }

            int x = 1;
            foreach (var item in xls.ExcelWorkSheets)
            {
                foreach (var subitem in item.Cells)
                {
                    expack.Workbook.Worksheets[x].Cells[subitem.RowPos, subitem.ColPos].Value = subitem.Value;
                }
                x++;
            }


            foreach (var item in xls.ExcelWorkSheets)
            {
                foreach (var subitem in item.ExcelCellStyles)
                {
                    ExcelNamedStyleXml estilo = expack.Workbook.Styles.CreateNamedStyle(subitem.ExcelStyleName);
                    estilo.Style.Font.Size = subitem.FontSize;
                    estilo.Style.Font.Bold = subitem.FontBold;
                    estilo.Style.Font.Name = subitem.FontName;
                    estilo.Style.Font.Color.SetColor(0, subitem.FontRGB.RColor, subitem.FontRGB.GColor, subitem.FontRGB.BColor);
                    estilo.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    estilo.Style.Fill.BackgroundColor.SetColor(0, subitem.BackGroundRGB.RColor, subitem.BackGroundRGB.GColor, subitem.BackGroundRGB.BColor);
                    estilo.Style.Border.Bottom.Style = (OfficeOpenXml.Style.ExcelBorderStyle)subitem.BottomBorderStyle;
                    estilo.Style.Border.Top.Style = (OfficeOpenXml.Style.ExcelBorderStyle)subitem.TopBorderStyle;
                    estilo.Style.Border.Right.Style = (OfficeOpenXml.Style.ExcelBorderStyle)subitem.RightBorderStyle;
                    estilo.Style.Border.Left.Style = (OfficeOpenXml.Style.ExcelBorderStyle)subitem.LeftBorderStyle;
                    estilo.Style.VerticalAlignment = (OfficeOpenXml.Style.ExcelVerticalAlignment)subitem.VerticalAlignment;
                }
            }

            return expack;

        }


        private List<Example> GetAllexamples()
        {
            List<Example> cellList = new List<Example>();
            cellList.Add(new Example { Id = 1, NameExample = "Alfa", PageNumber = 95 });
            cellList.Add(new Example { Id = 2, NameExample = "Beta", PageNumber = 96 });
            cellList.Add(new Example { Id = 3, NameExample = "Delta", PageNumber = 97 });
            cellList.Add(new Example { Id = 4, NameExample = "Gamma", PageNumber = 98 });

            return cellList;
        }




    }
}