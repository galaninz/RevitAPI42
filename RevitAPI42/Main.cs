using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Newtonsoft.Json;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPI42
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            string pipeInfo = string.Empty;

            var pipeList = new FilteredElementCollector(doc)
                .OfClass(typeof(Pipe))
                .Cast<Pipe>()
                .ToList();

            string excelPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "pipes.xlsx");

            using (FileStream stream = new FileStream(excelPath, FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("Лист1");

                int rowIndex = 0;

                foreach (Pipe pipe in pipeList)
                {
                    Parameter inDiameter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM);
                    double inDiameterDouble = inDiameter.AsDouble();
                    Parameter outDiameter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER);
                    double outDiameterDouble = outDiameter.AsDouble();
                    Parameter pipeLength = pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                    double pipeLengthDouble = pipeLength.AsDouble();

                    sheet.SetCellValue(rowIndex, columnIndex: 0, pipe.Name);
                    sheet.SetCellValue(rowIndex, columnIndex: 1, outDiameterDouble);
                    sheet.SetCellValue(rowIndex, columnIndex: 2, inDiameterDouble);
                    sheet.SetCellValue(rowIndex, columnIndex: 3, pipeLengthDouble);

                    rowIndex++;
                }
                    workbook.Write(stream);
                    workbook.Close();
            }

            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string csvPath = Path.Combine(desktopPath, "pipeInfo.csv");

            File.WriteAllText(csvPath, pipeInfo);

            return Result.Succeeded;
        }
    }
}
