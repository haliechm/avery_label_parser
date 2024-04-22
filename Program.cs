using AveryLabelParser;
using OfficeOpenXml;
using System.Text;

/*
 * application to parse avery label template and create sql insert script
 */

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

Console.WriteLine("Enter path to Avery label template: ");
var filePath = Console.ReadLine();

if (!File.Exists(filePath))
{
    Console.WriteLine("File not found.");
    return;
}

using (ExcelPackage package = new ExcelPackage(filePath))
{
    ExcelWorksheet ws = package.Workbook.Worksheets[0];

    int rowCount = ws.Dimension.Rows;
    int colCount = ws.Dimension.Columns;

    IList<AveryLabelTemplate> templates = new List<AveryLabelTemplate>();

    for (int i = 2; i <= rowCount; i++)
    {
        var template = new AveryLabelTemplate();
        template.Brand = (string)ws.Cells["A" + i].Value;
        template.ProductCode = (string)ws.Cells["B" + i].Value;
        template.ProductRange = (string)ws.Cells["C" + i].Value;
        template.UnitOfMeasure = (string)ws.Cells["D" + i].Value;
        template.Language = (string)ws.Cells["G" + i].Value;
        template.ProductDescription = (string)ws.Cells["H" + i].Value;
        template.ProductCategory = (string)ws.Cells["K" + i].Value;
        template.PrinterType = (string)ws.Cells["L" + i].Value;
        template.PageNo = (double)ws.Cells["M" + i].Value;
        template.PageDescription = (string)ws.Cells["N" + i].Value;
        template.PageWidth = (double)ws.Cells["O" + i].Value;
        template.PageHeight = (double)ws.Cells["P" + i].Value;
        template.PageOrientation = (string)ws.Cells["Q" + i].Value;
        template.PaperSize = (string)ws.Cells["R" + i].Value;
        template.DoubleSided = (string)ws.Cells["S" + i].Value == "X";
        template.MirrorPrinting = (string)ws.Cells["T" + i].Value == "X";
        template.PagePanelNo = (double?)ws.Cells["U" + i].Value;
        template.PanelDescription = (string)ws.Cells["V" + i].Value;
        template.PanelWidth = (double)ws.Cells["W" + i].Value;
        template.PanelHeight = (double)ws.Cells["X" + i].Value;
        template.PanelShape = (string)ws.Cells["Y" + i].Value;
        template.CornerRadius = (double)ws.Cells["Z" + i].Value;
        template.NumberAcross = (double)ws.Cells["AA" + i].Value;
        template.NumberDown = (double)ws.Cells["AB" + i].Value;
        template.PageMarginLeft = (double)ws.Cells["AC" + i].Value;
        template.PageMarginTop = (double)ws.Cells["AD" + i].Value;
        template.HorizontalPitch = (double)ws.Cells["AE" + i].Value;
        template.VerticalPitch = (double)ws.Cells["AF" + i].Value;

        templates.Add(template);
    }

    templates = templates
        .Where(x =>
            x.Brand == "Avery" &&
            x.ProductRange == "US Letter" &&
            x.UnitOfMeasure == "inches" &&
            x.Language == "en" &&
            x.PageNo == 1 &&
            x.PageDescription == "Sheet" &&
            x.PageOrientation == "Portrait" &&
            x.PaperSize == "Letter" &&
            x.DoubleSided != null &&
            !x.DoubleSided.Value &&
            x.MirrorPrinting != null &&
            !x.MirrorPrinting.Value &&
            x.PagePanelNo == 1 &&
            x.PanelDescription == "Label" &&
            x.PanelWidth >= 1 &&
            x.PanelHeight >= 1 &&
            x.PanelShape == "rect"
        )
        .ToList();


    // create sql statement
    StringBuilder sb = new StringBuilder();

    sb.Append("INSERT INTO [LabelTemplate] \r\n(templateidentifier, pagesizeid, description, topmargininches, sidemargininches, verticalpitchinches, horizontalpitchinches, labelheightinches, labelwidthinches, numberacross, numberdown, createdat, createdbyuserid, updatedat, updatedbyuserid) \r\nVALUES ");

    foreach (var template in templates)
    {
        sb.AppendLine(@$"('{template.ProductCode}', 48, '{template.ProductDescription}', {template.PageMarginTop}, {template.PageMarginLeft}, {template.VerticalPitch}, {template.HorizontalPitch}, {template.PanelHeight}, {template.PanelWidth}, {template.NumberAcross}, {template.NumberDown}, getutcdate(),  -1,  getutcdate(),  -1),");

    }

    Console.WriteLine(sb.ToString());
}


