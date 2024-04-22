using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AveryLabelParser
{
    public class AveryLabelTemplate
    {
        public string Brand { get; set; }
        public string ProductCode { get; set; }
        public string ProductRange { get; set; }
        public string UnitOfMeasure { get; set; }
        public string Language { get; set; }
        public string ProductDescription { get; set; }
        public string ProductCategory { get; set; }
        public string PrinterType { get; set; }
        public double PageNo { get; set; }
        public string PageDescription { get; set; }
        public double PageWidth { get; set; }
        public double PageHeight { get; set; }
        public string PageOrientation { get; set; }
        public string PaperSize { get; set; }
        public bool? DoubleSided { get; set; }
        public bool? MirrorPrinting { get; set; }
        public double? PagePanelNo { get; set; }
        public string PanelDescription { get; set; }
        public double PanelWidth { get; set; }
        public double PanelHeight { get; set; }
        public string PanelShape { get; set; }
        public double CornerRadius { get; set; }
        public double NumberAcross { get; set; }
        public double NumberDown { get; set; }
        public double PageMarginLeft { get; set; }
        public double PageMarginTop { get; set; }
        public double HorizontalPitch { get; set; }
        public double VerticalPitch { get; set; }
    }
}
