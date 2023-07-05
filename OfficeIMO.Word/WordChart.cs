using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using AxisId = DocumentFormat.OpenXml.Drawing.Charts.AxisId;
using Chart = DocumentFormat.OpenXml.Drawing.Charts.Chart;
using ChartSpace = DocumentFormat.OpenXml.Drawing.Charts.ChartSpace;
using DataLabels = DocumentFormat.OpenXml.Drawing.Charts.DataLabels;
using Formula = DocumentFormat.OpenXml.Drawing.Charts.Formula;
using Legend = DocumentFormat.OpenXml.Drawing.Charts.Legend;
using NumericValue = DocumentFormat.OpenXml.Drawing.Charts.NumericValue;
using PlotArea = DocumentFormat.OpenXml.Drawing.Charts.PlotArea;

namespace OfficeIMO.Word {
    public partial class WordChart {
        protected static WordDocument _document;
        protected static WordParagraph _paragraph;
        protected static ChartPart _chartPart;
        protected static Drawing _drawing;
        protected static Chart _chart;


        private const long EnglishMetricUnitsPerInch = 914400;
        private const long PixelsPerInch = 96;
        public WordChart(WordDocument document, Paragraph paragraph, Drawing drawing) {
            _document = document;
            _drawing = drawing;
        }

        private string _id {
            get {
                return _document._wordprocessingDocument.MainDocumentPart.GetIdOfPart(_chartPart);
            }
        }

        protected UInt32Value _index {
            get {
                var ids = new List<UInt32Value>();
                if (_chart != null) {
                    var lineChart = _chart.PlotArea.GetFirstChild<LineChart>();
                    var barChart = _chart.PlotArea.GetFirstChild<BarChart>();
                    var pieChart = _chart.PlotArea.GetFirstChild<PieChart>();
                    if (lineChart != null) {
                        var series = lineChart.ChildElements.OfType<LineChartSeries>();
                        foreach (var index in series) {
                            ids.Add(index.Index.Val);
                        }
                    } else if (pieChart != null) {
                        var series = pieChart.ChildElements.OfType<PieChartSeries>();
                        foreach (var index in series) {
                            ids.Add(index.Index.Val);
                        }
                    } else if (barChart != null) {
                        var series = barChart.ChildElements.OfType<BarChartSeries>();
                        foreach (var index in series) {
                            ids.Add(index.Index.Val);
                        }
                    }
                }
                if (ids.Count > 0) {
                    return ids.Max() + 1;
                } else {
                    return 0;
                }
            }
        }
        internal static CategoryAxis AddCategoryAxis() {
            CategoryAxis categoryAxis1 = new CategoryAxis();
            categoryAxis1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            AxisId axisId3 = new AxisId() { Val = (UInt32Value)148921728U };

            Scaling scaling1 = new Scaling();
            Orientation orientation1 = new Orientation() { Val = OrientationValues.MinMax };

            scaling1.Append(orientation1);
            Delete delete1 = new Delete() { Val = false };
            AxisPosition axisPosition1 = new AxisPosition() { Val = AxisPositionValues.Bottom };
            MajorTickMark majorTickMark1 = new MajorTickMark() { Val = TickMarkValues.Outside };
            MinorTickMark minorTickMark1 = new MinorTickMark() { Val = TickMarkValues.None };
            TickLabelPosition tickLabelPosition1 = new TickLabelPosition() { Val = TickLabelPositionValues.NextTo };
            CrossingAxis crossingAxis1 = new CrossingAxis() { Val = (UInt32Value)154227840U };
            Crosses crosses1 = new Crosses() { Val = CrossesValues.AutoZero };
            AutoLabeled autoLabeled1 = new AutoLabeled() { Val = true };
            LabelAlignment labelAlignment1 = new LabelAlignment() { Val = LabelAlignmentValues.Center };
            LabelOffset labelOffset1 = new LabelOffset() { Val = (UInt16Value)100U };
            NoMultiLevelLabels noMultiLevelLabels1 = new NoMultiLevelLabels() { Val = false };

            categoryAxis1.Append(axisId3);
            categoryAxis1.Append(scaling1);
            categoryAxis1.Append(delete1);
            categoryAxis1.Append(axisPosition1);
            categoryAxis1.Append(majorTickMark1);
            categoryAxis1.Append(minorTickMark1);
            categoryAxis1.Append(tickLabelPosition1);
            categoryAxis1.Append(crossingAxis1);
            categoryAxis1.Append(crosses1);
            categoryAxis1.Append(autoLabeled1);
            categoryAxis1.Append(labelAlignment1);
            categoryAxis1.Append(labelOffset1);
            categoryAxis1.Append(noMultiLevelLabels1);

            return categoryAxis1;
        }

       

        internal static ValueAxis AddValueAxis() {
            ValueAxis valueAxis1 = new ValueAxis();
            valueAxis1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            AxisId axisId4 = new AxisId() { Val = (UInt32Value)154227840U };

            Scaling scaling2 = new Scaling();
            Orientation orientation2 = new Orientation() { Val = OrientationValues.MinMax };

            scaling2.Append(orientation2);
            Delete delete2 = new Delete() { Val = false };
            AxisPosition axisPosition2 = new AxisPosition() { Val = AxisPositionValues.Left };
            DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat numberingFormat1 = new DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat() { FormatCode = "General", SourceLinked = false };
            MajorGridlines majorGridlines1 = new MajorGridlines();
            MajorTickMark majorTickMark2 = new MajorTickMark() { Val = TickMarkValues.Outside };
            MinorTickMark minorTickMark2 = new MinorTickMark() { Val = TickMarkValues.None };
            TickLabelPosition tickLabelPosition2 = new TickLabelPosition() { Val = TickLabelPositionValues.NextTo };
            CrossingAxis crossingAxis2 = new CrossingAxis() { Val = (UInt32Value)148921728U };
            Crosses crosses2 = new Crosses() { Val = CrossesValues.AutoZero };
            CrossBetween crossBetween1 = new CrossBetween() { Val = CrossBetweenValues.Between };

            valueAxis1.Append(axisId4);
            valueAxis1.Append(scaling2);
            valueAxis1.Append(delete2);
            valueAxis1.Append(axisPosition2);
            valueAxis1.Append(numberingFormat1);
            valueAxis1.Append(majorGridlines1);
            valueAxis1.Append(majorTickMark2);
            valueAxis1.Append(minorTickMark2);
            valueAxis1.Append(tickLabelPosition2);
            valueAxis1.Append(crossingAxis2);
            valueAxis1.Append(crosses2);
            valueAxis1.Append(crossBetween1);

            return valueAxis1;
        }

        internal static DataLabels AddDataLabel() {
            DataLabels dataLabels1 = new DataLabels();
            dataLabels1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            ShowLegendKey showLegendKey1 = new ShowLegendKey() { Val = false };
            ShowValue showValue1 = new ShowValue() { Val = false };
            ShowCategoryName showCategoryName1 = new ShowCategoryName() { Val = false };
            ShowSeriesName showSeriesName1 = new ShowSeriesName() { Val = false };
            ShowPercent showPercent1 = new ShowPercent() { Val = false };
            ShowBubbleSize showBubbleSize1 = new ShowBubbleSize() { Val = false };
            ShowLeaderLines showLeaderLines1 = new ShowLeaderLines() { Val = true };

            dataLabels1.Append(showLegendKey1);
            dataLabels1.Append(showValue1);
            dataLabels1.Append(showCategoryName1);
            dataLabels1.Append(showSeriesName1);
            dataLabels1.Append(showPercent1);
            dataLabels1.Append(showBubbleSize1);
            dataLabels1.Append(showLeaderLines1);
            return dataLabels1;
        }

        internal static WordParagraph InsertChart(WordDocument wordDocument, WordParagraph paragraph, Chart chart, bool roundedCorners,int width=600,int height=600) {
            ChartPart part = CreateChartPart(wordDocument, roundedCorners);
            _chartPart = part;
            var id = _document._wordprocessingDocument.MainDocumentPart.GetIdOfPart(_chartPart);

            Drawing chartDrawing = CreateChartDrawing(id,width,height);
            _drawing = chartDrawing;

            var run = new Run();
            run.Append(chartDrawing);
            paragraph._paragraph.Append(run);
            _chartPart.ChartSpace.Append(chart);

            _chart = chart;
            return paragraph;
        }

        internal static ChartPart CreateChartPart(WordDocument document, bool roundedCorners) {
            ChartPart part = document._wordprocessingDocument.MainDocumentPart.AddNewPart<ChartPart>(); //("rId1");

            ChartSpace chartSpace1 = new ChartSpace();
            chartSpace1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            chartSpace1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            chartSpace1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            part.ChartSpace = chartSpace1;
            part.ChartSpace.Append(new RoundedCorners() { Val = roundedCorners });
            return part;
        }

        internal static Chart GenerateChart() {
            Chart chart1 = new Chart();
            AutoTitleDeleted autoTitleDeleted1 = new AutoTitleDeleted() { Val = false };
            PlotArea plotArea1 = new PlotArea() { Layout = new Layout() };
            //Layout layout1 = new Layout();
            //plotArea1.Append(layout1);
            
            PlotVisibleOnly plotVisibleOnly1 = new PlotVisibleOnly() { Val = true };
            DisplayBlanksAs displayBlanksAs1 = new DisplayBlanksAs() { Val = DisplayBlanksAsValues.Gap };
            ShowDataLabelsOverMaximum showDataLabelsOverMaximum1 = new ShowDataLabelsOverMaximum() { Val = false };
            chart1.Append(autoTitleDeleted1);
            chart1.Append(plotVisibleOnly1);
            chart1.Append(displayBlanksAs1);
            chart1.Append(showDataLabelsOverMaximum1);
            chart1.Append(plotArea1);
            return chart1;
        }

        internal static Values AddValuesAxisData<T>(List<T> dataList)  {
            Formula formula3 = new Formula() { Text = "" };
            NumberReference numberReference1 = new NumberReference();
            NumberingCache numberingCache1 = new NumberingCache();
            FormatCode formatCode1 = new FormatCode() { Text = "General" };
            //PointCount pointCount2 = new PointCount() { Val = (UInt32Value)4U };
            numberingCache1.Append(formatCode1);
            var index = 0;
            foreach (var data in dataList) {
                var numericPoint = new NumericPoint() { Index = Convert.ToUInt32(index), NumericValue = new NumericValue() { Text = data.ToString() } };

                numberingCache1.Append(numericPoint);
                index++;
            }
            numberReference1.Append(formula3);
            numberReference1.Append(numberingCache1);

            Values values1 = new Values() { NumberReference = numberReference1 };
            return values1;
        }

        internal static CategoryAxisData AddCategoryAxisData(List<string> categories) {
            CategoryAxisData categoryAxisData1 = new CategoryAxisData();

            StringReference stringReference2 = new StringReference();
            Formula formula2 = new Formula() { Text = "" };

            StringCache stringCache2 = new StringCache();
            int index = 0;
            foreach (string category in categories) {
                // AddStringPoint(count, category);
                stringCache2.Append(
                    new StringPoint() { Index = Convert.ToUInt32(index), NumericValue = new DocumentFormat.OpenXml.Drawing.Charts.NumericValue() { Text = category } }
                );
                index++;
            }

            stringReference2.Append(formula2);
            stringReference2.Append(stringCache2);

            categoryAxisData1.Append(stringReference2);

            return categoryAxisData1;
        }

        internal static StringReference AddSeries(UInt32Value index, string series) {
            StringReference stringReference1 = new StringReference();

            Formula formula1 = new Formula() { Text = "" };
            NumericValue numericValue1 = new NumericValue() { Text = series };
            StringPoint stringPoint1 = new StringPoint() { Index = index };
            StringCache stringCache1 = new StringCache();

            stringPoint1.Append(numericValue1);
            stringCache1.Append(stringPoint1);
            stringReference1.Append(formula1);
            stringReference1.Append(stringCache1);
            return stringReference1;
        }

        internal static Drawing CreateChartDrawing(string id,int width=600,int height=600) {
            Drawing drawing1 = new Drawing();

            DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline inline1 = new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline();
            inline1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent extent1 = new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent() { Cx =(long)width*EnglishMetricUnitsPerInch/PixelsPerInch, Cy = (long)height * EnglishMetricUnitsPerInch / PixelsPerInch };
            DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent effectExtent1 = new DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 19050L, BottomEdge = 19050L };
            DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties docProperties1 = new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties() { Id = (UInt32Value)2U, Name = "chart" };

            DocumentFormat.OpenXml.Drawing.Graphic graphic1 = new DocumentFormat.OpenXml.Drawing.Graphic();
            graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            DocumentFormat.OpenXml.Drawing.GraphicData graphicData1 = new DocumentFormat.OpenXml.Drawing.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" };

            DocumentFormat.OpenXml.Drawing.Charts.ChartReference chartReference1 = new DocumentFormat.OpenXml.Drawing.Charts.ChartReference() { Id = id };
            chartReference1.AddNamespaceDeclaration("p6", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            chartReference1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            graphicData1.Append(chartReference1);

            graphic1.Append(graphicData1);

            inline1.Append(extent1);
            inline1.Append(effectExtent1);
            inline1.Append(docProperties1);
            inline1.Append(graphic1);

            drawing1.Append(inline1);
            return drawing1;
        }

    }
}
