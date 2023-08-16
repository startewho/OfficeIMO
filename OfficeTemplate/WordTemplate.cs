using Fluid;
using System.IO;
using System.Linq;
using System.Text.Encodings.Web;
using System.Threading.Tasks;
using System.Collections.Generic;
using Fluid.Parser;
using OfficeTemplate;
using OfficeTemplate.Filter;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO;
public  class WordTemplate : AbstractTemplate<WordTemplateDocument> {
    private static readonly FluidParser parser = new FluidParser();

    public WordTemplate(WordTemplateDocument templateDocument) : base(templateDocument) {

    }

    public override async Task<WordTemplateDocument> RenderAsync(TemplateContext context) {

        var comParas = TemplateDocument.GetCombineParas();
        foreach (var par in comParas) {
            if (!string.IsNullOrEmpty(par.Text)) {
                Console.WriteLine($"Index:{par.Paragraphs.First().Index},Text:{par.Text}");
            }
        }

        var p= comParas.FirstOrDefault(p => p.Text == "网格划分参数");
        if (p!=null) {
            var table= p.Paragraphs.Last().AddTableBefore(5, 2);
            table.Alignment = TableRowAlignmentValues.Center;
            table.Width = 3000;
            table.WidthType = TableWidthUnitValues.Pct;
            table.Alignment = TableRowAlignmentValues.Center;
            table.ColumnWidth = new List<int>() { 1500, 3500 };

            for (int i = 0; i < 5; i++) {

               
                var cell = table.Rows[i].FirstCell.Paragraphs[0];
                cell.Text = $"项目名称{i + 1}";
                cell.FontSize = 15;
                cell.FontFamily = "SimSun";
                cell.ParagraphAlignment = JustificationValues.Center;
                cell = table.Rows[i].Cells[1].Paragraphs[0];
                cell.SetFontFamily("SimSun");
                cell.SetFontSize(15);
                cell.ParagraphAlignment = JustificationValues.Center;
                cell.Text = $"值{i + 1}";
             
            }
        }
        

        var fluidContext = this.CreateFluidTemplateContext(TemplateDocument, context);
        var source = "Hello {{ Firstname }} {{ Lastname }} {{ Pic|Image}}";
        //source = "ABCD";
        if (parser.TryParse(source, out var template, out var error)) {
            
            Console.WriteLine(template.Render(fluidContext));
        }
        return this.TemplateDocument;
    }

    protected override IEnumerable<IAsyncFilter> GetInternalSyncFilters(WordTemplateDocument document) {
        yield return new ImageFilter(document);
    }

    protected override void PrepareTemplate() {


    }


    /// <summary>
    /// Removes superfluous elements around the interpolation ( {﻿{...}} )
    ///
    /// e.g. <text:p text:style-name="P1">{{<text:span text:style-name="T2">so</text:span>.<text:span text:style-name="T2">StringValue</text:span>}}</text:p>
    ///      is transformed in
    ///      <text:p text:style-name="P1">{{so.StringValue}}</text:p>
    /// </summary>
    /// <param name="mainContentText"></param>
    /// <returns>Sanitized text</returns>
    private static string Sanitize(string mainContentText) {
        var doc = XDocument.Parse(mainContentText);

        // TODO: Is very coarse grained, can probably be refined.
        foreach (var element in doc.Descendants().Where(
            x => x.Nodes().Any(y => y.NodeType == XmlNodeType.Text && ((XText)y).Value.Contains("{{")))) {
            element.Value = element.Value;
        }

        return doc.ToString();
    }
}
