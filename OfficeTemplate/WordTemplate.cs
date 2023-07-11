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
namespace OfficeIMO;
public  class WordTemplate : AbstractTemplate<WordTemplateDocument> {
    private static readonly FluidParser parser = new FluidParser();

    public WordTemplate(WordTemplateDocument templateDocument) : base(templateDocument) {

    }

    public override async Task<WordTemplateDocument> RenderAsync(TemplateContext context) {
        var fluidContext = this.CreateFluidTemplateContext(TemplateDocument, context);
        var source = "Hello {{ Firstname }} {{ Lastname }} {{ Pic|Image}}";

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
