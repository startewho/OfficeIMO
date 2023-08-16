using DocumentFormat.OpenXml.InkML;
using Fluid;
using OfficeIMO;
using OfficeTemplate.Filter;

namespace OfficeTemplate {
    internal class Program {
        static void Main(string[] args) {
            using (var stream = File.Open("AdvancedDocument3.docx", FileMode.OpenOrCreate)) {
                var wordDoc = new WordTemplateDocument();
                wordDoc.Load(stream);
                var template = new WordTemplate(wordDoc);
                var model = new { Firstname = "Bill", Lastname = "Gates", Pic = new ImageSource("pic.jpg", 100, 100) };
                var context = new TemplateContext(model);

                var result = template.Render(context);

                using   var filestream = File.Open("AdvancedDocument3_save.docx", FileMode.OpenOrCreate) ;
                    wordDoc.Save(filestream);

                Console.WriteLine("All done, checkout the generated document: {0}");
            }
        }
    }
}
