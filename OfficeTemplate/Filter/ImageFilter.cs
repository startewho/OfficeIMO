using Fluid;
using Fluid.Values;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Word;

namespace OfficeTemplate.Filter {

    public record class ImageSource(String Uri, int? Width, int? Height);

    internal class ImageFilter : IAsyncFilter {


        private WordTemplateDocument _templateDoc;
        public ImageFilter(WordTemplateDocument wordDocument) {
            _templateDoc = wordDocument;
        }

        public string Name => "image";

        public ValueTask<FluidValue> ExecuteAsync(FluidValue input, FilterArguments arguments, TemplateContext context) {


            throw new NotImplementedException();
        }
    }
}
