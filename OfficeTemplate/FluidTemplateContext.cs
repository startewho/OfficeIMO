using Fluid;
using System.Collections.Generic;

namespace OfficeTemplate {
    public class FluidTemplateContext<TDocument> : TemplateContext  where TDocument:IDocument {
        public TDocument Document { get; set; }

        public FluidTemplateContext(TemplateContext templateContext, TDocument document) : base(templateContext.Model,templateContext.Options) {
            Document = document;
           
        }
    }
}
