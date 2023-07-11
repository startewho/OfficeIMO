using Fluid;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;


namespace OfficeTemplate {
    public abstract class AbstractTemplate<TDocument>
        where TDocument : IDocument {
        private readonly TDocument _document;
        private readonly static IAsyncFilter[] s_emptySyncFilters = new IAsyncFilter[] { };

        public TDocument TemplateDocument => _document;

        public AbstractTemplate(TDocument document) {
            if (document.IsNew) {
                throw new ArgumentOutOfRangeException(nameof(document), "The template document must not be new(empty)");
            }

            _document = document;
            this.PrepareTemplate();
        }

        public TDocument Render(TemplateContext context) =>
            Task.Run(() => this.RenderAsync(context)).Result;

        public abstract Task<TDocument> RenderAsync(TemplateContext context);

        protected abstract void PrepareTemplate();

        protected static IAsyncFilter[] EmptySyncFilters => s_emptySyncFilters;



        protected virtual IEnumerable<IAsyncFilter> GetInternalSyncFilters(TDocument document) => s_emptySyncFilters;


        protected virtual FluidTemplateContext<TDocument> CreateFluidTemplateContext(TDocument document, TemplateContext context) {
            var ftc = new FluidTemplateContext<TDocument>(context,document);
            ftc.CultureInfo = context.CultureInfo;
            this.RegisterInternalFilters(document, ftc);
            return ftc;
        }

        private void RegisterInternalFilters(TDocument document, FluidTemplateContext<TDocument> templateContext) {
            foreach (IAsyncFilter filter in this.GetInternalSyncFilters(document)) {
                templateContext.Options.Filters.AddFilter(filter.Name, new FilterDelegate(filter.ExecuteAsync));
            }
        }

    }
}
