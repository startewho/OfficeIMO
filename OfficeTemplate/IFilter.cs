using Fluid;
using Fluid.Values;
using System.Threading.Tasks;

namespace OfficeTemplate {
    public interface IFilter {
        string Name { get; }
    }

    public interface IAsyncFilter : IFilter {
        ValueTask<FluidValue> ExecuteAsync(FluidValue input, FilterArguments arguments, TemplateContext context);
    }

}
