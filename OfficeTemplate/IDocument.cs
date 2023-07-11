using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace OfficeTemplate {
    public interface IDocument
    {
        bool IsNew { get; }
        byte[] AsBuffer();
        void Save(Stream outStream);
        Task SaveAsync(Stream outStream);
    }
}
