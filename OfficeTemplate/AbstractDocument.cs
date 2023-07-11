using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace OfficeTemplate
{

    public abstract class AbstractDocument<TDocument> : IDocument
        where TDocument : AbstractDocument<TDocument>, new()
    {
        public bool IsNew { get; private set; } = false;
        public abstract byte[] AsBuffer();
        public abstract void Load(Stream inStream);
        public abstract Task LoadAsync(Stream inStream);
        public abstract void Save(Stream outStream);
        public abstract Task SaveAsync(Stream outStream);

        protected virtual void OnLoaded()
        {
            this.IsNew = false;
        }

       

        public string ToBase64String() =>
            Convert.ToBase64String(this.AsBuffer());

    }

}
