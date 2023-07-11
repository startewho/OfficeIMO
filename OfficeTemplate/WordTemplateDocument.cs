using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using OfficeIMO.Word;

namespace OfficeTemplate {



    public class WordTemplateDocument : AbstractDocument<WordTemplateDocument> {

        private WordDocument _wordDocument;
        public override byte[] AsBuffer() {
            using (var ms = new MemoryStream()) {
                _wordDocument.Save(ms);
                return ms.ToArray();
            }
        }

        public override void Load(Stream inStream) {
            _wordDocument = WordDocument.Load(inStream, false, false);
            this.OnLoaded();
        }

        public override async Task LoadAsync(Stream inStream) {
            await Task.Factory.StartNew(() => Load(inStream));
        }

        public override void Save(Stream outStream) {
            _wordDocument.Save(outStream);
        }

        public override async Task SaveAsync(Stream outStream) {
            await Task.Factory.StartNew(() => this.Save(outStream));
        }
    }
}
