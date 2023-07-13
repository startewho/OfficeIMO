using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml.ExtendedProperties;
using OfficeIMO.Word;

namespace OfficeTemplate {

    /// <summary>
    /// 
    /// </summary>
    public class CombinePara {
        public string Text { get; set; }

        public List<WordParagraph> Paragraphs { get; set; }
    }

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

        public List<CombinePara> GetCombineParas() {
            var list = new List<CombinePara>();
            var paraGroup = _wordDocument.Paragraphs.GroupBy(p => p.Index);
            foreach (var g in paraGroup) {
                var paras = g.ToList();
                var textBuilder = new StringBuilder();
                var text = string.Join(string.Empty, paras.Select(p => p.Text));
                list.Add(new CombinePara { Text = text, Paragraphs = paras });
            }
            return list;
        }


    }
}
