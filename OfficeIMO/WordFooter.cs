﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO {
    public class WordFooter {
        public List<WordParagraph> Paragraphs {
            get {
                List<WordParagraph> paragraphs = new List<WordParagraph>();
                if (_footer != null) {
                    var list = _footer.ChildElements.OfType<Paragraph>();
                    foreach (var paragraph in list) {
                        paragraphs.Add(new WordParagraph(_document, paragraph, false));
                    }
                }

                return paragraphs;
            }
        }
        private readonly FooterPart _footerPart;
        private readonly Footer _footer;
        private string _id;
        private WordDocument _document;
        private readonly HeaderFooterValues _type;

        //private readonly Footer _footerFirst;
        //private readonly Footer _footerDefault;
        //private readonly Footer _footerEven;

        internal WordFooter(WordDocument document, FooterReference footerReference) {
            _document = document;
            _id = footerReference.Id;
            _type = WordSection.GetType(footerReference.Type);

            var listHeaders = document._wordprocessingDocument.MainDocumentPart.FooterParts.ToList();
            foreach (FooterPart footerPart in listHeaders) {
                var id = document._wordprocessingDocument.MainDocumentPart.GetIdOfPart(footerPart);
                if (id == _id) {
                    _footerPart = footerPart;
                    _footer = footerPart.Footer;
                }
            }

            if (_type == HeaderFooterValues.Default) {
                document._currentSection.Footer.Default = this;
            } else if (_type == HeaderFooterValues.Even) {
                document._currentSection.Footer.Even = this;
            } else if (_type == HeaderFooterValues.First) {
                document._currentSection.Footer.First = this;
            } else {
                throw new InvalidOperationException("Shouldn't happen?");
            }
        }

        internal WordFooter(WordDocument document, HeaderFooterValues type, Footer footerPartFooter) {
            _document = document;
            _footer = footerPartFooter;

            //if (type == HeaderFooterValues.First) {
            //    _footerFirst = footerPartFooter;
            //} else if (type == HeaderFooterValues.Default) {
            //    _footerDefault = footerPartFooter;
            //} else if (type == HeaderFooterValues.Even) {
            //    _footerEven = footerPartFooter;
            //}
            _type = type;
        }
        public WordParagraph InsertParagraph() {
            var wordParagraph = new WordParagraph();
            //if (_type == HeaderFooterValues.First) {
            //    _footerFirst.Append(wordParagraph._paragraph);
            //} else if (_type == HeaderFooterValues.Default) {
            //    _footerDefault.Append(wordParagraph._paragraph);
            //} else if (_type == HeaderFooterValues.Even) {
            //    _footerEven.Append(wordParagraph._paragraph);
            //}
            _footer.Append(wordParagraph._paragraph);
            //this.Paragraphs.Add(wordParagraph);
            return wordParagraph;
        }
    }
}
