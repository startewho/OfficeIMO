﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;

namespace OfficeIMO.Word {
    public class WordCustomProperties {
        private WordprocessingDocument _wordprocessingDocument = null;
        private WordDocument _document = null;
        private Properties _customProperties;

        public WordCustomProperties(WordDocument document, bool? create = null) {
            _document = document;
            _wordprocessingDocument = document._wordprocessingDocument;

            var customProps = _wordprocessingDocument.CustomFilePropertiesPart;
            if (customProps == null) {
                if (document.FileOpenAccess != FileAccess.Read) {
                    // No custom properties? Add the part, and the
                    // collection of properties now.
                    customProps = _wordprocessingDocument.AddCustomFilePropertiesPart();
                    customProps.Properties = new DocumentFormat.OpenXml.CustomProperties.Properties();
                } else {
                    throw new ArgumentException("Document is read only!");
                }
            } else {
                // we could play using add / remove, but better to just rebuild?
                //_wordprocessingDocument.DeletePart(customProps.Properties);
                //customProps = _wordprocessingDocument.AddCustomFilePropertiesPart();
                //if (create != null) {
                //   customProps.Properties = new DocumentFormat.OpenXml.CustomProperties.Properties();
                // }
            }

            _customProperties = customProps.Properties;

            document._customDocumentProperties = this;
            if (create == true) {
                CreateCustomProperty(document);
            } else {
                foreach (CustomDocumentProperty property in _customProperties.CustomFilePropertiesPart.Properties) {
                    WordCustomProperty wordCustomProperty = new WordCustomProperty(property);
                    document.CustomDocumentProperties.Add(property.Name, wordCustomProperty);
                }
            }
        }

        public CustomDocumentProperty Add(string name, object value, PropertyTypes propertyType) {
            var newProp = new CustomDocumentProperty();
            bool propSet = false;

            // Calculate the correct type.
            switch (propertyType) {
                case PropertyTypes.DateTime:
                    // Be sure you were passed a real date, 
                    // and if so, format in the correct way. 
                    // The date/time value passed in should 
                    // represent a UTC date/time.
                    if ((value) is DateTime) {
                        newProp.VTFileTime = new VTFileTime(string.Format("{0:s}Z", Convert.ToDateTime(value)));
                        propSet = true;
                    }

                    break;

                case PropertyTypes.NumberInteger:
                    if ((value) is int) {
                        newProp.VTInt32 = new VTInt32(value.ToString());
                        propSet = true;
                    }

                    break;

                case PropertyTypes.NumberDouble:
                    if (value is double) {
                        newProp.VTFloat = new VTFloat(value.ToString());
                        propSet = true;
                    }

                    break;

                case PropertyTypes.Text:
                    newProp.VTLPWSTR = new VTLPWSTR(value.ToString());
                    propSet = true;

                    break;

                case PropertyTypes.YesNo:
                    if (value is bool) {
                        // Must be lowercase.
                        newProp.VTBool = new VTBool(Convert.ToBoolean(value).ToString().ToLower());
                        propSet = true;
                    }

                    break;

                default:
                    if (value is bool) {
                        // Must be lowercase.
                        newProp.VTBool = new VTBool(Convert.ToBoolean(value).ToString().ToLower());
                        propSet = true;
                    } else if (value is string) {
                        newProp.VTLPWSTR = new VTLPWSTR(value.ToString());
                        propSet = true;
                    } else if (value is double) {
                        newProp.VTFloat = new VTFloat(value.ToString());
                        propSet = true;
                    } else if (value is int) {
                        newProp.VTInt32 = new VTInt32(value.ToString());
                        propSet = true;
                    } else if (value is DateTime) {
                        newProp.VTFileTime = new VTFileTime(string.Format("{0:s}Z", Convert.ToDateTime(value)));
                        propSet = true;
                    }

                    break;
            }

            if (!propSet) {
                // If the code was not able to convert the 
                // property to a valid value, throw an exception.
                throw new InvalidDataException("propertyValue of uknown ");
            }

            // Now that you have handled the parameters, start
            // working on the document.
            newProp.FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
            newProp.Name = name;
            return newProp;
        }

        private void Add(string name, CustomDocumentProperty newProp) {
            string returnValue = null;
            if (_customProperties != null) {
                // This will trigger an exception if the property's Name 
                // property is null, but if that happens, the property is damaged, 
                // and probably should raise an exception.
                var prop = _customProperties.Where(p => ((CustomDocumentProperty)p).Name.Value == name).FirstOrDefault();

                // Does the property exist? If so, get the return value, 
                // and then delete the property.
                if (prop != null) {
                    //returnValue = prop.InnerText;
                    prop.Remove();

                    //_customProperties.RemoveChild(prop);
                }

                // Append the new property, and 
                // fix up all the property ID values. 
                // The PropertyId value must start at 2.
                _customProperties.AppendChild(newProp);
                int pid = 2;
                foreach (CustomDocumentProperty item in _customProperties) {
                    item.PropertyId = pid++;
                }

                //this._wordprocessingDocument.CustomFilePropertiesPart.Properties.Save();
                _customProperties.Save();

                //this._wordprocessingDocument.CustomFilePropertiesPart.Properties = _customProperties;
            }

            //            return returnValue;
        }

        private void CreateCustomProperty(WordDocument document) {
            if (document.CustomDocumentProperties.Count > 0) {
                foreach (var property in document.CustomDocumentProperties.Keys) {
                    var prop = Add(property, document.CustomDocumentProperties[property].Value, document.CustomDocumentProperties[property].PropertyType);
                    Add(property, prop);
                }
            }
        }
    }
}
