using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXmlLab
{
	public class CustomProp
	{
		public string FileLocation { get; set; }
		public CustomProp(string fileLocation)
		{
			this.FileLocation = fileLocation;
		}

		public string SetCustomProperty(CustomPropertyEntity cProp)
		{
			return SetCustomProperty(cProp.Name, cProp.Value, PropertyTypes.Text);
		}


		public string SetCustomProperty(string propertyName, 
										object propertyValue,
										PropertyTypes propertyType)
		{

			// Given a document name, a property name/value, and the property type, 
			// add a custom property to a document. The method returns the original
			// value, if it existed.

			string returnValue = null;

			var newProp = new CustomDocumentProperty();
			bool propSet = false;

			// Calculate the correct type.
			switch (propertyType)
			{
				case PropertyTypes.DateTime:

					// Be sure you were passed a real date, 
					// and if so, format in the correct way. 
					// The date/time value passed in should 
					// represent a UTC date/time.
					if ((propertyValue) is DateTime)
					{
						newProp.VTFileTime = new VTFileTime(string.Format("{0:s}Z", Convert.ToDateTime(propertyValue)));
						propSet = true;
					}
					break;

				case PropertyTypes.NumberInteger:
					if ((propertyValue) is int)
					{
						newProp.VTInt32 = new VTInt32(propertyValue.ToString());
						propSet = true;
					}
					break;

				case PropertyTypes.NumberDouble:
					if (propertyValue is double)
					{
						newProp.VTFloat = new VTFloat(propertyValue.ToString());
						propSet = true;
					}
					break;

				case PropertyTypes.Text:
					newProp.VTLPWSTR = new VTLPWSTR(propertyValue.ToString());
					propSet = true;

					break;

				case PropertyTypes.YesNo:
					if (propertyValue is bool)
					{
						// Must be lowercase.
						newProp.VTBool = new VTBool(
						  Convert.ToBoolean(propertyValue).ToString().ToLower());
						propSet = true;
					}
					break;
			}

			if (!propSet)
			{
				// If the code was not able to convert the 
				// property to a valid value, throw an exception.
				throw new InvalidDataException("propertyValue");
			}

			// Now that you have handled the parameters, start
			// working on the document.
			newProp.FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
			newProp.Name = propertyName;

			using (var document = WordprocessingDocument.Open(this.FileLocation, true))
			{
				var customProps = document.CustomFilePropertiesPart;
				if (customProps == null)
				{
					// No custom properties? Add the part, and the
					// collection of properties now.
					customProps = document.AddCustomFilePropertiesPart();
					customProps.Properties = new DocumentFormat.OpenXml.CustomProperties.Properties();
				}

				var props = customProps.Properties;
				if (props != null)
				{
					// This will trigger an exception if the property's Name 
					// property is null, but if that happens, the property is damaged, 
					// and probably should raise an exception.
					var prop = props.Where( p => ((CustomDocumentProperty)p).Name.Value == propertyName).FirstOrDefault();

					// Does the property exist? If so, get the return value, 
					// and then delete the property.
					if (prop != null)
					{
						returnValue = prop.InnerText;
						prop.Remove();
					}

					// Append the new property, and 
					// fix up all the property ID values. 
					// The PropertyId value must start at 2.
					props.AppendChild(newProp);
					int pid = 2;
					foreach (CustomDocumentProperty item in props)
					{
						item.PropertyId = pid++;
					}
					props.Save();
				}
				SetUpdateOnOpen(document);
			}
			return returnValue;
		}

		public List<CustomPropertyEntity> GetCustomProperities()
		{
			List<CustomPropertyEntity> allCustomProperities = new List<CustomPropertyEntity>();


			// Given a document name, a property name/value, and the property type, 
			// add a custom property to a document. The method returns the original
			// value, if it existed.
						
			var cPropnewProp = new CustomDocumentProperty();
		

			// Now that you have handled the parameters, start
			// working on the document.
			//newProp.FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
			//newProp.Name = propertyName;

			using (var document = WordprocessingDocument.Open(this.FileLocation, true))
			{
				var customProps = document.CustomFilePropertiesPart;
				if (customProps == null)
				{
					// No custom properties? Add the part, and the
					// collection of properties now.
					return null;
				}

				var props = customProps.Properties;
				if (props != null)
				{
					// This will trigger an exception if the property's Name 
					// property is null, but if that happens, the property is damaged, 
					// and probably should raise an exception.
					foreach (CustomDocumentProperty p in props)
					{
						CustomPropertyEntity cProp = new CustomPropertyEntity();
						cProp.Name = p.Name.Value;
						cProp.Value = p.InnerText;
						allCustomProperities.Add(cProp);
					}
				}
				
			}
			return allCustomProperities;
		}

		public void SetUpdateOnOpen(WordprocessingDocument xmlDoc)
		{

			//Open Word Setting File
			DocumentSettingsPart settingsPart = xmlDoc.MainDocumentPart.GetPartsOfType<DocumentSettingsPart>().First();
			//Update Fields
			UpdateFieldsOnOpen updateFields = new UpdateFieldsOnOpen();
			updateFields.Val = new OnOffValue(true);

			settingsPart.Settings.PrependChild<UpdateFieldsOnOpen>(updateFields);
			settingsPart.Settings.Save();
		}
	}

	public class CustomPropertyEntity
	{
		public string Name { get; set; }
		public string Value { get; set; }

		public CustomPropertyEntity()
		{
			Name = "";
			Value = "";
		}

		public CustomPropertyEntity(string name, string value)
		{
			this.Name = name;
			this.Value = value;
		}
	}

	public enum PropertyTypes : int
	{
		YesNo,
		Text,
		DateTime,
		NumberInteger,
		NumberDouble
	}
}
