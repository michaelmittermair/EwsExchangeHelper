using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using Microsoft.Exchange.WebServices.Data;

namespace EwsExchangeHelper.ExchangeItems
{
    /// <remarks/>
    [Serializable]
    [XmlType(AnonymousType = true, Namespace = "CategoryList.xsd")]
    [XmlRoot(ElementName = "categories", Namespace = "CategoryList.xsd", IsNullable = false)]
    public class MasterCategoryList
    {
        private UserConfiguration _userConfigurationItem;

        /// <remarks/>
        [XmlElement("category")]
        public List<Category> Categories { get; set; }

        /// <remarks/>
        [XmlIgnore]
        public Guid? DefaultCategory { get; set; }

        [XmlAttribute("default")]
        [EditorBrowsable(EditorBrowsableState.Never)]
        public string DefaultCategoryText
        {
            get => DefaultCategory != null ? DefaultCategory.ToString() : string.Empty;
            set => DefaultCategory = Guid.TryParse(value, out var result) ? result : (Guid?)null;
        }


        /// <remarks/>
        [XmlAttribute("lastSavedSession")]
        public int LastSavedSession { get; set; }

        /// <remarks/>
        [XmlAttribute("lastSavedTime")]
        public DateTime LastSavedTime { get; set; }

        public static MasterCategoryList Bind(ExchangeService service)
        {
            var item = UserConfiguration.Bind(service, "CategoryList", WellKnownFolderName.Contacts,
                                               UserConfigurationProperties.XmlData);

            var reader = new StreamReader(new MemoryStream(item.XmlData), Encoding.UTF8, true);
            var serializer = new XmlSerializer(typeof(MasterCategoryList));
            var result = (MasterCategoryList)serializer.Deserialize(reader);
            result._userConfigurationItem = item;
            return result;
        }

        public void Update()
        {
            var stream = new MemoryStream();
            var writer = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = Encoding.UTF8 });
            var serializer = new XmlSerializer(typeof(MasterCategoryList));

            serializer.Serialize(writer, this);
            writer.Flush();
            _userConfigurationItem.XmlData = stream.ToArray();
            _userConfigurationItem.Update();
        }

        public static MasterCategoryList BindOrCreate(ExchangeService service)
        {
            UserConfiguration userConfiguration;
            MasterCategoryList result;

            try
            {
                userConfiguration = UserConfiguration.Bind(service, "CategoryList", WellKnownFolderName.Contacts, UserConfigurationProperties.XmlData);
            }
            catch
            {
                result = new MasterCategoryList
                {
                    Categories = new List<Category>(),
                    _userConfigurationItem = new UserConfiguration(service)
                };
                result._userConfigurationItem.Save("CategoryList", WellKnownFolderName.Contacts);
                return result;
            }

            if (userConfiguration.XmlData != null)
            {
                var reader = new StreamReader(new MemoryStream(userConfiguration.XmlData), Encoding.UTF8, true);
                var serializer = new XmlSerializer(typeof(MasterCategoryList));
                result = (MasterCategoryList)serializer.Deserialize(reader);
            }
            else
            {
                result = new MasterCategoryList();
            }

            result._userConfigurationItem = userConfiguration;

            if (result.Categories == null)
                result.Categories = new List<Category>();

            return result;
        }
    }
}