using System;
using System.Xml.Serialization;

namespace EwsExchangeHelper.ExchangeItems
{
	/// <remarks/>
	[Serializable]
	[XmlType(Namespace = "CategoryList.xsd")]
	public class Category
	{
		/// <remarks/>
		[XmlAttribute("name")]
		public string Name { get; set; }

		/// <remarks/>
		[XmlAttribute("color")]
		public CategoryColor Color { get; set; }

		/// <remarks/>
		[XmlAttribute("keyboardShortcut")]
		public CategoryKeyboardShortcut KeyboardShortcut { get; set; }

		/// <remarks/>
		[XmlAttribute("usageCount")]
		public int UsageCount { get; set; }

		/// <remarks/>
		[XmlAttribute("lastTimeUsedNotes")]
		public DateTime LastTimeUsedNotes { get; set; }

		/// <remarks/>
		[XmlAttribute("lastTimeUsedJournal")]
		public DateTime LastTimeUsedJournal { get; set; }

		/// <remarks/>
		[XmlAttribute("lastTimeUsedContacts")]
		public DateTime LastTimeUsedContacts { get; set; }

		/// <remarks/>
		[XmlAttribute("lastTimeUsedTasks")]
		public DateTime LastTimeUsedTasks { get; set; }

		/// <remarks/>
		[XmlAttribute("lastTimeUsedCalendar")]
		public DateTime LastTimeUsedCalendar { get; set; }

		/// <remarks/>
		[XmlAttribute("lastTimeUsedMail")]
		public DateTime LastTimeUsedMail { get; set; }

		/// <remarks/>
		[XmlAttribute("lastTimeUsed")]
		public DateTime LastTimeUsed { get; set; }

		/// <remarks/>
		[XmlAttribute("lastSessionUsed")]
		public int LastSessionUsed { get; set; }

		/// <remarks/>
		[XmlAttribute("guid")]
		public Guid Id { get; set; }

		/// <remarks/>
		[XmlAttribute("renameOnFirstUse")]
		public byte RenameOnFirstUse { get; set; }

		/// <remarks/>
		[XmlIgnore]
		public bool RenameOnFirstUseSpecified { get; set; }


		public Category()
		{
		}

		public Category(string name, CategoryColor color, CategoryKeyboardShortcut keyboardShortcut): this()
		{
			Name = name;
			Color = color;
			KeyboardShortcut = keyboardShortcut;
			Id = Guid.NewGuid();
		}

		public Category(string name, CategoryColor color)
			: this()
		{
			Name = name;
			Color = color;
			Id = Guid.NewGuid();
		}
	}
}