using System.Xml.Serialization;

namespace EwsExchangeHelper.ExchangeItems
{
	public enum CategoryKeyboardShortcut: byte
	{
		[XmlEnum("0")]
		None = 0,
		[XmlEnum("1")]
		CtrlF2 = 1,
		[XmlEnum("2")]
		CtrlF3 = 2,
		[XmlEnum("3")]
		CtrlF4 = 3,
		[XmlEnum("4")]
		CtrlF5 = 4,
		[XmlEnum("5")]
		CtrlF6 = 5,
		[XmlEnum("6")]
		CtrlF7 = 6,
		[XmlEnum("7")]
		CtrlF8 = 7,
		[XmlEnum("8")]
		CtrlF9 = 8,
		[XmlEnum("9")]
		CtrlF10 = 9,
		[XmlEnum("10")]
		CtrlF11 = 10,
		[XmlEnum("11")]
		CtrlF12 = 11,
	}
}