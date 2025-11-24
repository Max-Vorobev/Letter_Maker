using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;

[Serializable]
public class ListModel
{
    [XmlArray("ArrayOfAuthor")]
    [XmlArrayItem("Author")]
    public List<Author> Authors { get; set; }

    [XmlArray("listRailRoad")]
    [XmlArrayItem("RR")]
    public List<string> rrLst { get; set; }
}

[Serializable]
public class Author
{
    [XmlElement("authorName")]
    public string Name { get; set; }

    [XmlElement("phNumber")]
    public string PhoneNumber { get; set; }

    internal SortedDictionary<string, string> spisAuthor = new SortedDictionary<string, string>();
}

[Serializable]
public class rrList
{
    [XmlElement("RR")]
    public string rr { get; set; }

    internal List<string> spisRR = new List<string>();
}