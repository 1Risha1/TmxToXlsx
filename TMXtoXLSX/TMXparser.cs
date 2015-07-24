using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.IO;

namespace TMXtoXLSX
{
    class TmxParser
    {
        XmlTextReader reader;
        string path;
        XlsxWriter output;

        public TmxParser(string path, XlsxWriter output)
        {
            if (!File.Exists(path))
            {
                throw new FileNotFoundException("File " + path + " doesn't exist\n");
            }
            this.path = path;
            this.output = output;
        }
/*
        ~TmxParser()
        {
            //reader.Close();
            try
            {
                output.Close();
            }
            catch (Exception e)
            {
                System.Console.Error.WriteLine("File can't be written " + e.ToString());
            }
        }
        */
        Dictionary<string, int> languages;

        public void scanLanguages()
        {
            int rowsCount = 1;
            languages = new Dictionary<string, int>();
            using (reader = new XmlTextReader(path))
            {
                while (reader.Read())
                {
                    if (reader.IsStartElement() 
                        && reader.LocalName == "tu")
                    {
                        rowsCount++;
                    }
                    else if (reader.IsStartElement() && reader.LocalName == "tuv")
                    {
                        string lang = reader.GetAttribute("xml:lang");
                        if (lang == null)
                        {
                            lang = reader.GetAttribute("lang");
                            if (lang == null)
                            {
                                throw new IOException("File doesn't have proper structure");
                            }
                        }
                        if (!languages.ContainsKey(lang))
                        {
                            languages.Add(lang, languages.Count);
                        }
                    }
                }
            }
            output.writeWorkSheetStart(rowsCount, languages.Count);
            foreach (KeyValuePair<string, int> item in languages)
            {
                output.addCell(item.Value, item.Key);
            }
            output.writeRow();
        }

        public void makeJob()
        {
            if (languages.Count == 0)
            {
                scanLanguages();
            }

            using (reader = new XmlTextReader(path)) {
                while (reader.Read())
                {
                    if (reader.IsStartElement() && reader.LocalName == "tu")
                    {
                        XmlDocument xmlDoc = new XmlDocument();
                        xmlDoc.LoadXml(reader.ReadOuterXml());

                        foreach (XmlNode item in xmlDoc.FirstChild.ChildNodes)
                        {
                            if (item.LocalName == "tuv")
                            {
                                string lang;
                                if (null == item.Attributes["xml:lang"])
                                {
                                    lang = item.Attributes["lang"].Value;
                                }
                                else
                                {
                                    lang = item.Attributes["xml:lang"].Value;
                                }
                                int columnNumber = languages[lang];

                                foreach (XmlNode seg in item.ChildNodes)
                                {
                                    if (seg.LocalName == "seg")
                                    {
                                        foreach (XmlNode txt in seg.ChildNodes)
                                        {
                                            if (txt.NodeType == XmlNodeType.Text)
                                            {
                                                output.addCell(columnNumber, txt.Value);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        output.writeRow();
                    }                   
                }
            }
            output.Close();
        }
    }
}
