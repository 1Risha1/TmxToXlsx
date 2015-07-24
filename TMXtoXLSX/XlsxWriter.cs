using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.IO;
using System.IO.Compression;

namespace TMXtoXLSX
{
    class XlsxWriter
    {
        // хранилище строк
        class SharedStrings
        {
            private int count = 0;
            private Dictionary<string, int> strings = new Dictionary<string,int>();

            public void add(string key)
            {
                count++;
                if (!strings.ContainsKey(key))
                {
                    strings.Add(key, strings.Count);
                }
            }

            public string[] keysArray
            {
                get { return strings.OrderBy(pair => pair.Value).ToDictionary(pair => pair.Key, pair => pair.Value).Keys.ToArray(); }
            }

            public int uniqueCount
            {
                get { return strings.Count; }
            }
            
            public int Count
            {
                get {return count; }
            }

            public int IndexOf(string str)
            {
                return strings[str];
            }
        }

        private XmlTextWriter workSheetWriter;
        private Dictionary<string, string> cells = new Dictionary<string,string>();

        public void addCell(int columnIndex, string value)
        {
            sharedStrings.add(value);
            cells.Add(getColumnName(columnIndex) + rowIndex, value);
        }

        internal static string getColumnName(int columnIndex)
        {
            string colName = ((char)(((int)'A') + columnIndex % 26)).ToString();

            columnIndex = columnIndex / 26;

            while (columnIndex != 0)
            {
                colName = ((char)(((int)'A') - 1 + columnIndex)).ToString() + colName;
                columnIndex = columnIndex / 26;
            }

            return colName.ToString();
        }

        private string getTemporaryDirectory()
        {
            string tempDirectory = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
            while (File.Exists(tempDirectory))
            {
                tempDirectory = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
            }
            Directory.CreateDirectory(tempDirectory);
            return tempDirectory;
        }

        static string tmpPath;
        private string outputFilePath;

        public XlsxWriter(string outputFilePath)
        {
            tmpPath = getTemporaryDirectory();
            this.outputFilePath = outputFilePath;
            Directory.CreateDirectory(tmpPath + "//xl//worksheets");
            Directory.CreateDirectory(tmpPath + "//docProps");
            Directory.CreateDirectory(tmpPath + "//_rels");
            Directory.CreateDirectory(tmpPath + "//xl//_rels");
            Directory.CreateDirectory(tmpPath + "//xl//theme");

        }
   /*     ~XlsxWriter() {
            workSheetWriter.Close();
            try
            {
                DirectoryInfo directoryinfo = new DirectoryInfo(tmpPath);
                directoryinfo.Delete(true);
            }
            catch { }
        }
        */
        int rowIndex = 1;
        SharedStrings sharedStrings = new SharedStrings();

        bool isFirst = true;
        int columnsCount;

        public void writeRow()
        {
            if (isFirst) {
                columnsCount = cells.Count();
                isFirst = false;
            }    

            workSheetWriter.WriteStartElement("row");
            workSheetWriter.WriteAttributeString("r", rowIndex.ToString());
            workSheetWriter.WriteAttributeString("spans", "1:" + columnsCount);
        
            foreach (KeyValuePair<string, string> item in cells)
            {                
                workSheetWriter.WriteStartElement("c");
                workSheetWriter.WriteAttributeString("r", item.Key);
                workSheetWriter.WriteAttributeString("t", "s");
                workSheetWriter.WriteElementString("v", sharedStrings.IndexOf(item.Value).ToString());
                workSheetWriter.WriteEndElement();
            }
            workSheetWriter.WriteEndElement();
            cells.Clear();
            rowIndex++;
        }


        private void writeSharedString()
        {
            using (XmlTextWriter sharedStringWriter = new XmlTextWriter(File.CreateText(tmpPath + "//xl//sharedStrings.xml")))
            {
                sharedStringWriter.WriteStartDocument(true);
                sharedStringWriter.WriteStartElement("sst");
                sharedStringWriter.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                sharedStringWriter.WriteAttributeString("count", sharedStrings.Count.ToString());
                sharedStringWriter.WriteAttributeString("uniqueCount", sharedStrings.uniqueCount.ToString());

                foreach (string str in sharedStrings.keysArray)
                {
                    sharedStringWriter.WriteStartElement("si");
                    sharedStringWriter.WriteElementString("t", str);
                    sharedStringWriter.WriteEndElement();
                }

                sharedStringWriter.WriteEndElement();
            }
        }

        public void writeWorkSheetStart(int rowsCount, int columnsCount)
        {
            workSheetWriter = new XmlTextWriter(File.CreateText(tmpPath + "//xl//worksheets//sheet1.xml"));
            workSheetWriter.WriteStartDocument(true);

            workSheetWriter.WriteStartElement("worksheet");
            workSheetWriter.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            workSheetWriter.WriteAttributeString("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            workSheetWriter.WriteStartElement("dimension");
            workSheetWriter.WriteAttributeString("ref", "A1:" + getColumnName(columnsCount) + rowsCount);
            workSheetWriter.WriteEndElement();

            workSheetWriter.WriteStartElement("sheetViews");
            workSheetWriter.WriteStartElement("sheetView");
            workSheetWriter.WriteAttributeString("tabSelected", "1");
            workSheetWriter.WriteAttributeString("workbookViewId", "0");
            workSheetWriter.WriteEndElement();
            workSheetWriter.WriteEndElement();

            workSheetWriter.WriteStartElement("sheetFormatPr");
            workSheetWriter.WriteAttributeString("defaultRowHeight", "15");
            workSheetWriter.WriteEndElement();

            workSheetWriter.WriteStartElement("sheetData");
    }
        bool isClosed = false;

        public void Close()
        {
            if (isClosed)
                return;

            writeEndOfWorksheet();
            writeSharedString();
        
            //файлы, необходимы для формата xlsx
            writeContentXml();
            writeWorkBook();
            writeRelations();
            writeApp();
            writeCore();
            writeStyleSheet();
            writeWorkBookRelations();
            writeTheme();
            System.IO.Compression.ZipFile.CreateFromDirectory(tmpPath, outputFilePath);

            isClosed = true;
        }

        public void writeEndOfWorksheet()
        {
            workSheetWriter.WriteEndElement();

            workSheetWriter.WriteStartElement("pageMargins");
            workSheetWriter.WriteAttributeString("left", "0.7");
            workSheetWriter.WriteAttributeString("right", "0.7");
            workSheetWriter.WriteAttributeString("top", "0.75");
            workSheetWriter.WriteAttributeString("bottom", "0.75");
            workSheetWriter.WriteAttributeString("header", "0.3");
            workSheetWriter.WriteAttributeString("footer", "0.3");
            workSheetWriter.WriteEndElement();

            workSheetWriter.WriteEndElement();

            workSheetWriter.Close();
        }

        void writeApp()
        {
            using (StreamWriter writer = new StreamWriter(tmpPath + "//docProps//app.xml"))
            {
                writer.Write(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes"" ?> 
<Properties xmlns=""http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"" xmlns:vt=""http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"">
  <Application>Microsoft Excel</Application> 
  <DocSecurity>0</DocSecurity> 
  <ScaleCrop>false</ScaleCrop> 
 <HeadingPairs>
 <vt:vector size=""2"" baseType=""variant"">
 <vt:variant>
  <vt:lpstr>Worksheets</vt:lpstr> 
  </vt:variant>
 <vt:variant>
<vt:i4>1</vt:i4></vt:variant>
  </vt:vector>
  </HeadingPairs>
 <TitlesOfParts>
 <vt:vector size=""1"" baseType=""lpstr""><vt:lpstr>Лист1</vt:lpstr>
   </vt:vector>
  </TitlesOfParts>
  <LinksUpToDate>false</LinksUpToDate> 
  <SharedDoc>false</SharedDoc> 
  <HyperlinksChanged>false</HyperlinksChanged> 
  <AppVersion>14.0300</AppVersion> 
  </Properties>");
            }
        }

        void writeCore()
        {
            using (StreamWriter writer = new StreamWriter(tmpPath + "//docProps//core.xml"))
            {
                writer.Write(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes"" ?> 
 <cp:coreProperties xmlns:cp=""http://schemas.openxmlformats.org/package/2006/metadata/core-properties"" xmlns:dc=""http://purl.org/dc/elements/1.1/"" xmlns:dcterms=""http://purl.org/dc/terms/"" xmlns:dcmitype=""http://purl.org/dc/dcmitype/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">
  <dc:creator>Administrator</dc:creator> 
  <cp:lastModifiedBy>Windows User</cp:lastModifiedBy> 
  <dcterms:created xsi:type=""dcterms:W3CDTF"">2011-06-16T08:14:40Z</dcterms:created> 
  <dcterms:modified xsi:type=""dcterms:W3CDTF"">2011-06-16T08:14:40Z</dcterms:modified> 
  </cp:coreProperties>");
            }
        }

        void writeStyleSheet()
        {
            using (StreamWriter writer = new StreamWriter(tmpPath + "//xl//styles.xml"))
            {
                writer.Write(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<styleSheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"" mc:Ignorable=""x14ac"" xmlns:x14ac=""http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac""><fonts count=""18"" x14ac:knownFonts=""1""><font><sz val=""11""/><color theme=""1""/><name val=""Calibri""/><family val=""2""/><scheme val=""minor""/></font><font><sz val=""11""/><color theme=""1""/><name val=""Calibri""/><family val=""2""/><scheme val=""minor""/></font><font><b/><sz val=""18""/><color theme=""3""/><name val=""Cambria""/><family val=""2""/><scheme val=""major""/></font><font><b/><sz val=""15""/><color theme=""3""/><name val=""Calibri""/><family val=""2""/><scheme val=""minor""/></font><font><b/><sz val=""13""/><color theme=""3""/><name val=""Calibri""/><family val=""2""/><scheme val=""minor""/></font><font><b/><sz val=""11""/><color theme=""3""/><name val=""Calibri""/><family val=""2""/><scheme val=""minor""/></font><font><sz val=""11""/><color rgb=""FF006100""/><name val=""Calibri""/><family val=""2""/><scheme val=""minor""/></font><font><sz val=""11""/><color rgb=""FF9C0006""/><name val=""Calibri""/><family val=""2""/><scheme val=""minor""/></font><font><sz val=""11""/><color rgb=""FF9C6500""/><name val=""Calibri""/><family val=""2""/><scheme val=""minor""/></font><font><sz val=""11""/><color rgb=""FF3F3F76""/><name val=""Calibri""/><family val=""2""/><scheme val=""minor""/></font><font><b/><sz val=""11""/><color rgb=""FF3F3F3F""/><name val=""Calibri""/><family val=""2""/><scheme val=""minor""/></font><font><b/><sz val=""11""/><color rgb=""FFFA7D00""/><name val=""Calibri""/><family val=""2""/><scheme val=""minor""/></font><font><sz val=""11""/><color rgb=""FFFA7D00""/><name val=""Calibri""/><family val=""2""/><scheme val=""minor""/></font><font><b/><sz val=""11""/><color theme=""0""/><name val=""Calibri""/><family val=""2""/><scheme val=""minor""/></font><font><sz val=""11""/><color rgb=""FFFF0000""/><name val=""Calibri""/><family val=""2""/><scheme val=""minor""/></font><font><i/><sz val=""11""/><color rgb=""FF7F7F7F""/><name val=""Calibri""/><family val=""2""/><scheme val=""minor""/></font><font><b/><sz val=""11""/><color theme=""1""/><name val=""Calibri""/><family val=""2""/><scheme val=""minor""/></font><font><sz val=""11""/><color theme=""0""/><name val=""Calibri""/><family val=""2""/><scheme val=""minor""/></font></fonts><fills count=""33""><fill><patternFill patternType=""none""/></fill><fill><patternFill patternType=""gray125""/></fill><fill><patternFill patternType=""solid""><fgColor rgb=""FFC6EFCE""/></patternFill></fill><fill><patternFill patternType=""solid""><fgColor rgb=""FFFFC7CE""/></patternFill></fill><fill><patternFill patternType=""solid""><fgColor rgb=""FFFFEB9C""/></patternFill></fill><fill><patternFill patternType=""solid""><fgColor rgb=""FFFFCC99""/></patternFill></fill><fill><patternFill patternType=""solid""><fgColor rgb=""FFF2F2F2""/></patternFill></fill><fill><patternFill patternType=""solid""><fgColor rgb=""FFA5A5A5""/></patternFill></fill><fill><patternFill patternType=""solid""><fgColor rgb=""FFFFFFCC""/></patternFill></fill><fill><patternFill patternType=""solid""><fgColor theme=""4""/></patternFill></fill><fill><patternFill patternType=""solid""><fgColor theme=""4"" tint=""0.79998168889431442""/><bgColor indexed=""65""/></patternFill></fill><fill><patternFill patternType=""solid""><fgColor theme=""4"" tint=""0.59999389629810485""/><bgColor indexed=""65""/></patternFill></fill><fill><patternFill patternType=""solid""><fgColor theme=""4"" tint=""0.39997558519241921""/><bgColor indexed=""65""/></patternFill></fill><fill><patternFill patternType=""solid""><fgColor theme=""5""/></patternFill></fill><fill><patternFill patternType=""solid""><fgColor theme=""5"" tint=""0.79998168889431442""/><bgColor indexed=""65""/></patternFill></fill><fill><patternFill patternType=""solid""><fgColor theme=""5"" tint=""0.59999389629810485""/><bgColor indexed=""65""/></patternFill></fill><fill><patternFill patternType=""solid""><fgColor theme=""5"" tint=""0.39997558519241921""/><bgColor indexed=""65""/></patternFill></fill><fill><patternFill patternType=""solid""><fgColor theme=""6""/></patternFill></fill><fill><patternFill patternType=""solid""><fgColor theme=""6"" tint=""0.79998168889431442""/><bgColor indexed=""65""/></patternFill></fill><fill><patternFill patternType=""solid""><fgColor theme=""6"" tint=""0.59999389629810485""/><bgColor indexed=""65""/></patternFill></fill><fill><patternFill patternType=""solid""><fgColor theme=""6"" tint=""0.39997558519241921""/><bgColor indexed=""65""/></patternFill></fill><fill><patternFill patternType=""solid""><fgColor theme=""7""/></patternFill></fill><fill><patternFill patternType=""solid""><fgColor theme=""7"" tint=""0.79998168889431442""/><bgColor indexed=""65""/></patternFill></fill><fill><patternFill patternType=""solid""><fgColor theme=""7"" tint=""0.59999389629810485""/><bgColor indexed=""65""/></patternFill></fill><fill><patternFill patternType=""solid""><fgColor theme=""7"" tint=""0.39997558519241921""/><bgColor indexed=""65""/></patternFill></fill><fill><patternFill patternType=""solid""><fgColor theme=""8""/></patternFill></fill><fill><patternFill patternType=""solid""><fgColor theme=""8"" tint=""0.79998168889431442""/><bgColor indexed=""65""/></patternFill></fill><fill><patternFill patternType=""solid""><fgColor theme=""8"" tint=""0.59999389629810485""/><bgColor indexed=""65""/></patternFill></fill><fill><patternFill patternType=""solid""><fgColor theme=""8"" tint=""0.39997558519241921""/><bgColor indexed=""65""/></patternFill></fill><fill><patternFill patternType=""solid""><fgColor theme=""9""/></patternFill></fill><fill><patternFill patternType=""solid""><fgColor theme=""9"" tint=""0.79998168889431442""/><bgColor indexed=""65""/></patternFill></fill><fill><patternFill patternType=""solid""><fgColor theme=""9"" tint=""0.59999389629810485""/><bgColor indexed=""65""/></patternFill></fill><fill><patternFill patternType=""solid""><fgColor theme=""9"" tint=""0.39997558519241921""/><bgColor indexed=""65""/></patternFill></fill></fills><borders count=""10""><border><left/><right/><top/><bottom/><diagonal/></border><border><left/><right/><top/><bottom style=""thick""><color theme=""4""/></bottom><diagonal/></border><border><left/><right/><top/><bottom style=""thick""><color theme=""4"" tint=""0.499984740745262""/></bottom><diagonal/></border><border><left/><right/><top/><bottom style=""medium""><color theme=""4"" tint=""0.39997558519241921""/></bottom><diagonal/></border><border><left style=""thin""><color rgb=""FF7F7F7F""/></left><right style=""thin""><color rgb=""FF7F7F7F""/></right><top style=""thin""><color rgb=""FF7F7F7F""/></top><bottom style=""thin""><color rgb=""FF7F7F7F""/></bottom><diagonal/></border><border><left style=""thin""><color rgb=""FF3F3F3F""/></left><right style=""thin""><color rgb=""FF3F3F3F""/></right><top style=""thin""><color rgb=""FF3F3F3F""/></top><bottom style=""thin""><color rgb=""FF3F3F3F""/></bottom><diagonal/></border><border><left/><right/><top/><bottom style=""double""><color rgb=""FFFF8001""/></bottom><diagonal/></border><border><left style=""double""><color rgb=""FF3F3F3F""/></left><right style=""double""><color rgb=""FF3F3F3F""/></right><top style=""double""><color rgb=""FF3F3F3F""/></top><bottom style=""double""><color rgb=""FF3F3F3F""/></bottom><diagonal/></border><border><left style=""thin""><color rgb=""FFB2B2B2""/></left><right style=""thin""><color rgb=""FFB2B2B2""/></right><top style=""thin""><color rgb=""FFB2B2B2""/></top><bottom style=""thin""><color rgb=""FFB2B2B2""/></bottom><diagonal/></border><border><left/><right/><top style=""thin""><color theme=""4""/></top><bottom style=""double""><color theme=""4""/></bottom><diagonal/></border></borders><cellStyleXfs count=""42""><xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0""/><xf numFmtId=""0"" fontId=""2"" fillId=""0"" borderId=""0"" applyNumberFormat=""0"" applyFill=""0"" applyBorder=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""3"" fillId=""0"" borderId=""1"" applyNumberFormat=""0"" applyFill=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""4"" fillId=""0"" borderId=""2"" applyNumberFormat=""0"" applyFill=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""5"" fillId=""0"" borderId=""3"" applyNumberFormat=""0"" applyFill=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""5"" fillId=""0"" borderId=""0"" applyNumberFormat=""0"" applyFill=""0"" applyBorder=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""6"" fillId=""2"" borderId=""0"" applyNumberFormat=""0"" applyBorder=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""7"" fillId=""3"" borderId=""0"" applyNumberFormat=""0"" applyBorder=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""8"" fillId=""4"" borderId=""0"" applyNumberFormat=""0"" applyBorder=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""9"" fillId=""5"" borderId=""4"" applyNumberFormat=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""10"" fillId=""6"" borderId=""5"" applyNumberFormat=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""11"" fillId=""6"" borderId=""4"" applyNumberFormat=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""12"" fillId=""0"" borderId=""6"" applyNumberFormat=""0"" applyFill=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""13"" fillId=""7"" borderId=""7"" applyNumberFormat=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""14"" fillId=""0"" borderId=""0"" applyNumberFormat=""0"" applyFill=""0"" applyBorder=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""1"" fillId=""8"" borderId=""8"" applyNumberFormat=""0"" applyFont=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""15"" fillId=""0"" borderId=""0"" applyNumberFormat=""0"" applyFill=""0"" applyBorder=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""16"" fillId=""0"" borderId=""9"" applyNumberFormat=""0"" applyFill=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""17"" fillId=""9"" borderId=""0"" applyNumberFormat=""0"" applyBorder=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""1"" fillId=""10"" borderId=""0"" applyNumberFormat=""0"" applyBorder=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""1"" fillId=""11"" borderId=""0"" applyNumberFormat=""0"" applyBorder=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""17"" fillId=""12"" borderId=""0"" applyNumberFormat=""0"" applyBorder=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""17"" fillId=""13"" borderId=""0"" applyNumberFormat=""0"" applyBorder=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""1"" fillId=""14"" borderId=""0"" applyNumberFormat=""0"" applyBorder=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""1"" fillId=""15"" borderId=""0"" applyNumberFormat=""0"" applyBorder=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""17"" fillId=""16"" borderId=""0"" applyNumberFormat=""0"" applyBorder=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""17"" fillId=""17"" borderId=""0"" applyNumberFormat=""0"" applyBorder=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""1"" fillId=""18"" borderId=""0"" applyNumberFormat=""0"" applyBorder=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""1"" fillId=""19"" borderId=""0"" applyNumberFormat=""0"" applyBorder=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""17"" fillId=""20"" borderId=""0"" applyNumberFormat=""0"" applyBorder=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""17"" fillId=""21"" borderId=""0"" applyNumberFormat=""0"" applyBorder=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""1"" fillId=""22"" borderId=""0"" applyNumberFormat=""0"" applyBorder=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""1"" fillId=""23"" borderId=""0"" applyNumberFormat=""0"" applyBorder=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""17"" fillId=""24"" borderId=""0"" applyNumberFormat=""0"" applyBorder=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""17"" fillId=""25"" borderId=""0"" applyNumberFormat=""0"" applyBorder=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""1"" fillId=""26"" borderId=""0"" applyNumberFormat=""0"" applyBorder=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""1"" fillId=""27"" borderId=""0"" applyNumberFormat=""0"" applyBorder=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""17"" fillId=""28"" borderId=""0"" applyNumberFormat=""0"" applyBorder=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""17"" fillId=""29"" borderId=""0"" applyNumberFormat=""0"" applyBorder=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""1"" fillId=""30"" borderId=""0"" applyNumberFormat=""0"" applyBorder=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""1"" fillId=""31"" borderId=""0"" applyNumberFormat=""0"" applyBorder=""0"" applyAlignment=""0"" applyProtection=""0""/><xf numFmtId=""0"" fontId=""17"" fillId=""32"" borderId=""0"" applyNumberFormat=""0"" applyBorder=""0"" applyAlignment=""0"" applyProtection=""0""/></cellStyleXfs><cellXfs count=""1""><xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0"" xfId=""0""/></cellXfs><cellStyles count=""42""><cellStyle name=""20% - Accent1"" xfId=""19"" builtinId=""30"" customBuiltin=""1""/><cellStyle name=""20% - Accent2"" xfId=""23"" builtinId=""34"" customBuiltin=""1""/><cellStyle name=""20% - Accent3"" xfId=""27"" builtinId=""38"" customBuiltin=""1""/><cellStyle name=""20% - Accent4"" xfId=""31"" builtinId=""42"" customBuiltin=""1""/><cellStyle name=""20% - Accent5"" xfId=""35"" builtinId=""46"" customBuiltin=""1""/><cellStyle name=""20% - Accent6"" xfId=""39"" builtinId=""50"" customBuiltin=""1""/><cellStyle name=""40% - Accent1"" xfId=""20"" builtinId=""31"" customBuiltin=""1""/><cellStyle name=""40% - Accent2"" xfId=""24"" builtinId=""35"" customBuiltin=""1""/><cellStyle name=""40% - Accent3"" xfId=""28"" builtinId=""39"" customBuiltin=""1""/><cellStyle name=""40% - Accent4"" xfId=""32"" builtinId=""43"" customBuiltin=""1""/><cellStyle name=""40% - Accent5"" xfId=""36"" builtinId=""47"" customBuiltin=""1""/><cellStyle name=""40% - Accent6"" xfId=""40"" builtinId=""51"" customBuiltin=""1""/><cellStyle name=""60% - Accent1"" xfId=""21"" builtinId=""32"" customBuiltin=""1""/><cellStyle name=""60% - Accent2"" xfId=""25"" builtinId=""36"" customBuiltin=""1""/><cellStyle name=""60% - Accent3"" xfId=""29"" builtinId=""40"" customBuiltin=""1""/><cellStyle name=""60% - Accent4"" xfId=""33"" builtinId=""44"" customBuiltin=""1""/><cellStyle name=""60% - Accent5"" xfId=""37"" builtinId=""48"" customBuiltin=""1""/><cellStyle name=""60% - Accent6"" xfId=""41"" builtinId=""52"" customBuiltin=""1""/><cellStyle name=""Accent1"" xfId=""18"" builtinId=""29"" customBuiltin=""1""/><cellStyle name=""Accent2"" xfId=""22"" builtinId=""33"" customBuiltin=""1""/><cellStyle name=""Accent3"" xfId=""26"" builtinId=""37"" customBuiltin=""1""/><cellStyle name=""Accent4"" xfId=""30"" builtinId=""41"" customBuiltin=""1""/><cellStyle name=""Accent5"" xfId=""34"" builtinId=""45"" customBuiltin=""1""/><cellStyle name=""Accent6"" xfId=""38"" builtinId=""49"" customBuiltin=""1""/><cellStyle name=""Bad"" xfId=""7"" builtinId=""27"" customBuiltin=""1""/><cellStyle name=""Calculation"" xfId=""11"" builtinId=""22"" customBuiltin=""1""/><cellStyle name=""Check Cell"" xfId=""13"" builtinId=""23"" customBuiltin=""1""/><cellStyle name=""Explanatory Text"" xfId=""16"" builtinId=""53"" customBuiltin=""1""/><cellStyle name=""Good"" xfId=""6"" builtinId=""26"" customBuiltin=""1""/><cellStyle name=""Heading 1"" xfId=""2"" builtinId=""16"" customBuiltin=""1""/><cellStyle name=""Heading 2"" xfId=""3"" builtinId=""17"" customBuiltin=""1""/><cellStyle name=""Heading 3"" xfId=""4"" builtinId=""18"" customBuiltin=""1""/><cellStyle name=""Heading 4"" xfId=""5"" builtinId=""19"" customBuiltin=""1""/><cellStyle name=""Input"" xfId=""9"" builtinId=""20"" customBuiltin=""1""/><cellStyle name=""Linked Cell"" xfId=""12"" builtinId=""24"" customBuiltin=""1""/><cellStyle name=""Neutral"" xfId=""8"" builtinId=""28"" customBuiltin=""1""/><cellStyle name=""Normal"" xfId=""0"" builtinId=""0""/><cellStyle name=""Note"" xfId=""15"" builtinId=""10"" customBuiltin=""1""/><cellStyle name=""Output"" xfId=""10"" builtinId=""21"" customBuiltin=""1""/><cellStyle name=""Title"" xfId=""1"" builtinId=""15"" customBuiltin=""1""/><cellStyle name=""Total"" xfId=""17"" builtinId=""25"" customBuiltin=""1""/><cellStyle name=""Warning Text"" xfId=""14"" builtinId=""11"" customBuiltin=""1""/></cellStyles><dxfs count=""0""/><tableStyles count=""0"" defaultTableStyle=""TableStyleMedium2"" defaultPivotStyle=""PivotStyleLight16""/><extLst><ext uri=""{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}"" xmlns:x14=""http://schemas.microsoft.com/office/spreadsheetml/2009/9/main""><x14:slicerStyles defaultSlicerStyle=""SlicerStyleLight1""/></ext></extLst></styleSheet>
");
            }
        }

        void writeRelations()
        {
            using (StreamWriter writer = new StreamWriter(tmpPath + "//_rels//.rels"))
            {
                writer.Write(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes"" ?> 
 <Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
  <Relationship Id=""rId3"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"" Target=""docProps/app.xml"" /> 
  <Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"" Target=""docProps/core.xml"" /> 
  <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""xl/workbook.xml"" /> 
  </Relationships>");
            }
        }

        void writeWorkBookRelations()
        {
            using (StreamWriter writer = new StreamWriter(tmpPath + "//xl//_rels//workbook.xml.rels"))
            {
                writer.Write(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes"" ?> 
 <Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">");
                writer.Write(@"<Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet1.xml"" /> ");
                 
                writer.Write(@"<Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"" Target=""theme/theme1.xml"" /> ");
                writer.Write(@"<Relationship Id=""rId3"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""styles.xml"" />");
                writer.Write(@"<Relationship Id=""rId4"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"" Target=""sharedStrings.xml"" />");
                writer.Write("</Relationships>");
            }
        }

        void writeTheme()
        {
            using (StreamWriter writer = new StreamWriter(tmpPath + "//xl//theme//theme1.xml"))
            {
                writer.Write(@" <?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<a:theme xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" name=""Office Theme""><a:themeElements><a:clrScheme name=""Office""><a:dk1><a:sysClr val=""windowText"" lastClr=""000000""/></a:dk1><a:lt1><a:sysClr val=""window"" lastClr=""FFFFFF""/></a:lt1><a:dk2><a:srgbClr val=""1F497D""/></a:dk2><a:lt2><a:srgbClr val=""EEECE1""/></a:lt2><a:accent1><a:srgbClr val=""4F81BD""/></a:accent1><a:accent2><a:srgbClr val=""C0504D""/></a:accent2><a:accent3><a:srgbClr val=""9BBB59""/></a:accent3><a:accent4><a:srgbClr val=""8064A2""/></a:accent4><a:accent5><a:srgbClr val=""4BACC6""/></a:accent5><a:accent6><a:srgbClr val=""F79646""/></a:accent6><a:hlink><a:srgbClr val=""0000FF""/></a:hlink><a:folHlink><a:srgbClr val=""800080""/></a:folHlink></a:clrScheme><a:fontScheme name=""Office""><a:majorFont><a:latin typeface=""Cambria""/><a:ea typeface=""""/><a:cs typeface=""""/><a:font script=""Jpan"" typeface=""ＭＳ Ｐゴシック""/><a:font script=""Hang"" typeface=""맑은 고딕""/><a:font script=""Hans"" typeface=""宋体""/><a:font script=""Hant"" typeface=""新細明體""/><a:font script=""Arab"" typeface=""Times New Roman""/><a:font script=""Hebr"" typeface=""Times New Roman""/><a:font script=""Thai"" typeface=""Tahoma""/><a:font script=""Ethi"" typeface=""Nyala""/><a:font script=""Beng"" typeface=""Vrinda""/><a:font script=""Gujr"" typeface=""Shruti""/><a:font script=""Khmr"" typeface=""MoolBoran""/><a:font script=""Knda"" typeface=""Tunga""/><a:font script=""Guru"" typeface=""Raavi""/><a:font script=""Cans"" typeface=""Euphemia""/><a:font script=""Cher"" typeface=""Plantagenet Cherokee""/><a:font script=""Yiii"" typeface=""Microsoft Yi Baiti""/><a:font script=""Tibt"" typeface=""Microsoft Himalaya""/><a:font script=""Thaa"" typeface=""MV Boli""/><a:font script=""Deva"" typeface=""Mangal""/><a:font script=""Telu"" typeface=""Gautami""/><a:font script=""Taml"" typeface=""Latha""/><a:font script=""Syrc"" typeface=""Estrangelo Edessa""/><a:font script=""Orya"" typeface=""Kalinga""/><a:font script=""Mlym"" typeface=""Kartika""/><a:font script=""Laoo"" typeface=""DokChampa""/><a:font script=""Sinh"" typeface=""Iskoola Pota""/><a:font script=""Mong"" typeface=""Mongolian Baiti""/><a:font script=""Viet"" typeface=""Times New Roman""/><a:font script=""Uigh"" typeface=""Microsoft Uighur""/><a:font script=""Geor"" typeface=""Sylfaen""/></a:majorFont><a:minorFont><a:latin typeface=""Calibri""/><a:ea typeface=""""/><a:cs typeface=""""/><a:font script=""Jpan"" typeface=""ＭＳ Ｐゴシック""/><a:font script=""Hang"" typeface=""맑은 고딕""/><a:font script=""Hans"" typeface=""宋体""/><a:font script=""Hant"" typeface=""新細明體""/><a:font script=""Arab"" typeface=""Arial""/><a:font script=""Hebr"" typeface=""Arial""/><a:font script=""Thai"" typeface=""Tahoma""/><a:font script=""Ethi"" typeface=""Nyala""/><a:font script=""Beng"" typeface=""Vrinda""/><a:font script=""Gujr"" typeface=""Shruti""/><a:font script=""Khmr"" typeface=""DaunPenh""/><a:font script=""Knda"" typeface=""Tunga""/><a:font script=""Guru"" typeface=""Raavi""/><a:font script=""Cans"" typeface=""Euphemia""/><a:font script=""Cher"" typeface=""Plantagenet Cherokee""/><a:font script=""Yiii"" typeface=""Microsoft Yi Baiti""/><a:font script=""Tibt"" typeface=""Microsoft Himalaya""/><a:font script=""Thaa"" typeface=""MV Boli""/><a:font script=""Deva"" typeface=""Mangal""/><a:font script=""Telu"" typeface=""Gautami""/><a:font script=""Taml"" typeface=""Latha""/><a:font script=""Syrc"" typeface=""Estrangelo Edessa""/><a:font script=""Orya"" typeface=""Kalinga""/><a:font script=""Mlym"" typeface=""Kartika""/><a:font script=""Laoo"" typeface=""DokChampa""/><a:font script=""Sinh"" typeface=""Iskoola Pota""/><a:font script=""Mong"" typeface=""Mongolian Baiti""/><a:font script=""Viet"" typeface=""Arial""/><a:font script=""Uigh"" typeface=""Microsoft Uighur""/><a:font script=""Geor"" typeface=""Sylfaen""/></a:minorFont></a:fontScheme><a:fmtScheme name=""Office""><a:fillStyleLst><a:solidFill><a:schemeClr val=""phClr""/></a:solidFill><a:gradFill rotWithShape=""1""><a:gsLst><a:gs pos=""0""><a:schemeClr val=""phClr""><a:tint val=""50000""/><a:satMod val=""300000""/></a:schemeClr></a:gs><a:gs pos=""35000""><a:schemeClr val=""phClr""><a:tint val=""37000""/><a:satMod val=""300000""/></a:schemeClr></a:gs><a:gs pos=""100000""><a:schemeClr val=""phClr""><a:tint val=""15000""/><a:satMod val=""350000""/></a:schemeClr></a:gs></a:gsLst><a:lin ang=""16200000"" scaled=""1""/></a:gradFill><a:gradFill rotWithShape=""1""><a:gsLst><a:gs pos=""0""><a:schemeClr val=""phClr""><a:shade val=""51000""/><a:satMod val=""130000""/></a:schemeClr></a:gs><a:gs pos=""80000""><a:schemeClr val=""phClr""><a:shade val=""93000""/><a:satMod val=""130000""/></a:schemeClr></a:gs><a:gs pos=""100000""><a:schemeClr val=""phClr""><a:shade val=""94000""/><a:satMod val=""135000""/></a:schemeClr></a:gs></a:gsLst><a:lin ang=""16200000"" scaled=""0""/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w=""9525"" cap=""flat"" cmpd=""sng"" algn=""ctr""><a:solidFill><a:schemeClr val=""phClr""><a:shade val=""95000""/><a:satMod val=""105000""/></a:schemeClr></a:solidFill><a:prstDash val=""solid""/></a:ln><a:ln w=""25400"" cap=""flat"" cmpd=""sng"" algn=""ctr""><a:solidFill><a:schemeClr val=""phClr""/></a:solidFill><a:prstDash val=""solid""/></a:ln><a:ln w=""38100"" cap=""flat"" cmpd=""sng"" algn=""ctr""><a:solidFill><a:schemeClr val=""phClr""/></a:solidFill><a:prstDash val=""solid""/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad=""40000"" dist=""20000"" dir=""5400000"" rotWithShape=""0""><a:srgbClr val=""000000""><a:alpha val=""38000""/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad=""40000"" dist=""23000"" dir=""5400000"" rotWithShape=""0""><a:srgbClr val=""000000""><a:alpha val=""35000""/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad=""40000"" dist=""23000"" dir=""5400000"" rotWithShape=""0""><a:srgbClr val=""000000""><a:alpha val=""35000""/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst=""orthographicFront""><a:rot lat=""0"" lon=""0"" rev=""0""/></a:camera><a:lightRig rig=""threePt"" dir=""t""><a:rot lat=""0"" lon=""0"" rev=""1200000""/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w=""63500"" h=""25400""/></a:sp3d></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val=""phClr""/></a:solidFill><a:gradFill rotWithShape=""1""><a:gsLst><a:gs pos=""0""><a:schemeClr val=""phClr""><a:tint val=""40000""/><a:satMod val=""350000""/></a:schemeClr></a:gs><a:gs pos=""40000""><a:schemeClr val=""phClr""><a:tint val=""45000""/><a:shade val=""99000""/><a:satMod val=""350000""/></a:schemeClr></a:gs><a:gs pos=""100000""><a:schemeClr val=""phClr""><a:shade val=""20000""/><a:satMod val=""255000""/></a:schemeClr></a:gs></a:gsLst><a:path path=""circle""><a:fillToRect l=""50000"" t=""-80000"" r=""50000"" b=""180000""/></a:path></a:gradFill><a:gradFill rotWithShape=""1""><a:gsLst><a:gs pos=""0""><a:schemeClr val=""phClr""><a:tint val=""80000""/><a:satMod val=""300000""/></a:schemeClr></a:gs><a:gs pos=""100000""><a:schemeClr val=""phClr""><a:shade val=""30000""/><a:satMod val=""200000""/></a:schemeClr></a:gs></a:gsLst><a:path path=""circle""><a:fillToRect l=""50000"" t=""50000"" r=""50000"" b=""50000""/></a:path></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>");
            }
        }

        void writeWorkBook()
        {
            using (StreamWriter writer = new StreamWriter(tmpPath + "//xl//workbook.xml"))
            {
                writer.Write(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
<fileVersion appName=""xl"" lastEdited=""5"" lowestEdited=""5"" rupBuild=""9303""/>
<workbookPr defaultThemeVersion=""124226""/><bookViews><workbookView xWindow=""360"" yWindow=""105"" windowWidth=""14355""
windowHeight=""4695""/></bookViews><sheets><sheet name=""Лист1"" sheetId=""1"" r:id=""rId1""/>
</sheets><calcPr calcId=""145621""/></workbook>");
            }
        }
        void writeContentXml()
        {
            using (StreamWriter writer = new StreamWriter(tmpPath + "//[Content_Types].xml"))
            {
                writer.Write(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes"" ?> 
  <Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
  <Override PartName=""/xl/theme/theme1.xml"" ContentType=""application/vnd.openxmlformats-officedocument.theme+xml"" /> 
  <Override PartName=""/xl/styles.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"" /> 
  <Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml"" /> 
  <Default Extension=""xml"" ContentType=""application/xml"" /> 
  <Override PartName=""/xl/workbook.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"" /> 
  <Override PartName=""/docProps/app.xml"" ContentType=""application/vnd.openxmlformats-officedocument.extended-properties+xml"" /> 
<Override PartName=""/xl/worksheets/sheet1.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"" />
<Override PartName=""/xl/sharedStrings.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"" /> 
  <Override PartName=""/docProps/core.xml"" ContentType=""application/vnd.openxmlformats-package.core-properties+xml"" /> 
  </Types>");

            }
        }
    
    }
}
