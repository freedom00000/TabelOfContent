using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Xml.Linq;
using System.IO;
using System.Xml;

namespace TableOfContent
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        string filepath = "test.docx";
        static XNamespace w = @"http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        private void button_Click(object sender, RoutedEventArgs e)
        {

            DirectoryInfo di = new DirectoryInfo(".");
            foreach (var file in di.GetFiles("*.docx"))
                file.Delete();
            DirectoryInfo di2 = new DirectoryInfo("../../../");
            foreach (var file in di2.GetFiles("*.docx"))
                file.CopyTo(di.FullName + "/" + file.Name);


            WordprocessingDocument d = WordprocessingDocument.Open(filepath, true);
            var doc = d.MainDocumentPart.Document;
            OpenXmlElement block = doc.Descendants<DocPartGallery>().
  Where(b => b.Val.HasValue &&
    (b.Val.Value == "Table of Contents")).FirstOrDefault();
            if (block == null) {
                d.Close();
                using (WordprocessingDocument docWord = WordprocessingDocument.Open(filepath, true)) {

                    XElement firstPara =
                         GetXDocument(docWord.MainDocumentPart)
                        .Descendants(w + "p")
                        .FirstOrDefault();

                    if (firstPara.Parent.Name.LocalName != "sdtContent")
                        AddToc(docWord, firstPara,
        @"TOC \o '1-3' \h \z \u", "Оглавление", null);


                }
            } 
        }


        public static void AddToc(WordprocessingDocument doc, XElement addBefore, string switches, string title, int? rightTabPos)
        {

            if (title == null)
                title = "Contents";
            if (rightTabPos == null)
                rightTabPos = 9350;

            // {0} tocTitle (default = "Contents")
            // {1} rightTabPosition (default = 9350)
            // {2} switches

            String xmlString =
@"<w:sdt xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
  <w:sdtPr>
    <w:docPartObj>
      <w:docPartGallery w:val='Table of Contents'/>
      <w:docPartUnique/>
    </w:docPartObj>
  </w:sdtPr>
  <w:sdtEndPr>
    <w:rPr>
     <w:rFonts w:asciiTheme='minorHAnsi' w:cstheme='minorBidi' w:eastAsiaTheme='minorHAnsi' w:hAnsiTheme='minorHAnsi'/>
     <w:color w:val='auto'/>
     <w:sz w:val='22'/>
     <w:szCs w:val='22'/>
     <w:lang w:eastAsia='en-US'/>
    </w:rPr>
  </w:sdtEndPr>
  <w:sdtContent>
    <w:p>
      <w:pPr>
        <w:pStyle w:val='TOCHeading'/>
      </w:pPr>
      <w:r>
        <w:t>{0}</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val='TOC1'/>
        <w:tabs>
          <w:tab w:val='right' w:leader='dot' w:pos='{1}'/>
        </w:tabs>
        <w:rPr>
          <w:noProof/>
        </w:rPr>
      </w:pPr>
      <w:r>
        <w:fldChar w:fldCharType='begin' w:dirty='true'/>
      </w:r>
      <w:r>
        <w:instrText xml:space='preserve'> {2} </w:instrText>
      </w:r>
      <w:r>
        <w:fldChar w:fldCharType='separate'/>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:rPr>
          <w:b/>
          <w:bCs/>
          <w:noProof/>
        </w:rPr>
        <w:fldChar w:fldCharType='end'/>
      </w:r>
    </w:p>
  </w:sdtContent>
</w:sdt>";

            XElement sdt = XElement.Parse(String.Format(xmlString, title, rightTabPos, switches));
            addBefore.AddBeforeSelf(sdt);
            PutXDocument(doc.MainDocumentPart);

        }

        public static XDocument GetXDocument(OpenXmlPart part)
        {

            XDocument partXDocument = part.Annotation<XDocument>();
            if (partXDocument != null)
                return partXDocument;
            using (Stream partStream = part.GetStream())
            using (XmlReader partXmlReader = XmlReader.Create(partStream))
                partXDocument = XDocument.Load(partXmlReader);
            part.AddAnnotation(partXDocument);
            return partXDocument;
        }

        public static void PutXDocument(OpenXmlPart part)
        {
            XDocument partXDocument = GetXDocument(part);
            if (partXDocument != null)
            {
                using (Stream partStream = part.GetStream(FileMode.Create, FileAccess.Write))
                using (XmlWriter partXmlWriter = XmlWriter.Create(partStream))
                    partXDocument.Save(partXmlWriter);
            }
        }

    }

  

}
