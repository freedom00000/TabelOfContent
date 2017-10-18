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
  
            using (WordprocessingDocument docWord = WordprocessingDocument.Open(filepath,true)) {
             //   var doc = docWord.MainDocumentPart.Document;
                XElement firstPara = 
                     GetXDocument(docWord.MainDocumentPart)
                    .Descendants(w + "p")
                    .FirstOrDefault();
                AddToc(docWord, firstPara,
@"TOC \o '1-3' \h \z \u", "Оглавление", null);
                //OpenXmlElement block = doc.Descendants<DocPartGallery>().
                //          Where(b => b.Val.HasValue &&
                //            (b.Val.Value == "Table of Contents")).FirstOrDefault();

            } 
        }


        public static void AddToc(WordprocessingDocument doc, XElement addBefore, string switches, string title, int? rightTabPos)
        {
            // UpdateFontTablePart(doc);
            // UpdateStylesPartForToc(doc);
            //UpdateStylesWithEffectsPartForToc(doc);

            if (title == null)
                title = "Оглавsdaweление";
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

            //XDocument settingsXDoc = doc.MainDocumentPart.DocumentSettingsPart.GetXDocument();
            //XElement updateFields = settingsXDoc.Descendants(W.updateFields).FirstOrDefault();
            //if (updateFields != null)
            //    updateFields.Attribute(W.val).Value = "true";
            //else
            //{
            //    updateFields = new XElement(W.updateFields,
            //        new XAttribute(W.val, "true"));
            //    settingsXDoc.Root.Add(updateFields);
            //}
            //doc.MainDocumentPart.DocumentSettingsPart.PutXDocument();
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

        //void UpdateFontTablePart(WordprocessingDocument doc)
        //{
        //    FontTablePart fontTablePart = doc.MainDocumentPart.FontTablePart;

        //    XDocument fontTableXDoc = GetXDocument(fontTablePart);

        //    AddElementIfMissing(fontTableXDoc,
        //        fontTableXDoc
        //            .Root
        //            .Elements(w + "font")
        //            .Where(e => (string)e.Attribute(w + "name") == "Tahoma")
        //            .FirstOrDefault(),
        //        @"<w:font w:name='Tahoma' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
        //             <w:panose1 w:val='020B0604030504040204'/>
        //             <w:charset w:val='00'/>
        //             <w:family w:val='swiss'/>
        //             <w:pitch w:val='variable'/>
        //             <w:sig w:usb0='E1002EFF' w:usb1='C000605B' w:usb2='00000029' w:usb3='00000000' w:csb0='000101FF' w:csb1='00000000'/>
        //           </w:font>");

        //    PutXDocument(fontTablePart);
        //}

        //private  void AddElementIfMissing(XDocument partXDoc, XElement existing, string newElement)
        //{
        //    if (existing != null)
        //        return;
        //    XElement newXElement = XElement.Parse(newElement);
        //    newXElement.Attributes().Where(a => a.IsNamespaceDeclaration).Remove();
        //    partXDoc.Root.Add(newXElement);
        //}

        //private  void UpdateStylesPartForToc(WordprocessingDocument doc)
        //{
        //    StylesPart stylesPart = doc.MainDocumentPart.StyleDefinitionsPart;

        //    XDocument stylesXDoc = GetXDocument(stylesPart);
        //    UpdateAStylePartForToc(stylesXDoc);
        //    PutXDocument(stylesPart);
        //}

        //private  void UpdateStylesWithEffectsPartForToc(WordprocessingDocument doc)
        //{
        //    StylesWithEffectsPart stylesWithEffectsPart = doc.MainDocumentPart.StylesWithEffectsPart;

        //    XDocument stylesWithEffectsXDoc = GetXDocument(stylesWithEffectsPart);
        //    UpdateAStylePartForToc(stylesWithEffectsXDoc);
        //    PutXDocument(stylesWithEffectsPart);
        //}

        //private  void UpdateAStylePartForToc(XDocument partXDoc)
        //{
        //    AddElementIfMissing(
        //        partXDoc,
        //        partXDoc.Root.Elements(w + "style")
        //            .Where(e => (string)e.Attribute(w + "type") == "paragraph" && (string)e.Attribute(w + "styleId") == "TOCHeading")
        //            .FirstOrDefault(),
        //        @"<w:style w:type='paragraph' w:styleId='TOCHeading' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
        //            <w:name w:val='TOC Heading'/>
        //            <w:basedOn w:val='Heading1'/>
        //            <w:next w:val='Normal'/>
        //            <w:uiPriority w:val='39'/>
        //            <w:semiHidden/>
        //            <w:unhideWhenUsed/>
        //            <w:qFormat/>
        //            <w:pPr>
        //              <w:outlineLvl w:val='9'/>
        //            </w:pPr>
        //            <w:rPr>
        //              <w:lang w:eastAsia='ja-JP'/>
        //            </w:rPr>
        //          </w:style>");

        //    for (int i = 1; i <= 6; ++i)
        //    {
        //        AddElementIfMissing(
        //            partXDoc,
        //            partXDoc.Root.Elements(w + "style")
        //                .Where(e => (string)e.Attribute(w + "type") == "paragraph" && (string)e.Attribute(w + "styleId") == ("TOC" + i.ToString()))
        //                .FirstOrDefault(),
        //            String.Format(
        //                @"<w:style w:type='paragraph' w:styleId='TOC{0}' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
        //                    <w:name w:val='toc {0}'/>
        //                    <w:basedOn w:val='Normal'/>
        //                    <w:next w:val='Normal'/>
        //                    <w:autoRedefine/>
        //                    <w:uiPriority w:val='39'/>
        //                    <w:unhideWhenUsed/>
        //                    <w:pPr>
        //                      <w:spacing w:after='100'/>
        //                    </w:pPr>
        //                  </w:style>", i));
        //    }

        //    AddElementIfMissing(
        //        partXDoc,
        //        partXDoc.Root.Elements(w + "style")
        //            .Where(e => (string)e.Attribute(w + "type") == "character" && (string)e.Attribute(w + "styleId") == "Hyperlink")
        //            .FirstOrDefault(),
        //        @"<w:style w:type='character' w:styleId='Hyperlink' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
        //             <w:name w:val='Hyperlink'/>
        //             <w:basedOn w:val='DefaultParagraphFont'/>
        //             <w:uiPriority w:val='99'/>
        //             <w:unhideWhenUsed/>
        //             <w:rPr>
        //               <w:color w:val='0000FF' w:themeColor='hyperlink'/>
        //               <w:u w:val='single'/>
        //             </w:rPr>
        //           </w:style>");

        //    AddElementIfMissing(
        //        partXDoc,
        //        partXDoc.Root.Elements(w + "style")
        //            .Where(e => (string)e.Attribute(w + "type") == "paragraph" && (string)e.Attribute(w + "styleId") == "BalloonText")
        //            .FirstOrDefault(),
        //        @"<w:style w:type='paragraph' w:styleId='BalloonText' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
        //            <w:name w:val='Balloon Text'/>
        //            <w:basedOn w:val='Normal'/>
        //            <w:link w:val='BalloonTextChar'/>
        //            <w:uiPriority w:val='99'/>
        //            <w:semiHidden/>
        //            <w:unhideWhenUsed/>
        //            <w:pPr>
        //              <w:spacing w:after='0' w:line='240' w:lineRule='auto'/>
        //            </w:pPr>
        //            <w:rPr>
        //              <w:rFonts w:ascii='Tahoma' w:hAnsi='Tahoma' w:cs='Tahoma'/>
        //              <w:sz w:val='16'/>
        //              <w:szCs w:val='16'/>
        //            </w:rPr>
        //          </w:style>");

        //    AddElementIfMissing(
        //        partXDoc,
        //        partXDoc.Root.Elements(w + "style")
        //            .Where(e => (string)e.Attribute(w + "type") == "character" &&
        //                (bool?)e.Attribute(w + "customStyle") == true && (string)e.Attribute(w + "styleId") == "BalloonTextChar")
        //            .FirstOrDefault(),
        //        @"<w:style w:type='character' w:customStyle='1' w:styleId='BalloonTextChar' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
        //            <w:name w:val='Balloon Text Char'/>
        //            <w:basedOn w:val='DefaultParagraphFont'/>
        //            <w:link w:val='BalloonText'/>
        //            <w:uiPriority w:val='99'/>
        //            <w:semiHidden/>
        //            <w:rPr>
        //              <w:rFonts w:ascii='Tahoma' w:hAnsi='Tahoma' w:cs='Tahoma'/>
        //              <w:sz w:val='16'/>
        //              <w:szCs w:val='16'/>
        //            </w:rPr>
        //          </w:style>");
        //}
    }

  

}
