using HtmlAgilityPack;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Word;
using QuoteAddIn.Properties;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;
using Office = Microsoft.Office.Core;

namespace QuoteAddIn
{
    [ComVisible(true)]
    public class QuoteRibbon : Office.IRibbonExtensibility
    {
        private IRibbonUI ribbon;

        public QuoteRibbon()
        {
        }

        #region Элементы IRibbonExtensibility

        public string GetCustomUI(string ribbonID)
        {
            if (ribbonID == "Microsoft.Outlook.Mail.Compose")
                return GetResourceText("QuoteAddIn.QuoteRibbon.xml");
            else 
                return "";
        }

        #endregion

        #region Обратные вызовы ленты

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Вспомогательные методы

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        public Bitmap GetImage(IRibbonControl control)
        {
            switch (control.Id)
            {
                case "btnQuote":
                    return Resources.Quote;
            }
            return null;
        }

        #endregion

        public void QuoteCommandClick(IRibbonControl e)
        {
           var app =  Globals.ThisAddIn.Application;
           if (app == null)
                return;
            
            var inspector = app.ActiveInspector();
            var explorer = app.ActiveExplorer();

            if (inspector == null)
               return;

            if (inspector.CurrentItem != null)
            {
                MailItem mailItem = inspector.CurrentItem;
                if (mailItem.BodyFormat == OlBodyFormat.olFormatHTML)
                {

                    if (inspector.EditorType == OlEditorType.olEditorHTML)
                    {
                        var editor = inspector.HTMLEditor;
                    }
                    else if (inspector.EditorType == OlEditorType.olEditorWord)
                    {
                        Microsoft.Office.Interop.Outlook.Selection oSel = explorer.Selection;

                        List<string> rawQuotes = new List<string>();

                        string author = "Кто-то";

                        for (int i = 1; i <= oSel.Count; i++)
                        {
                            var oItem = oSel[i];
                            MailItem oMail = (MailItem)oItem;
                            Inspector localInspector = oMail.GetInspector;
                            Microsoft.Office.Interop.Word.Document document =
                                (Microsoft.Office.Interop.Word.Document)inspector.WordEditor;
                            string quoteText = document.Application.Selection.Text;

                            List<ParserDescriptor> parserDescriptors = new List<ParserDescriptor>();
                            parserDescriptors.Add(new ParserDescriptor("From","Sent"));
                            parserDescriptors.Add(new ParserDescriptor("От","Отправлено:"));

                            foreach (var pd in parserDescriptors)
                            {

                                Object findText = pd.SearchKey;
                                Object missing = Type.Missing;

                                var range = document.Range(document.Content.Start, document.Application.Selection.Range.Start);
                                if (range.Find.Execute(
                                    ref findText, ref missing, ref missing, true, ref missing,
                                    ref missing, false, ref missing, ref missing, ref missing,
                                    ref missing, ref missing, ref missing, ref missing, ref missing))
                                {
                                    pd.Result = document.Range(range.Start, range.End);
                                }
                            }

                            var optimalParser = parserDescriptors.OrderByDescending(pd => pd.Result == null ? 0 : pd.Result.Start).FirstOrDefault();
                            if (optimalParser != null)
                            {
                                string parsedAuthor = optimalParser.ParseSender(document);
                                if (!String.IsNullOrWhiteSpace(parsedAuthor))
                                    author = parsedAuthor;
                            }

                            rawQuotes.Add(quoteText);
                        }

                        string styledQuote = "";

                        foreach (var rawQuote in rawQuotes)
                        {
                            styledQuote += Properties.Resources.QuoteTemplate
                                .Replace("[%QUOTE_ID%]", Guid.NewGuid().ToString().Replace("-", ""))
                                .Replace("[%QUOTE_AUTHOR%]", author)
                                .Replace("[%QUOTE_TEXT%]", rawQuote);
                        }

                        mailItem.HTMLBody = QuoteInsert(styledQuote, mailItem.HTMLBody);
                    }
                }
            }
        }

        /// <summary>
        /// Размещает сформированный блок-цитату вверху редактируемого письма.
        /// TODO: каждую следующую цитату стоит размещать внизу последней
        /// </summary>
        /// <param name="item"></param>
        /// <param name="body"></param>
        /// <returns></returns>
        private string QuoteInsert(string item, string body)
        {
            HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(body);
            var bodyNode = doc.DocumentNode.SelectSingleNode("//body");

            HtmlNode newNode = HtmlNode.CreateNode(item);
            bodyNode.PrependChild(HtmlNode.CreateNode("<p class=MsoNormal><o:p></o:p></p>"));
            bodyNode.PrependChild(newNode);

            return doc.DocumentNode.OuterHtml;
        }
    }

    internal class ParserDescriptor
    {
        private string fromKeyword = "";
        private string sentKeyword = "";
        
        public ParserDescriptor(string from, string sent)
        {
            this.fromKeyword = from;
            this.sentKeyword = sent;
        }

        public string SearchKey => fromKeyword + ":*" + sentKeyword + ":";
        public int LeftPad => fromKeyword.Length + 1;
        public int RightPad => sentKeyword.Length + 1;

        public Range Result { get; set; }

        internal string ParseSender(Document document)
        {
            if (Result == null)
                return "";

            return document.Range(Result.Start + 5, Result.End - 5).Text.Trim();
        }
    }
}
