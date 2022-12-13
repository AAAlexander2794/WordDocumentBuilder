using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDocumentBuilder
{
    /// <summary>
    /// Класс работы непосредственно с файлом .docx.
    /// </summary>
    /// <remarks>
    /// Не должен ничего знать про бизнес-логику.
    /// </remarks>
    public class WordDocument
    {
        private WordprocessingDocument Document { get; }

        public IDictionary<string, BookmarkStart> BookmarkMap { get; }

        public WordDocument(string templatePath)
        {
            Document = WordprocessingDocument.CreateFromTemplate(templatePath);
            BookmarkMap = GetBookmarkMap(Document);
        }

        public void Save(string path)
        {
            Document.SaveAs(path);
        }

        public void Close()
        {
            Document.Close();
        }

        /// <summary>
        /// Возвращает карту закладок файла .docx
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns>Словарь закладок файла .docx</returns>
        private static IDictionary<string, BookmarkStart> GetBookmarkMap(string filePath)
        {
            //
            IDictionary<string, BookmarkStart> bookmarkMap = new Dictionary<string, BookmarkStart>();
            //
            using (WordprocessingDocument document = WordprocessingDocument.CreateFromTemplate(filePath))
            {
                bookmarkMap = GetBookmarkMap(document);
            }
            //
            return bookmarkMap;
        }

        /// <summary>
        /// Возвращает карту закладок шаблона
        /// </summary>
        /// <param name="document"></param>
        /// <returns>Словарь закладок шаблона</returns>
        private static IDictionary<string, BookmarkStart> GetBookmarkMap(WordprocessingDocument document)
        {
            //
            IDictionary<string, BookmarkStart> bookmarkMap = new Dictionary<string, BookmarkStart>();
            //
            foreach (BookmarkStart bookmarkStart in document.MainDocumentPart.RootElement.Descendants<BookmarkStart>())
            {
                bookmarkMap[bookmarkStart.Name] = bookmarkStart;
            }
            //
            return bookmarkMap;
        }

        public void SetBookmarkText(string bookmarkName, string text)
        {
            try
            {
                SetBookmarkText(BookmarkMap[bookmarkName], text);
            }
            catch (Exception ex)
            {
                throw new Exception("Закладка не найдена.\n" + ex.Message);
            };
        }

        private static void SetBookmarkText(BookmarkStart bookmarkStart, string text)
        {
            Run bookmarkText = bookmarkStart.NextSibling<Run>();
            if (bookmarkText != null)
            {
                bookmarkText.GetFirstChild<Text>().Text = text;
            }
        }

        /// <summary>
        /// Вставляет таблицу после закладки
        /// </summary>
        /// <param name="bookmarkName"></param>
        /// <param name="table"></param>
        public void SetBookmarkTable(string bookmarkName, Table table)
        {
            if (table == null) return;
            var mainPart = Document.MainDocumentPart;
            var res = from bm in mainPart.Document.Body.Descendants<BookmarkStart>()
                      where bm.Name == bookmarkName
                      select bm;
            var bookmark = res.SingleOrDefault();
            if (bookmark != null)
            {
                // Родитель закладки
                var parent = bookmark.Parent;
                // 
                parent.InsertAfterSelf(table);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="mergeFieldName">Название поля для слияния.</param>
        /// <param name="text"></param>
        public void SetMergeFieldText(string mergeFieldName, string text)
        {
            //
            string FieldDelimeter = " MERGEFIELD ";
            string FieldDelimeterEnd = " \\* MERGEFORMAT ";

            foreach (FieldCode field in Document.MainDocumentPart.RootElement.Descendants<FieldCode>())
            {
                var fieldNameStart = field.Text.LastIndexOf(FieldDelimeter, System.StringComparison.Ordinal);
                var fieldName = field.Text.Substring(fieldNameStart + FieldDelimeter.Length).Replace(FieldDelimeterEnd, "").Trim();

                if (fieldName == mergeFieldName)
                {
                    foreach (Run run in Document.MainDocumentPart.Document.Descendants<Run>())
                    {
                        foreach (Text txtFromRun in run.Descendants<Text>().Where(a => a.Text == $"«{fieldName}»"))
                        {
                            txtFromRun.Text = text;
                        }
                    }
                }
            }
        }

    }
}
