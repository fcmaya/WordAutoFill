using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MinutaDocx.Core
{
    public static class DocumentFieldFiller
    {
        public static void PreencherDocumento(string documentPath, dynamic dados)
        {
            using (WordprocessingDocument document = WordprocessingDocument.Open(documentPath, true))
            {
                MainDocumentPart? mainPart = document.MainDocumentPart;
                if (mainPart?.Document?.Body == null)
                {
                    throw new InvalidOperationException("Documento inválido ou corrompido.");
                }

                // Method 1: Fill Content Controls by Title/Tag
                PreencherContentControls(mainPart, dados);

                // Method 2: Fill Form Fields (Legacy Form Fields)
                PreencherFormFields(mainPart, dados);

                // Method 3: Replace text placeholders
                SubstituirPlaceholders(mainPart, dados);

                // Method 4: Fill Bookmarks
                PreencherBookmarks(mainPart, dados);

                // Method 5: Fill Merge Fields
                PreencherMergeFields(mainPart, dados);

                // Save changes
                mainPart.Document.Save();
            }
        }

        private static void PreencherContentControls(MainDocumentPart mainPart, dynamic dados)
        {
            try
            {
                var contentControls = new List<OpenXmlElement>();
                contentControls.AddRange(mainPart.Document.Body.Descendants<SdtBlock>());
                contentControls.AddRange(mainPart.Document.Body.Descendants<SdtRun>());
                contentControls.AddRange(mainPart.Document.Body.Descendants<SdtCell>());

                Console.WriteLine($"Processando {contentControls.Count} Content Controls...");

                foreach (var contentControl in contentControls)
                {
                    var sdtProperties = contentControl.Descendants<SdtProperties>().FirstOrDefault();
                    if (sdtProperties == null) continue;

                    var tag = sdtProperties.Descendants<Tag>().FirstOrDefault()?.Val?.Value;
                    var alias = sdtProperties.Descendants<SdtAlias>().FirstOrDefault()?.Val?.Value;

                    string fieldName = tag ?? alias ?? "";

                    if (string.IsNullOrEmpty(fieldName)) continue;

                    var sdtContent = contentControl.Descendants<SdtContentRun>().FirstOrDefault() ??
                                   contentControl.Descendants<SdtContentBlock>().FirstOrDefault() ??
                                   (OpenXmlElement?)contentControl.Descendants<SdtContentCell>().FirstOrDefault();

                    if (sdtContent == null) continue;

                    string newValue = FieldValueMapper.GetValueForField(fieldName, dados);

                    if (!string.IsNullOrEmpty(newValue))
                    {
                        Console.WriteLine($"  Preenchendo Content Control '{fieldName}' com '{newValue}'");

                        // Clear existing content
                        sdtContent.RemoveAllChildren<Paragraph>();
                        sdtContent.RemoveAllChildren<Run>();
                        sdtContent.RemoveAllChildren<Text>();

                        // Add new content
                        if (sdtContent is SdtContentRun)
                        {
                            sdtContent.AppendChild(new Run(new Text(newValue)));
                        }
                        else
                        {
                            sdtContent.AppendChild(new Paragraph(new Run(new Text(newValue))));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao processar Content Controls: {ex.Message}");
            }
        }

        private static void PreencherFormFields(MainDocumentPart mainPart, dynamic dados)
        {
            try
            {
                var formFields = mainPart.Document.Body.Descendants<FormFieldData>().ToList();
                Console.WriteLine($"Processando {formFields.Count} Form Fields...");

                foreach (var formField in formFields)
                {
                    var fieldName = formField.XName?.LocalName;
                    if (string.IsNullOrEmpty(fieldName)) continue;

                    string newValue = FieldValueMapper.GetValueForField(fieldName, dados);

                    if (!string.IsNullOrEmpty(newValue))
                    {
                        Console.WriteLine($"  Preenchendo Form Field '{fieldName}' com '{newValue}'");

                        var fieldChar = formField.Parent as FieldChar;
                        if (fieldChar != null)
                        {
                            var nextElement = fieldChar.NextSibling();
                            while (nextElement != null)
                            {
                                if (nextElement is Run run && run.Descendants<Text>().Any())
                                {
                                    var text = run.Descendants<Text>().FirstOrDefault();
                                    if (text != null)
                                    {
                                        text.Text = newValue;
                                        break;
                                    }
                                }
                                nextElement = nextElement.NextSibling();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao processar Form Fields: {ex.Message}");
            }
        }

        private static void PreencherBookmarks(MainDocumentPart mainPart, dynamic dados)
        {
            try
            {
                var bookmarks = mainPart.Document.Body.Descendants<BookmarkStart>().ToList();
                Console.WriteLine($"Processando {bookmarks.Count} Bookmarks...");

                foreach (var bookmark in bookmarks)
                {
                    var bookmarkName = bookmark.Name?.Value;
                    if (string.IsNullOrEmpty(bookmarkName) || bookmarkName.StartsWith("_")) continue;

                    string newValue = FieldValueMapper.GetValueForField(bookmarkName, dados);

                    if (!string.IsNullOrEmpty(newValue))
                    {
                        Console.WriteLine($"  Preenchendo Bookmark '{bookmarkName}' com '{newValue}'");

                        // Encontrar o conteúdo entre BookmarkStart e BookmarkEnd
                        var bookmarkEnd = mainPart.Document.Body.Descendants<BookmarkEnd>()
                            .FirstOrDefault(be => be.Id == bookmark.Id);

                        if (bookmarkEnd != null)
                        {
                            // Inserir texto após o bookmark
                            var run = new Run(new Text(newValue));
                            bookmark.InsertAfterSelf(run);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao processar Bookmarks: {ex.Message}");
            }
        }

        private static void PreencherMergeFields(MainDocumentPart mainPart, dynamic dados)
        {
            try
            {
                var fieldCodes = mainPart.Document.Body.Descendants<FieldCode>().ToList();
                var mergeFieldsCount = fieldCodes.Count(fc => fc.Text.Contains("MERGEFIELD"));
                Console.WriteLine($"Processando {mergeFieldsCount} Merge Fields...");

                foreach (var fieldCode in fieldCodes)
                {
                    var code = fieldCode.Text;
                    if (code.Contains("MERGEFIELD"))
                    {
                        var parts = code.Split(' ');
                        if (parts.Length > 1)
                        {
                            var fieldName = parts[1].Trim('"');
                            string newValue = FieldValueMapper.GetValueForField(fieldName, dados);

                            if (!string.IsNullOrEmpty(newValue))
                            {
                                Console.WriteLine($"  Preenchendo Merge Field '{fieldName}' com '{newValue}'");

                                // Substituir o campo pelo valor
                                var run = fieldCode.Ancestors<Run>().FirstOrDefault();
                                if (run != null)
                                {
                                    run.RemoveAllChildren();
                                    run.AppendChild(new Text(newValue));
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao processar Merge Fields: {ex.Message}");
            }
        }

        private static void SubstituirPlaceholders(MainDocumentPart mainPart, dynamic dados)
        {
            try
            {
                var textElements = mainPart.Document.Body.Descendants<Text>().ToList();
                var placeholdersCount = textElements.Count(te => te.Text.Contains("{{") && te.Text.Contains("}}"));
                Console.WriteLine($"Processando {placeholdersCount} elementos com placeholders...");

                foreach (var textElement in textElements)
                {
                    if (textElement.Text.Contains("{{") && textElement.Text.Contains("}}"))
                    {
                        string originalText = textElement.Text;
                        string text = originalText;

                        // Replace placeholders usando reflection para acessar propriedades dinâmicas
                        var properties = dados.GetType().GetProperties();
                        foreach (var property in properties)
                        {
                            string placeholder = $"{{{{{property.Name}}}}}";
                            string value = property.GetValue(dados)?.ToString() ?? "";

                            if (text.Contains(placeholder))
                            {
                                text = text.Replace(placeholder, value);
                                Console.WriteLine($"  Substituindo '{placeholder}' por '{value}'");
                            }
                        }

                        textElement.Text = text;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao substituir placeholders: {ex.Message}");
            }
        }
    }
}