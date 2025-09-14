using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordAutoFill.Utilities;

namespace WordAutoFill.Core
{
    public static class DocumentFieldAnalyzer
    {
        public static void AnalisarCamposEditaveis(string documentPath)
        {
            using (WordprocessingDocument document = WordprocessingDocument.Open(documentPath, false))
            {
                MainDocumentPart? mainPart = document.MainDocumentPart;
                if (mainPart?.Document?.Body == null)
                {
                    Console.WriteLine("Documento inválido ou corrompido.");
                    return;
                }

                var camposEncontrados = new List<string>();

                // 1. Analisar Content Controls (Controles de Conteúdo)
                AnalyzeContentControls(mainPart, camposEncontrados);

                // 2. Analisar Form Fields (Campos de Formulário Legacy)
                AnalyzeFormFields(mainPart, camposEncontrados);

                // 3. Analisar Bookmarks (Marcadores)
                AnalyzeBookmarks(mainPart, camposEncontrados);

                // 4. Procurar por placeholders de texto como {{Campo}}
                AnalyzeTextPlaceholders(mainPart, camposEncontrados);

                // 5. Analisar campos MERGEFIELD
                AnalyzeMergeFields(mainPart, camposEncontrados);

                // Resumo
                DisplaySummary(camposEncontrados);
            }
        }

        private static void AnalyzeContentControls(MainDocumentPart mainPart, List<string> camposEncontrados)
        {
            Console.WriteLine("\n1. CONTENT CONTROLS (Controles de Conteúdo):");
            var contentControls = new List<OpenXmlElement>();
            contentControls.AddRange(mainPart.Document.Body.Descendants<SdtBlock>());
            contentControls.AddRange(mainPart.Document.Body.Descendants<SdtRun>());
            contentControls.AddRange(mainPart.Document.Body.Descendants<SdtCell>());

            if (contentControls.Any())
            {
                foreach (var contentControl in contentControls)
                {
                    var sdtProperties = contentControl.Descendants<SdtProperties>().FirstOrDefault();
                    if (sdtProperties != null)
                    {
                        var tag = sdtProperties.Descendants<Tag>().FirstOrDefault()?.Val?.Value;
                        var alias = sdtProperties.Descendants<SdtAlias>().FirstOrDefault()?.Val?.Value;

                        var currentText = DocumentFieldHelper.ObterTextoContentControl(contentControl);

                        Console.WriteLine($"   • Tag: '{tag ?? "N/A"}', Alias: '{alias ?? "N/A"}', Texto Atual: '{currentText}'");

                        if (!string.IsNullOrEmpty(tag))
                            camposEncontrados.Add($"ContentControl-Tag: {tag}");
                    }
                }
            }
            else
            {
                Console.WriteLine("   Nenhum Content Control encontrado.");
            }
        }

        private static void AnalyzeFormFields(MainDocumentPart mainPart, List<string> camposEncontrados)
        {
            Console.WriteLine("\n2. FORM FIELDS (Campos de Formulário Legacy):");
            var formFields = mainPart.Document.Body.Descendants<FormFieldData>().ToList();

            if (formFields.Any())
            {
                foreach (var formField in formFields)
                {
                    var fieldName = formField.XName?.LocalName;
                    var fieldType = "Desconhecido";

                    // Identificar tipo do campo
                    if (formField.Parent?.Parent is Run run)
                    {
                        var fieldChar = run.Descendants<FieldChar>().FirstOrDefault();
                        if (fieldChar != null)
                        {
                            fieldType = "Text Field";
                        }
                    }

                    Console.WriteLine($"   • Nome: '{fieldName ?? "N/A"}', Tipo: {fieldType}");

                    if (!string.IsNullOrEmpty(fieldName))
                        camposEncontrados.Add($"FormField: {fieldName}");
                }
            }
            else
            {
                Console.WriteLine("   Nenhum Form Field encontrado.");
            }
        }

        private static void AnalyzeBookmarks(MainDocumentPart mainPart, List<string> camposEncontrados)
        {
            Console.WriteLine("\n3. BOOKMARKS (Marcadores):");
            var bookmarks = mainPart.Document.Body.Descendants<BookmarkStart>().ToList();

            if (bookmarks.Any())
            {
                foreach (var bookmark in bookmarks)
                {
                    var bookmarkName = bookmark.Name?.Value;
                    if (!string.IsNullOrEmpty(bookmarkName) && !bookmarkName.StartsWith("_"))
                    {
                        Console.WriteLine($"   • Nome: '{bookmarkName}'");
                        camposEncontrados.Add($"Bookmark: {bookmarkName}");
                    }
                }
            }
            else
            {
                Console.WriteLine("   Nenhum Bookmark personalizado encontrado.");
            }
        }

        private static void AnalyzeTextPlaceholders(MainDocumentPart mainPart, List<string> camposEncontrados)
        {
            Console.WriteLine("\n4. PLACEHOLDERS DE TEXTO ({{campo}}):");
            var paragraphs = mainPart.Document.Body.Descendants<Paragraph>();
            var placeholders = new HashSet<string>();

            foreach (var paragraph in paragraphs)
            {
                var fullText = string.Concat(paragraph.Descendants<Text>().Select(t => t.Text));
                var matches = System.Text.RegularExpressions.Regex.Matches(fullText, @"\{\{[^{}]+\}\}");

                foreach (System.Text.RegularExpressions.Match match in matches)
                {
                    placeholders.Add(match.Value);
                }
            }

            if (placeholders.Any())
            {
                foreach (var placeholder in placeholders)
                {
                    Console.WriteLine($"   • Placeholder: '{placeholder}'");
                    camposEncontrados.Add($"TextPlaceholder: {placeholder}");
                }
            }
            else
            {
                Console.WriteLine("   Nenhum placeholder de texto encontrado.");
            }
        }

        private static void AnalyzeMergeFields(MainDocumentPart mainPart, List<string> camposEncontrados)
        {
            Console.WriteLine("\n5. MERGE FIELDS (Campos de Mala Direta):");
            var fieldCodes = mainPart.Document.Body.Descendants<FieldCode>().ToList();

            if (fieldCodes.Any())
            {
                foreach (var fieldCode in fieldCodes)
                {
                    var code = fieldCode.Text;
                    if (code.Contains("MERGEFIELD"))
                    {
                        // Extrair nome do campo MERGEFIELD
                        var parts = code.Split(' ');
                        if (parts.Length > 1)
                        {
                            var fieldName = parts[1].Trim('"');
                            Console.WriteLine($"   • MergeField: '{fieldName}'");
                            camposEncontrados.Add($"MergeField: {fieldName}");
                        }
                    }
                }
            }
            else
            {
                Console.WriteLine("   Nenhum Merge Field encontrado.");
            }
        }

        private static void DisplaySummary(List<string> camposEncontrados)
        {
            Console.WriteLine($"\n=== RESUMO ===");
            Console.WriteLine($"Total de campos editáveis encontrados: {camposEncontrados.Count}");

            if (camposEncontrados.Any())
            {
                Console.WriteLine("\nLista completa de campos:");
                foreach (var campo in camposEncontrados)
                {
                    Console.WriteLine($"  - {campo}");
                }
            }
        }

        private static Dictionary<string, string> DadosParaDicionario(dynamic dados)
        {
            var dict = new Dictionary<string, string>();
            var properties = dados.GetType().GetProperties();
            foreach (var property in properties)
            {
                string placeholder = $"{{{{{property.Name}}}}}";
                string value = property.GetValue(dados)?.ToString() ?? "";
                dict[placeholder] = value;
            }
            return dict;
        }
    }
}