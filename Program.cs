using WordAutoFill.Core;
using WordAutoFill.Models;
using WordAutoFill.Utilities;

namespace WordAutoFill
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Word document - using solution directory
            string solutionDirectory = DocumentFieldHelper.GetSolutionDirectory();
            string documentPath = Path.Combine(solutionDirectory, "Exemplo_Campos_Editaveis.docx");
            string outputPath = Path.Combine(solutionDirectory, "Documento_Preenchido.docx");

            try
            {
                Console.WriteLine($"Diretório da solution: {solutionDirectory}");
                Console.WriteLine($"Procurando arquivo em: {documentPath}");

                // Check if source file exists
                if (!File.Exists(documentPath))
                {
                    Console.WriteLine($"Arquivo não encontrado: {documentPath}");
                    Console.WriteLine("Certifique-se de que o arquivo 'Exemplo_Campos_Editaveis.docx' está na pasta da solution (.sln)");
                    return;
                }

                Console.WriteLine("✅ Arquivo encontrado!");

                // First, analyze the document to identify editable fields
                Console.WriteLine("\n=== ANÁLISE DOS CAMPOS EDITÁVEIS ===");
                DocumentFieldAnalyzer.AnalisarCamposEditaveis(documentPath);

                // Create a copy of the original document to work with
                File.Copy(documentPath, outputPath, true);

                // Get sample data
                var dadosPreenchimento = DocumentData.GetSampleData();

                Console.WriteLine("\n=== PREENCHIMENTO DO DOCUMENTO ===");
                // Fill the document
                DocumentFieldFiller.PreencherDocumento(outputPath, dadosPreenchimento);

                Console.WriteLine("Documento preenchido com sucesso!");
                Console.WriteLine($"Arquivo salvo como: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao processar o documento: {ex.Message}");
                Console.WriteLine($"Detalhes: {ex.InnerException?.Message}");
            }
        }
    }
}