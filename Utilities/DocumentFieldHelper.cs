using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordAutoFill.Utilities
{
    public static class DocumentFieldHelper
    {
        public static string GetSolutionDirectory()
        {
            // Obter o diret�rio atual da aplica��o
            string currentDirectory = Directory.GetCurrentDirectory();

            // Navegar para cima at� encontrar o arquivo .sln
            DirectoryInfo directory = new DirectoryInfo(currentDirectory);

            while (directory != null)
            {
                // Procurar por arquivos .sln na pasta atual
                if (directory.GetFiles("*.sln").Length > 0)
                {
                    return directory.FullName;
                }

                // Subir um n�vel
                directory = directory.Parent;
            }

            // Se n�o encontrou a solution, usar o diret�rio atual
            Console.WriteLine("?? Arquivo .sln n�o encontrado. Usando diret�rio atual.");
            return currentDirectory;
        }

        public static string ObterTextoContentControl(OpenXmlElement contentControl)
        {
            var textElements = contentControl.Descendants<Text>();
            return string.Join("", textElements.Select(t => t.Text));
        }
    }
}