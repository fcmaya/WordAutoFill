using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordAutoFill.Utilities
{
    public static class DocumentFieldHelper
    {
        public static string GetSolutionDirectory()
        {
            // Obter o diretório atual da aplicação
            string currentDirectory = Directory.GetCurrentDirectory();

            // Navegar para cima até encontrar o arquivo .sln
            DirectoryInfo directory = new DirectoryInfo(currentDirectory);

            while (directory != null)
            {
                // Procurar por arquivos .sln na pasta atual
                if (directory.GetFiles("*.sln").Length > 0)
                {
                    return directory.FullName;
                }

                // Subir um nível
                directory = directory.Parent;
            }

            // Se não encontrou a solution, usar o diretório atual
            Console.WriteLine("?? Arquivo .sln não encontrado. Usando diretório atual.");
            return currentDirectory;
        }

        public static string ObterTextoContentControl(OpenXmlElement contentControl)
        {
            var textElements = contentControl.Descendants<Text>();
            return string.Join("", textElements.Select(t => t.Text));
        }
    }
}