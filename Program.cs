using System.Text;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;
using Tesseract; // Necessário instalar o pacote NuGet: Tesseract

namespace RenomeadorPDF
{
    class Program
    {

        private static readonly string[] identificadoresPdfValidos = { "ASO", "ATESTADO DE SAÚDE OCUPACIONAL" };

        // Configure o caminho para a pasta tessdata (onde ficam os arquivos de linguagem do OCR)
        // Baixe o arquivo 'por.traineddata' ou 'eng.traineddata' e coloque numa pasta "tessdata" junto com o executável
        private static readonly string TessDataPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tessdata");
        private static readonly string Language = "por"; // 'por' para português, 'eng' para inglês

        static void Main(string[] args)
        {
            Console.WriteLine("=== Iniciando Processador de PDFs Híbrido (Nativo + OCR) ===");

            string currentDir = Directory.GetCurrentDirectory();
            var files = Directory.GetFiles(currentDir, "*.pdf");

            if (files.Length == 0)
            {
                Console.WriteLine("Nenhum arquivo PDF encontrado na pasta atual.");
                return;
            }

            Console.WriteLine($"Encontrados {files.Length} arquivos. Iniciando processamento...\n");

            files.ToList().ForEach(ProcessarArquivo);

            Console.WriteLine("\nProcessamento concluído. Pressione qualquer tecla para sair.");
            Console.ReadKey();
        }

        static void ProcessarArquivo(string filePath)
        {
            string nomeArquivoOriginal = Path.GetFileName(filePath);
            Console.WriteLine($"-> Lendo: {nomeArquivoOriginal}");

            try
            {
                int qtdPaginas = 0;
                string terceiraPalavra = "NaoIdentificado";
                string textoExtraido = ExtrairTExtoPdfNativo(filePath, out qtdPaginas);

                if (String.IsNullOrEmpty(textoExtraido)) textoExtraido = ExtrairTextoPdfOcr(filePath, out qtdPaginas);

                if (String.IsNullOrEmpty(textoExtraido))
                {
                    Console.WriteLine("   [FALHA] Não foi possível extrair texto do PDF, mesmo via OCR. Pulando arquivo.\n");
                    return;
                }

                if (!textoPdfAsoValida(textoExtraido))
                {
                    Console.WriteLine("   [AVISO] Pdf não reconhecido como ASO ou PRONTUARIO DIGITALIZADO .\n");
                    return;
                }

                string nomeFuncionario = ExtrairNomeFuncionario(normalizeString(textoExtraido)).ToUpper();
                string dataAso = ExtrairDataAso(normalizeString(textoExtraido)).ToUpper();
                string tipoAso = ExtrairTipoAso(normalizeString(textoExtraido)).ToUpper();

                string inicio = qtdPaginas > 1 ? "PRONTUARIO DIGITALIZADO" : "ASO DIGITALIZADO";

                // Monta o novo nome: {QtdPaginas}_{TerceiraPalavra}_{Data}.pdf
                string dataAtual = DateTime.Now.ToString("yyyyMMdd");
                string novoNome = $"{inicio} - {nomeFuncionario} {tipoAso} {dataAtual}.pdf";


                Path.GetInvalidFileNameChars().ToList().ForEach(c => { novoNome = novoNome.Replace(c, '_'); });

                string novoCaminho = Path.Combine(Path.GetDirectoryName(filePath), novoNome);

                // Evita sobrescrever se já existir (adiciona contador)
                int count = 1;
                while (File.Exists(novoCaminho) && novoCaminho != filePath)
                {
                    string nomeSemExt = Path.GetFileNameWithoutExtension(novoNome);
                    novoCaminho = Path.Combine(Path.GetDirectoryName(filePath), $"{nomeSemExt}_{count}.pdf");
                    count++;
                }

                if (filePath != novoCaminho)
                {
                    File.Move(filePath, novoCaminho);
                    Console.WriteLine($"   RENOMEADO PARA: {Path.GetFileName(novoCaminho)}\n");
                }
                else
                {
                    Console.WriteLine("   Arquivo já está com o nome correto ou não pôde ser alterado.\n");
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine($"   [ERRO CRÍTICO] Não foi possível processar o arquivo: {ex.Message}\n");
            }
        }

        private static String ExtrairTipoAso(string textoExtraido)
        {
            return textoExtraido.Contains("Admissional".ToUpper()) ? "AMD" : "PER";
        }

        private static string ExtrairDataAso(string textoExtraido)
        {
            return "21112025";
        }

        private static string ExtrairNomeFuncionario(string textoExtraido)
        {
            return "Lucas Coutinho Bezerra";
        }

        private static bool textoPdfAsoValida(string textoExtraido)
        {
            if (string.IsNullOrEmpty(textoExtraido)) return false;

            return identificadoresPdfValidos.Any(x => textoExtraido.Contains(x, StringComparison.OrdinalIgnoreCase));

        }

        private static string ExtrairTExtoPdfNativo(string filePath, out int qtdPaginas)
        {
            using (PdfDocument pdf = PdfDocument.Open(filePath))
            {
                qtdPaginas = pdf.NumberOfPages;
                if (qtdPaginas > 0)
                {
                    var page = pdf.GetPage(1);
                    string textoNativo = page.Text;

                    if (!string.IsNullOrWhiteSpace(textoNativo) && textoNativo.Length > 10) return textoNativo;
                }
            }
            return String.Empty;
        }

        private static string ExtrairTextoPdfOcr(string filePath, out int qtdPaginas)
        {
            using (PdfDocument pdf = PdfDocument.Open(filePath))
            {
                qtdPaginas = pdf.NumberOfPages;
                StringBuilder sbTextoOcr = new StringBuilder();

                // Verifica se a pasta do Tesseract existe antes de entrar no loop
                if (Directory.Exists(TessDataPath))
                {
                    using (var engine = new TesseractEngine(TessDataPath, Language, EngineMode.Default))
                    {
                        // Loop por TODAS as páginas
                        for (int i = 1; i <= qtdPaginas; i++)
                        {
                            // Se já achamos texto suficiente (ex: mais de 50 caracteres), paramos para economizar tempo
                            // (Remova este 'if' se quiser OBRIGAR a ler o arquivo inteiro mesmo já tendo a 3ª palavra)
                            if (sbTextoOcr.Length > 300) break;

                            var page = pdf.GetPage(i);
                            var images = page.GetImages(); // Pega TODAS as imagens da página

                            if (images.Count() > 0)
                            {
                                Console.WriteLine($"     -> Processando pág {i} ({images.Count()} imagens)...");

                                foreach (var image in images)
                                {
                                    try
                                    {
                                        byte[] imageBytes = image.RawBytes.ToArray();
                                        using (var img = Pix.LoadFromMemory(imageBytes))
                                        {
                                            using (var pageOcr = engine.Process(img))
                                            {
                                                string txt = pageOcr.GetText();
                                                if (!string.IsNullOrWhiteSpace(txt))
                                                {
                                                    sbTextoOcr.Append(txt + " ");
                                                }
                                            }
                                        }
                                    }
                                    catch
                                    {
                                        // Ignora imagens que derem erro (formatos não suportados pelo Tesseract direto)
                                        continue;
                                    }
                                }
                            }
                        }
                    }

                    string resultadoOcr = sbTextoOcr.ToString();
                    Console.WriteLine($"   [OCR Sucesso] Texto recuperado via OCR.");
                    return resultadoOcr;
                }
            }
            return String.Empty;
        }

        private static String normalizeString(string texto)
        {
            if (string.IsNullOrWhiteSpace(texto)) return "Vazio";

            // Normaliza quebras de linha e espaços extras
            texto = Regex.Replace(texto, @"\s+", " ").Trim();


            texto = texto.Split(' ')
                                .Where(p => !string.IsNullOrWhiteSpace(p) && p.Length > 1) // Filtra letras soltas ou espaços
                                .Select(p => Regex.Replace(p, @"[^\w\d]", "")) // Remove pontuação
                                .ToArray().Aggregate((a, b) => a + " " + b);

            return texto;
        }

    }
}
