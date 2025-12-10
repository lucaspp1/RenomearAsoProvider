using System.Text;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;
using Tesseract;

namespace RenomeadorPDF
{
    class Program
    {
        // ASO- ATESTADO DE SAÚDE
        private static readonly string[] _identificadoresPdfValidos = { "ASO", "ATESTADO DE SAÚDE OCUPACIONAL" };
        // Identificadores execos ASO Ocupacional 
        private static readonly string[] _identificadoresAsoValidas = { "Ocupacional" };

        // Configure o caminho para a pasta tessdata (onde ficam os arquivos de linguagem do OCR)
        // Baixe o arquivo 'por.traineddata' ou 'eng.traineddata' e coloque numa pasta "tessdata" junto com o executável
        private static readonly string TessDataPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tessdata");
        private static readonly string Language = "por"; // 'por' para português, 'eng' para inglês
        private static readonly string _TIPO_ASO_INDEFINIDO = "TIPO-INDEFINIDO";
        private static readonly string _SEM_DATA = "SEMDATA";
        private static readonly bool showLog = true;

        // 2. DEFINIÇÃO DOS GATILHOS (Onde começa e onde termina o nome)
        // Palavras que indicam que o nome vem a seguir
        private static readonly string _gatilhosInicio = @"(Empregado|Funcionário|Nome:|Colaborador|TRABALHADOR Nome|Paciente)";

        // Palavras que indicam que o nome acabou (o que costuma vir depois do nome no ASO)
        // \d+ pega qualquer número (como idade, matrícula ou CPF começando)
        // M\s pega "M " de matrícula
        private static readonly string _gatilhosFim = @"(Deptosetor|SEXO|SEQUENCIAL|Depto|Setor|Cargo|CPF|RG|CNPJ|Data|Nasc|Idade|M\s|\d|CBO|SEXO|Código|Código /Matrícula)";

        private static string getFatherCurrentDirectory()
        {
            // Obtém o caminho base do AppDomain (onde o executável está rodando)
            string caminhoDoExecutavel = AppDomain.CurrentDomain.BaseDirectory;

            // Usa Directory.GetParent para obter o diretório pai (DirectoryInfo)
            DirectoryInfo? diretorioPai = Directory.GetParent(caminhoDoExecutavel);

            if (diretorioPai != null) return diretorioPai.FullName;
            else return caminhoDoExecutavel;
        }

        static void Main(string[] args)
        {
            Console.WriteLine("=== Iniciando Processador de PDFs Híbrido (Nativo + OCR) ===");

            string currentDir = getFatherCurrentDirectory();
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
                string textoExtraido = ExtrairTExtoPdfNativo(filePath, out qtdPaginas);

                if (String.IsNullOrEmpty(textoExtraido)) textoExtraido = ExtrairTextoPdfOcr(filePath, out qtdPaginas);

                if (String.IsNullOrEmpty(textoExtraido))
                {
                    Console.WriteLine("   [FALHA] Não foi possível extrair texto do PDF, mesmo via OCR. Pulando arquivo.\n");
                    return;
                }

                textoExtraido = normalizeString(textoExtraido);

                if (!textoPdfAsoValida(textoExtraido))
                {
                    if (showLog) Console.WriteLine(" \n textoPdfAsoValida ====== \n" + textoExtraido + " \n ====== \n");
                    Console.WriteLine("   [AVISO] Pdf não reconhecido como ASO ou PRONTUARIO DIGITALIZADO .\n");
                    return;
                }


                string nomeFuncionario = ExtrairNomeFuncionario(textoExtraido).ToUpper();
                string dataAso = ExtrairDataAso(textoExtraido).ToUpper();
                string tipoAso = ExtrairTipoAso(textoExtraido).ToUpper();

                bool erroExtrairDados = false;


                if (String.IsNullOrEmpty(nomeFuncionario) || dataAso == _SEM_DATA || tipoAso == _TIPO_ASO_INDEFINIDO)
                {
                    if (showLog)
                    { Console.WriteLine("   [AVISO] Texto extraido: \n " + textoExtraido + "\n"); }
                    erroExtrairDados = true;
                }

                string inicio = qtdPaginas > 3 ? "PRONTUARIO DIGITALIZADO" : "ASO DIGITALIZADO";
                string dataAtual = DateTime.Now.ToString("yyyyMMdd");
                string novoNome = $"{inicio} - {nomeFuncionario} {tipoAso} {dataAso}.pdf";

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

                if (erroExtrairDados)
                { MovFileSubFolder(novoCaminho, "ERROS"); }
                else
                { MovFileSubFolder(novoCaminho, "PROCESSADOS"); }

            }
            catch (Exception ex)
            {
                Console.WriteLine($"   [ERRO CRÍTICO] Não foi possível processar o arquivo: {ex.Message}\n");
            }
        }

        private static string ExtrairTipoAso(string textoExtraido)
        {
            string texto = textoExtraido?.ToUpper() ?? "";

            return texto switch
            {

                var x when x.Contains("ADMISSIONAL") => "ADM",
                var x when x.Contains("PERIODICO") => "PER",
                var x when x.Contains("Periódico".ToUpper()) => "PER",
                var x when x.Contains("RETORNO") => "RT",
                var x when x.Contains("MUDANCA") && x.Contains("FUNCAO") => "MF",
                var x when x.Contains("DEMISSIONAL") => "DEM",
                _ => _TIPO_ASO_INDEFINIDO // Caso não ache nenhum
            };
        }

        private static string ExtrairDataAso(string textoExtraido)
        {
            if (string.IsNullOrEmpty(textoExtraido)) return "Texto vazio";

            // 1. PRÉ-PROCESSAMENTO
            // Remove caracteres que não ajudam na data, mantendo números, barras e pontos
            // Mas mantemos o contexto para ver se é nascimento
            string textoLimpo = Regex.Replace(textoExtraido, @"\s+", " ");

            // 2. REGEX PARA CAPTURAR DATAS
            // Padrão 1: Datas com separadores (13/11/2025 ou 13.11.25)
            // Padrão 2: Datas coladas (13112025) - comum no seu Texto 1
            // Grupo 1: Dia, Grupo 2: Mês, Grupo 3: Ano
            string pattern = @"\b(?<dia>\d{2})[\/\.-]?(?<mes>\d{2})[\/\.-]?(?<ano>\d{4})\b";

            var matches = Regex.Matches(textoLimpo, pattern);
            var datasCandidatas = new List<DateTime>();

            foreach (Match match in matches)
            {
                int dia = int.Parse(match.Groups["dia"].Value);
                int mes = int.Parse(match.Groups["mes"].Value);
                int ano = int.Parse(match.Groups["ano"].Value);

                // 3. VALIDAÇÃO BÁSICA (É uma data válida?)
                if (dia < 1 || dia > 31 || mes < 1 || mes > 12) continue;

                // Tenta criar um objeto DateTime para garantir que o dia existe no mês (ex: 30/02 falharia)
                if (DateTime.TryParse($"{dia}/{mes}/{ano}", out DateTime dataValida))
                {
                    // 4. FILTRAGEM DE CONTEXTO (IMPORTANTE)

                    // REGRA A: Ignorar Datas Antigas (Provavelmente Nascimento)
                    // Assumimos que um ASO válido é de no máximo 5 anos atrás.
                    if (dataValida.Year < DateTime.Now.Year - 5) continue;

                    // REGRA B: Verificar se a palavra "Nasc" ou "Nascimento" está imediatamente antes
                    int indexMatch = match.Index;
                    int contextoInicio = Math.Max(0, indexMatch - 20); // Olha 20 caracteres para trás
                    string contextoAnterior = textoLimpo.Substring(contextoInicio, indexMatch - contextoInicio).ToLower();

                    if (contextoAnterior.Contains("nasc")) continue;

                    datasCandidatas.Add(dataValida);
                }
            }

            // 5. SELEÇÃO DA MELHOR DATA
            if (datasCandidatas.Count > 0)
            {
                // Geralmente a data do ASO aparece mais de uma vez ou no final (assinatura).
                // A estratégia segura é pegar a data mais recente encontrada que não seja futura demais.
                // Ou simplesmente pegar a última ocorrência válida no documento (assinatura).

                string resultado = datasCandidatas.Last().ToString("ddMMyyyy");
                if (showLog) Console.WriteLine("data: " + resultado);

                if (textoExtraido.LastIndexOf(resultado) < 10) return _SEM_DATA;
                return resultado;
            }

            return _SEM_DATA;
        }

        private static string ExtrairNomeFuncionario(string textoExtraido)
        {
            if (string.IsNullOrEmpty(textoExtraido)) return String.Empty;

            // 1. PRÉ-PROCESSAMENTO
            // Remove quebras de linha e múltiplos espaços para facilitar a busca
            string textoLimpo = Regex.Replace(textoExtraido, @"\s+", " ").Trim();


            // 3. MONTAGEM DA REGEX
            // Explicação da Regex:
            // (?i)          -> Case insensitive (ignora maiúsculas/minúsculas)
            // \b{inicio}\b  -> Encontra a palavra chave inteira (ex: "Empregado")
            // [:\s]* -> Aceita opcionais dois pontos ou espaços após a palavra chave
            // (?<nome>.*?)  -> CAPTURA o conteúdo (o nome) de forma não-gulosa (para o mais cedo possível)
            // (?=\s+{fim})  -> Lookahead: Para a captura quando encontrar um espaço seguido de um gatilho de fim
            string pattern = $"(?i)\\b{_gatilhosInicio}\\b[:\\s]*(?<nome>.*?)(?=\\s+{_gatilhosFim})";

            Match match = Regex.Match(textoLimpo, pattern);

            if (match.Success) return match.Groups["nome"].Value.Trim().ToUpper().Replace("NOME ", "");

            return string.Empty;
        }

        private static bool textoPdfAsoValida(string textoExtraido)
        {
            if (string.IsNullOrEmpty(textoExtraido)) return false;

            bool isPdfValido = _identificadoresPdfValidos.Any(x => textoExtraido.Contains(x, StringComparison.OrdinalIgnoreCase));

            if (!isPdfValido)
            {
                _identificadoresAsoValidas.ToList().ForEach(identificador =>
                {
                    if (textoExtraido.IndexOf(identificador, StringComparison.OrdinalIgnoreCase) <= 10) isPdfValido = true;
                });
            }

            return isPdfValido;

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
                            if (sbTextoOcr.Length > 500) break;

                            var page = pdf.GetPage(i);
                            var images = page.GetImages(); // Pega TODAS as imagens da página

                            if (images.Count() > 0)
                            {
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

        private static bool MovFileSubFolder(string sourceFilePath, string destinationFolder)
        {
            try
            {
                if (!Directory.Exists(destinationFolder))
                    Directory.CreateDirectory(destinationFolder);

                string fileName = Path.GetFileName(sourceFilePath);
                string destFilePath = Path.Combine(destinationFolder, fileName);
                File.Move(sourceFilePath, destFilePath);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   [ERRO] Não foi possível mover o arquivo: {ex.Message}\n");
                return false;
            }

        }

    }
}