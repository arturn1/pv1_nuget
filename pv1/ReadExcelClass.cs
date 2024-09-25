using ClosedXML.Excel;
using OfficeOpenXml;
using System.Dynamic;
using System.Text.Json;

namespace ExcelUtilityLib
{
    public class ExcelHandler
    {
        /// <summary>
        /// Lê e combina dados de múltiplos arquivos Excel.
        /// </summary>
        /// <param name="filePaths">Lista de caminhos para os arquivos Excel.</param>
        /// <returns>Retorna uma lista de objetos dinâmicos contendo os dados do Excel.</returns>
        public static List<dynamic> ReadMultipleExcelFiles(List<string> filePaths)
        {
            var listaDeObjetos = new List<dynamic>();

            foreach (var filePath in filePaths)
            {
                using (var workbook = new XLWorkbook(filePath))
                {
                    var worksheet = workbook.Worksheets.Worksheet(1); // Assume que estamos lendo a primeira planilha
                    var ultimaLinha = worksheet.LastRowUsed().RowNumber();
                    var ultimaColuna = worksheet.LastColumnUsed().ColumnNumber();

                    // Leitura dos cabeçalhos (nomes das colunas)
                    var nomesPropriedades = new List<string>();
                    for (int coluna = 1; coluna <= ultimaColuna; coluna++)
                    {
                        nomesPropriedades.Add(worksheet.Cell(1, coluna).GetString().Trim());
                    }

                    // Leitura dos valores do Excel
                    for (int linha = 2; linha <= ultimaLinha; linha++)
                    {
                        dynamic objeto = new ExpandoObject();
                        var dicionarioObjeto = (IDictionary<string, object>)objeto;

                        for (int coluna = 1; coluna <= ultimaColuna; coluna++)
                        {
                            var valorCelula = worksheet.Cell(linha, coluna).Value.ToString().Trim();
                            dicionarioObjeto[nomesPropriedades[coluna - 1]] = valorCelula;
                        }

                        listaDeObjetos.Add(objeto);
                    }
                }
            }

            return listaDeObjetos;
        }

        /// <summary>
        /// Gera um arquivo Excel a partir de uma lista de objetos dinâmicos.
        /// </summary>
        /// <param name="data">Lista de dicionários representando os dados (chave-valor).</param>
        /// <param name="worksheetName">Nome da planilha que será gerada.</param>
        /// <returns>Retorna o arquivo Excel como um array de bytes.</returns>
        public static byte[] GenerateExcelFromData(List<Dictionary<string, object>> data, string worksheetName = "Sheet1")
        {
            if (data == null || !data.Any())
            {
                throw new ArgumentException("Os dados fornecidos estão vazios.");
            }

            // Cria um novo pacote Excel
            using (var package = new ExcelPackage())
            {
                // Adiciona a planilha com o nome fornecido
                var worksheet = package.Workbook.Worksheets.Add(worksheetName);

                // Extrai os cabeçalhos do primeiro objeto da lista
                var headers = data.First().Keys.ToList();

                // Adiciona os cabeçalhos à primeira linha
                for (int i = 0; i < headers.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = headers[i];
                }

                // Adiciona os dados às linhas subsequentes
                for (int i = 0; i < data.Count; i++)
                {
                    var rowData = data[i];
                    for (int j = 0; j < headers.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1].Value = rowData[headers[j]]?.ToString();
                    }
                }

                // Auto ajusta as colunas de acordo com o conteúdo
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                // Retorna o arquivo Excel como um array de bytes
                return package.GetAsByteArray();
            }
        }

        /// <summary>
        /// Gera um arquivo Excel a partir de um JSON string.
        /// </summary>
        /// <param name="jsonInput">O JSON string contendo os dados para o Excel.</param>
        /// <param name="worksheetName">Nome da planilha que será gerada.</param>
        /// <returns>Retorna o arquivo Excel como um array de bytes.</returns>
        public static byte[] GenerateExcelFromJson(string jsonInput, string worksheetName = "Sheet1")
        {
            if (string.IsNullOrEmpty(jsonInput))
            {
                throw new ArgumentException("JSON input is required.");
            }

            try
            {
                // Desserializa o JSON para uma lista de dicionários
                var data = JsonSerializer.Deserialize<List<Dictionary<string, object>>>(jsonInput);

                // Chama o método genérico para gerar o Excel a partir dos dados
                return GenerateExcelFromData(data, worksheetName);
            }
            catch (JsonException)
            {
                throw new InvalidOperationException("Formato de JSON inválido.");
            }
        }
    }
}
