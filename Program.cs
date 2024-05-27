using System;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;

class Program
{
    static void Main()
    {
        // Configurar para usar EPPlus sem licença comercial
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        // Criar um novo pacote Excel
        using (ExcelPackage package = new ExcelPackage())
        {
            // Adicionar uma nova worksheet
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

            // Cabeçalhos da tabela
            string[] headers = { "Nome", "Sobrenome", "NomeSocial", "TelefoneCelular", "Email", "CPF", "Sexo", "DataNascimento", "CEP", "Rua", "Numero", "Complemento", "Bairro", "Cidade", "Estado" };

            // Adicionar os cabeçalhos na primeira linha
            for (int i = 0; i < headers.Length; i++)
            {
                worksheet.Cells[1, i + 1].Value = headers[i];
                worksheet.Cells[1, i + 1].Style.Font.Bold = true;
            }

            // Gerar dados randomicamente para as colunas Nome, Sobrenome e Email
            Random random = new Random();
            int rowCount = 3000; // Número de linhas a serem geradas

            for (int row = 2; row <= rowCount + 1; row++)
            {
                string nome = GenerateRandomString(random, 5, 10);
                string sobrenome = GenerateRandomString(random, 5, 10);
                string email = $"crmteste{row - 1:D2}@gmail.com";
                string telefoneCelular = $"3199999{(row - 1).ToString("D4")}";

                worksheet.Cells[row, 1].Value = nome;
                worksheet.Cells[row, 2].Value = sobrenome;
                worksheet.Cells[row, 4].Value = telefoneCelular;
                worksheet.Cells[row, 5].Value = email;
            }

            // Obter o caminho da área de trabalho do usuário
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string filePath = Path.Combine(desktopPath, "Resultado.xlsx");

            // Salvar o arquivo Excel
            FileInfo file = new FileInfo(filePath);
            package.SaveAs(file);
        }

        Console.WriteLine("Arquivo Excel criado com sucesso!");
    }

    static string GenerateRandomString(Random random, int minLength, int maxLength)
    {
        int length = random.Next(minLength, maxLength + 1);
        char[] letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz".ToCharArray();
        char[] randomString = new char[length];

        for (int i = 0; i < length; i++)
        {
            randomString[i] = letters[random.Next(letters.Length)];
        }

        return new string(randomString);
    }
}
