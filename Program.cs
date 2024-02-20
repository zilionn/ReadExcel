using Newtonsoft.Json;
using OfficeOpenXml;
using ReadExcel.Entities;

namespace ReadExcel {
    class Program {
        static void Main(string[] args) {
            var verificacoes = LerExcel();

            foreach(var item in verificacoes ) {
                Console.WriteLine($"Descrição: {item.Descricao}\n" +
                    $"Dipositivo: {item.Dispositivo}\n" +
                    $"Tolerância: {item.Tolerancia}\n" +
                    $"Peso: {item.Peso}");

                string json = JsonConvert.SerializeObject(item);
                Console.WriteLine($"{json}\n");
            }

            Console.ReadKey();
        }

        private static List<Verificacao> LerExcel() {
            var dados = new List<Verificacao>();

            FileInfo arquivo = new FileInfo(fileName: "C:\\Users\\fernando.correia\\Downloads\\criarModelos.xlsx");
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage package = new ExcelPackage(arquivo)) {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int colCont = worksheet.Dimension.End.Column;
                
                int rowCont = worksheet.Dimension.End.Row;

                for(int row = 2; row <= rowCont; row++) {
                    var verificacao = new Verificacao();
                    verificacao.Descricao = worksheet.Cells[row, Col: 1].Value.ToString();
                    verificacao.Peso = Convert.ToInt32(worksheet.Cells[row, Col: 2].Value.ToString());
                    verificacao.Dispositivo = worksheet.Cells[row, Col: 3].Value.ToString();
                    verificacao.Tolerancia = worksheet.Cells[row, Col: 4].Value.ToString();

                    dados.Add(verificacao);
                }
            }
            return dados;
        }
    }

}