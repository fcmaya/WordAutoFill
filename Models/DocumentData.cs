namespace WordAutoFill.Models
{
    public static class DocumentData
    {
        public static dynamic GetSampleData()
        {
            return new
            {
                // Identificação das Partes
                Nome = "João Silva Santos",
                CPF = "123.456.789-00",
                Endereco = "Rua das Flores, 123",
                Cidade = "São Paulo",
                UF = "SP",
                CEP = "01234-567",
                Telefone = "(11) 99999-9999",
                Email = "joao.silva@email.com",
                // Valor de Crédito
                ValorConcedido = "R$ 150.000,00",
                PrazoPagamento = "60 meses",
                QuantidadeParcelas = "60",
                ValorParcela = "R$ 3.247,82",
                // Taxa de Juros e Encargos
                JurosMes = "1,25%",
                JurosAno = "15,00%",
                OutrosEncargos = "IOF: 0,38% | Tarifa de Cadastro: R$ 350,00",
                // Garantias
                Garantias = "Alienação fiduciária do veículo",
                // Foro
                Local = "São Paulo",
                Data = DateTime.Now.ToString("dd/MM/yyyy")
            };
        }
    }
}