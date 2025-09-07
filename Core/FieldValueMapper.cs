namespace MinutaDocx.Core
{
    public static class FieldValueMapper
    {
        public static string GetValueForField(string fieldName, dynamic dados)
        {
            if (string.IsNullOrEmpty(fieldName)) return "";

            try
            {
                var properties = dados.GetType().GetProperties();
                foreach (var property in properties)
                {
                    if (string.Equals(property.Name, fieldName, StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(property.Name, fieldName.Replace("_", ""), StringComparison.OrdinalIgnoreCase))
                    {
                        return property.GetValue(dados)?.ToString() ?? "";
                    }
                }

                // Mapeamento manual organizado por categorias (conforme DocumentData)
                return fieldName.ToLower() switch
                {
                    // === IDENTIFICAÇÃO DAS PARTES ===
                    "nome" or "name" or "nomecompleto" or "fullname" => GetPropertyValue(dados, "Nome"),
                    "cpf" or "documento" or "document" => GetPropertyValue(dados, "CPF"),
                    "endereco" or "address" => GetPropertyValue(dados, "Endereco"),
                    "cidade" or "city" => GetPropertyValue(dados, "Cidade"),
                    "uf" or "estado" or "unidadefederativa" => GetPropertyValue(dados, "UF"),
                    "cep" or "postalcode" or "zipcode" => GetPropertyValue(dados, "CEP"),
                    "telefone" or "phone" or "celular" or "telephone" => GetPropertyValue(dados, "Telefone"),
                    "email" or "emailaddress" or "e-mail" => GetPropertyValue(dados, "Email"),

                    // === VALOR DE CRÉDITO ===
                    "valorconcedido" or "valor_concedido" or "valoremprestimo" or "credito" => GetPropertyValue(dados, "ValorConcedido"),
                    "prazopagamento" or "prazo_pagamento" or "prazo" => GetPropertyValue(dados, "PrazoPagamento"),
                    "quantidadeparcelas" or "quantidade_parcelas" or "parcelas" or "numeroparcelas" => GetPropertyValue(dados, "QuantidadeParcelas"),
                    "valorparcela" or "valor_parcela" or "parcela" => GetPropertyValue(dados, "ValorParcela"),

                    // === TAXA DE JUROS E ENCARGOS ===
                    "jurosmes" or "juros_mes" or "jurosmensal" or "taxames" => GetPropertyValue(dados, "JurosMes"),
                    "jurosano" or "juros_ano" or "jurosanual" or "taxaano" => GetPropertyValue(dados, "JurosAno"),
                    "outrosencargos" or "outros_encargos" or "encargos" or "taxas" => GetPropertyValue(dados, "OutrosEncargos"),

                    // === GARANTIAS ===
                    "garantias" or "garantia" or "caucao" => GetPropertyValue(dados, "Garantias"),

                    // === FORO ===
                    "local" or "localidade" or "foro" or "jurisdicao" => GetPropertyValue(dados, "Local"),
                    "data" or "dataassinatura" or "data_assinatura" or "datacontrato" => GetPropertyValue(dados, "Data"),

                    // === CAMPOS REMOVIDOS (retornar vazio para compatibilidade) ===
                    "empresa" or "company" or "organizacao" or "organization" => "",
                    "cargo" or "position" or "funcao" or "job" => "",
                    "datanascimento" or "nascimento" or "birthdate" or "dateofbirth" => "",
                    "observacoes" or "comments" or "notas" or "notes" => "",

                    _ => ""
                };
            }
            catch
            {
                return "";
            }
        }

        private static string GetPropertyValue(dynamic obj, string propertyName)
        {
            try
            {
                var property = obj.GetType().GetProperty(propertyName);
                return property?.GetValue(obj)?.ToString() ?? "";
            }
            catch
            {
                return "";
            }
        }
    }
}