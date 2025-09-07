# WordAutoFill - Preenchimento Automático de Documentos Word

Este projeto permite preencher automaticamente campos editáveis em documentos Microsoft Word (.docx) utilizando C# e a biblioteca OpenXML.

## 📝 Funcionalidades

O programa é capaz de identificar e preencher os seguintes tipos de campos em documentos Word:

- **Content Controls** (Controles de Conteúdo) - Tags e Alias
- **Form Fields** (Campos de Formulário Legacy)  
- **Bookmarks** (Marcadores)
- **Text Placeholders** (Placeholders de texto como `{{Campo}}`)
- **Merge Fields** (Campos de Mala Direta)

## 🏗️ Arquitetura do Projeto

O projeto está organizado em módulos separados para melhor manutenibilidade:

### Descrição dos Módulos

| Módulo | Responsabilidade |
|--------|-----------------|
| `Program.cs` | Coordena a execução do programa |
| `Core/DocumentFieldAnalyzer.cs` | Identifica campos editáveis no documento |
| `Core/DocumentFieldFiller.cs` | Preenche os campos com dados |
| `Core/FieldValueMapper.cs` | Mapeia nomes de campos para valores de dados |
| `Models/DocumentData.cs` | Fornece dados de exemplo para teste |
| `Utilities/DocumentFieldHelper.cs` | Funções utilitárias (localizar solution, extrair texto) |

## 🚀 Como Executar

### Pré-requisitos

- **.NET 8 SDK** ou superior
- **Visual Studio 2022** (recomendado) ou qualquer editor C#
- Documento Word (`.docx`) com campos editáveis

### Instalação

1. **Clone ou baixe o projeto**
2. **Instale o pacote NuGet necessário:**

3. **Coloque seu documento Word na pasta da solution:**
- O arquivo deve se chamar `Exemplo_Campos_Editaveis.docx`
- Deve estar na mesma pasta do arquivo `.sln`

### Executando o Programa

#### Via Visual Studio
1. Abra o projeto no Visual Studio
2. Pressione `F5` ou clique em "Iniciar"

#### Via Command Line

## 📝 Preparando o Documento Word

### 🔧 Habilitando a Aba Desenvolvedor no Word

**Antes de criar campos editáveis, você precisa habilitar a aba Desenvolvedor:**

#### Word 365/2021/2019:
1. Abra o Microsoft Word
2. Vá em **Arquivo** > **Opções**
3. Na janela que abrir, clique em **Personalizar Faixa de Opções**
4. No lado direito, marque a caixa **📋 Desenvolvedor**
5. Clique em **OK**

#### Word 2016:
1. Abra o Microsoft Word
2. Clique na aba **Arquivo**
3. Selecione **Opções** no menu lateral
4. Clique em **Personalizar Faixa de Opções**
5. Marque **📋 Desenvolvedor** na lista à direita
6. Clique em **OK**

**✅ Agora você terá a aba "Desenvolvedor" disponível na faixa de opções!**

### 1. Content Controls (Recomendado) ⭐

1. **Habilite a aba Desenvolvedor** (conforme instruções acima)
2. No Word, vá em **Desenvolvedor** > **Controles**
3. Clique em **Texto Rico** (ícone Aa)
4. Selecione o controle inserido e clique em **Propriedades**
5. Configure:
   - **Título**: Nome descritivo (ex: "Nome do Cliente")
   - **Tag**: Nome técnico (ex: `nome`, `email`, `telefone`)
   - ✅ Marque "O controle não pode ser excluído"
   - ✅ Marque "O conteúdo não pode ser editado"

**🔍 Exemplo prático:**
- Tag: `nome` → Será preenchido com "João Silva Santos"
- Tag: `valorconcedido` → Será preenchido com "R$ 150.000,00"

### 2. Bookmarks 🔖
1. Selecione o texto onde quer o campo
2. Vá em **Inserir** > **Links** > **Indicador**
3. Digite o nome (ex: `Nome`, `Email`, `ValorConcedido`)
4. Clique em **Adicionar**

### 3. Text Placeholders 📝
Simplesmente digite no documento:

### 4. Merge Fields
1. Vá em **Correspondências** > **Inserir Campo de Mesclagem**
2. Digite o nome do campo

## 📋 Campos Suportados

O programa reconhece automaticamente os seguintes nomes de campos organizados por categoria:

### 👤 Identificação das Partes
- `nome`, `name`, `nomecompleto`, `fullname`
- `cpf`, `documento`, `document`
- `endereco`, `address`
- `cidade`, `city`
- `uf`, `estado`, `unidadefederativa`
- `cep`, `postalcode`, `zipcode`
- `telefone`, `phone`, `celular`, `telephone`
- `email`, `emailaddress`, `e-mail`

### 💰 Valor de Crédito
- `valorconcedido`, `valor_concedido`, `valoremprestimo`, `credito`
- `prazopagamento`, `prazo_pagamento`, `prazo`
- `quantidadeparcelas`, `quantidade_parcelas`, `parcelas`, `numeroparcelas`
- `valorparcela`, `valor_parcela`, `parcela`

### 📊 Taxa de Juros e Encargos
- `jurosmes`, `juros_mes`, `jurosmensal`, `taxames`
- `jurosano`, `juros_ano`, `jurosanual`, `taxaano`
- `outrosencargos`, `outros_encargos`, `encargos`, `taxas`

### 🛡️ Garantias
- `garantias`, `garantia`, `caucao`

### 🏛️ Foro
- `local`, `localidade`, `foro`, `jurisdicao`
- `data`, `data
