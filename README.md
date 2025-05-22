Compilador de Artigos - WPF
Esta aplicação WPF permite compilar múltiplos artigos em Word (.docx) ou texto (.txt) em um único documento formatado, incluindo lista de autores, índice automático, editorial e todos os artigos organizados.
Funcionalidades Implementadas
1. Lista de Autores Automática

Extrai automaticamente informações dos autores dos documentos
Procura por nomes, emails, IDs e instituições nos primeiros parágrafos
Remove duplicatas e organiza por artigo
Cria uma página dedicada com todos os autores listados

2. Índice Automático do Word

Gera um campo TOC (Table of Contents) nativo do Word
Permite atualização automática quando o documento é aberto no Word
Inclui títulos de seções automaticamente numerados

3. Editorial Personalizado

Interface dedicada para inserção do conteúdo editorial
Formatação automática do texto do editorial
Posicionamento estratégico após o índice

4. Compilação de Artigos

Suporte para arquivos .docx e .txt
Ordem personalizável via drag & drop
Formatação automática com estilos consistentes
Numeração automática dos artigos

Como Usar
Passo 1: Editorial

Ao iniciar a aplicação, uma janela de Editorial será exibida
Digite o conteúdo do editorial no campo de texto
Clique em "Continuar" para prosseguir ou "Cancelar" para sair

Passo 2: Seleção de Artigos

Na janela principal, adicione arquivos usando:

Botão "Adicionar Ficheiro"
Drag & Drop de arquivos do Windows Explorer


Reordene os arquivos conforme necessário:

Drag & Drop dentro da lista
Botões "Mover para Cima/Baixo"


Remova arquivos indesejados com "Remover Ficheiro"

Passo 3: Compilação

Clique no botão "Compilar"
Escolha o local e nome do arquivo final
Aguarde o processamento

Estrutura do Documento Final
O documento compilado terá a seguinte estrutura:

Página de Autores

Lista completa de todos os autores
Organizada por artigo
Inclui emails, instituições e IDs quando disponíveis


Índice Automático

Campo TOC nativo do Word
Atualiza automaticamente os números de página
Inclui todos os títulos e seções


Editorial

Conteúdo editorial inserido pelo usuário
Formatação profissional


Artigos

Cada artigo em uma página separada
Numeração automática (ARTIGO 1, ARTIGO 2, etc.)
Preservação da formatação original
Títulos e seções bem estruturados



Formatos Suportados

Entrada: .docx (Word), .txt (Texto)
Saída: .docx (Word)

Requisitos Técnicos

.NET 8.0 ou superior
Windows 10/11
Microsoft Word (recomendado para melhor visualização do índice)

Dependências

DocumentFormat.OpenXml - Manipulação de documentos Word
GongSolutions.Wpf.DragDrop - Funcionalidade drag & drop
Newtonsoft.Json - Serialização de dados

Dicas de Uso
Para Melhor Extração de Autores:

Mantenha informações de autores nos primeiros parágrafos
Use formato: "Nome - Email - Instituição"
Evite formatação excessiva nos dados dos autores

Para Índice Automático:

Após abrir o documento no Word, clique com botão direito no índice
Selecione "Atualizar campo" → "Atualizar página inteira"
Isso garantirá numeração de páginas correta

Para Melhor Formatação:

Use títulos claros e consistentes nos artigos
Evite formatação manual excessiva
Deixe a aplicação aplicar os estilos automaticamente

Solução de Problemas
Erro na Compilação:

Verifique se todos os arquivos existem e não estão abertos
Certifique-se de ter permissões de escrita no local de destino
Tente com menos arquivos para identificar problemas específicos

Autores Não Detectados:

Verifique se as informações estão nos primeiros parágrafos
Use formato simples: "Nome - Email - Escola"
Evite formatação complexa nos dados dos autores

Índice Não Atualizado:

Abra o documento no Microsoft Word
Clique com botão direito no índice e selecione "Atualizar campo"

Versão
Versão 2.0 - Implementação completa com todas as funcionalidades solicitadas.
