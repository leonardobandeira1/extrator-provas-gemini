# Extrator de Provas com Google Sheets e Gemini API

Este projeto é uma aplicação para Google Sheets que utiliza a API Gemini (Google AI) para extrair automaticamente nomes de alunos e médias de imagens de folhas de resposta de provas.

## Funcionalidades

- Processamento de imagens de folhas de resposta de provas diretamente do Google Drive
- Extração automática do nome do aluno e da média usando IA
- Inserção dos dados extraídos em uma planilha do Google Sheets
- Suporte para processamento em lote de múltiplas imagens

## Requisitos

- Conta Google com acesso ao Google Sheets e Google Drive
- Chave de API do Google Gemini
- Imagens de folhas de resposta armazenadas no Google Drive

## Como usar

1. Copie o código do script para o editor do Google Apps Script em sua planilha
2. Substitua `'SUA_CHAVE_API_AQUI'` pela sua chave API real do Gemini
3. Salve o script e recarregue sua planilha
4. Use o menu "Extrator de Provas" que aparecerá para processar suas imagens

## Obtendo uma chave API do Gemini

1. Acesse o [Google AI Studio]
2. Faça login com sua conta Google
3. Crie uma nova chave API
4. Copie a chave e use-a no script

## Limitações

- Apenas imagens JPEG e PNG são suportadas no momento
- Os arquivos PDF ainda não são processados automaticamente
- A API Gemini tem limites de uso, consulte a documentação para mais detalhes

## Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para abrir um issue ou enviar um pull request.
