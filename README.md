# Controle de Empresas — planilha de obrigações contábeis

Sistema de acompanhamento mensal de obrigações contábeis, feito sobre
**Google Sheets + Google Apps Script**. Controla, por empresa e por mês,
as etapas de Folha, SPED, Faturamento e Consultas, com painel de
acompanhamento, ficha por empresa e bloco de notas.

## Arquivos do repositório

| Arquivo | O que é |
|---|---|
| `Codigo.gs` | Todo o script do Apps Script (menu, cadastro, ficha da empresa, notas…). |
| `MANUAL_CONTROLE_EMPRESAS.md` | Manual completo: instalação, rotina, solução de problemas. |
| `CONTROLE_EMPRESAS_MODELO.xlsx` | **Modelo vazio** da planilha (estrutura, fórmulas e formatação, **sem dados de cliente**). |
| `gerar_modelo.py` | Gera o modelo limpo a partir de um arquivo real e **recusa** se sobrar dado. |
| `.gitignore` | Barra qualquer `.xlsx` real — só o modelo limpo pode subir. |

## ⚠️ Regra de ouro (dados de cliente)

A planilha real contém **CNPJ, inscrições e SENHAS de clientes**. Isso
**nunca** vai para o repositório. Só sobe o `CONTROLE_EMPRESAS_MODELO.xlsx`,
que é gerado e **conferido** por `gerar_modelo.py` (o script apaga tudo e se
recusa a gerar se sobrar qualquer dado).

## Como atualizar o modelo

1. No Google Sheets: **Arquivo → Fazer download → Microsoft Excel (.xlsx)**.
2. Salve o arquivo nesta pasta.
3. Rode:
   ```
   py gerar_modelo.py "nome_do_arquivo_baixado.xlsx"
   ```
   Ele cria o `CONTROLE_EMPRESAS_MODELO.xlsx` limpo e confere que não sobrou dado.
4. Confirme o commit apenas do modelo (o arquivo real fica de fora pelo `.gitignore`).

## Instalação da planilha

Veja o **[manual](MANUAL_CONTROLE_EMPRESAS.md)**, seção 2.
