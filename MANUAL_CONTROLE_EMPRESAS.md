# Manual — Controle de Empresas

Guia de instalação, rotina mensal e solução de problemas da planilha de controle de obrigações.

**Arquivos:**
- `CONTROLE_EMPRESAS_2026.xlsx` — a planilha
- `Codigo.gs` — o script que automatiza o trabalho

---

## 1. Como a planilha é organizada

A ideia central: **o que é fixo fica separado do que muda todo mês.**

| Aba | Para que serve |
|---|---|
| **PAINEL** | Tela de acompanhamento. Escolha o mês e veja o que falta em cada etapa. À direita fica o **calendário de vencimentos** do mês. |
| **CADASTRO** | Dados fixos da empresa (CNPJ, regime, senha) e em quais etapas ela entra. Preenchido uma vez. |
| **1. FOLHA** | Folhas e encargos. Só empresas com FAZ FOLHA = Sim. |
| **2. SPED** | SPED ICMS/IPI/Contribuições. Só empresas com FAZ SPED = Sim. |
| **3. FATURAMENTO** | Faturamento do mês. Todas as empresas ativas. |
| **4. CONSULTAS** | Consultas fiscais. Todas as empresas ativas. |
| **COMPARATIVO** | Faturamento de cada empresa mês a mês + pendências por mês. |
| **VENCIMENTOS** | Prazos das obrigações que alimentam o calendário do PAINEL, e a lista de categorias/cores. Você lança à mão. (Seção 8) |
| **COMO USAR** | Resumo rápido dentro da própria planilha. |
| **LISTAS** | Bastidores (as listas suspensas). Fica oculta — não mexa. |
| **NOTAS** | Guarda as anotações do bloco de notas. Fica oculta — o script cuida dela. (Seção 9) |

Nas abas de etapa, a **coluna MÊS** é o que guarda o histórico: os meses ficam empilhados, nada é apagado.

**A planilha é de um ano.** Em janeiro, use *Virar o ano* (seção 10) para criar o arquivo do ano seguinte.

---

## 2. Instalação (uma vez só)

1. Suba o `.xlsx` para o Google Drive.
2. Abra o arquivo e vá em **Arquivo → Salvar como Planilhas Google**.
   *Obrigatório.* Se ficar como .xlsx, o script não funciona.
3. Na planilha convertida: **Extensões → Apps Script**.
4. Apague o `function myFunction() {}` que vem lá (Ctrl+A → Delete) e cole o conteúdo do `Codigo.gs`. Salve no 💾.
5. No seletor do topo, escolha a função `onOpen` e clique em **Executar**.
6. O Google vai pedir permissão. Aparece "app não verificado" — é normal, o script é seu:
   **Avançado → Acessar (não seguro) → Permitir**.
7. Volte para a planilha e **recarregue a página (F5)**.

Pronto: aparece o menu **🧮 Modo Contador** no topo.

> É tudo num arquivo só. Não precisa criar arquivos HTML: as janelas são geradas pelo próprio script.

---

## 3. Começando do zero

1. No **PAINEL**, escolha o mês em que você vai começar (a célula azul).
   *Importante: o script cadastra a empresa no mês que estiver selecionado aqui.*
   Ao abrir a planilha, esse seletor já vem no **mês atual** — mas você troca quando quiser.
2. **🧮 Modo Contador → Cadastrar empresa nova**. Abre um formulário com tudo na mesma tela:

| Campo | Observação |
|---|---|
| Nome da empresa | É a chave que liga tudo. Escolha com cuidado. |
| CNPJ | — |
| Regime | Simples Nacional, Simples Híbrido, Lucro Presumido ou MEI |
| **Entra a partir de** | Em qual mês ela começa. Veja abaixo. |
| Faz a FOLHA? | Define se ela aparece na aba 1. FOLHA |
| Faz o SPED? | Define se ela aparece na aba 2. SPED |

3. Clique em **Cadastrar**. O formulário **continua aberto e já limpo** para a próxima empresa — dá para cadastrar várias seguidas sem reabrir nada. O Enter no campo do nome também cadastra.

O script grava no CADASTRO (marcando ATIVA? = Sim), insere a empresa nos meses escolhidos e no COMPARATIVO.

**Senha, perfil e inscrições** não são perguntados: preencha direto no CADASTRO depois.

### O campo "Entra a partir de"

A lista mostra só os **meses já abertos** na planilha. A empresa entra no mês escolhido **e em todos os meses abertos depois dele**.

| Você escolhe | O que acontece (planilha com JAN a JUL abertos) |
|---|---|
| **JULHO** (o padrão, mês mais recente) | Entra só em julho. É o caso normal: cliente novo que chegou agora. |
| **JANEIRO** | Entra em janeiro, fevereiro... até julho. Para lançar retroativo o ano todo. |
| **MAIO** | Entra em maio, junho e julho. |
| **— Não adicionar a nenhum mês —** | Fica só no cadastro. Entra quando você usar *Abrir novo mês*. |

Se nenhum mês foi aberto ainda (planilha nova), a empresa fica só no cadastro — é o esperado. Cadastre todas e depois use *Abrir novo mês*.

---

## 4. Rotina mensal

1. **🧮 Modo Contador → Abrir novo mês** → escolha o mês na lista → **Abrir mês**.

   O script monta o mês novo em todas as abas com as empresas certas (só as **ativas**, e respeitando quem faz folha e quem faz SPED), deixa os status em branco e muda o PAINEL para o mês novo. **Os meses anteriores continuam intactos.**

2. Trabalhe nas abas na ordem: Folha → SPED → Faturamento → Consultas.
   Use o filtro da coluna MÊS para ver só o mês atual.

3. Marque os status pela lista suspensa (a cor aparece sozinha).

4. Acompanhe pelo **PAINEL**: ele mostra quantas empresas faltam em cada etapa.

> O script não deixa duplicar mês. Se você abrir ABRIL duas vezes, ele avisa "já existia, pulei".

---

## 5. Status — o que cada um significa

| Status | Quando usar |
|---|---|
| **Concluído** | Feito e conferido |
| **Pendente** | Ainda a fazer / declaração sem enviar |
| **Agendada** | Notas agendadas (SPED) |
| **Baixada** | Notas baixadas (SPED) |
| **Importada** | Notas importadas (SPED) |
| **Sem movimento** | Empresa sem movimento no mês |
| **Não se aplica** | Não se aplica àquela empresa |
| **Retificada** | Declaração retificada |
| **Erro** | Travou / precisa de atenção |

O que conta como **pendência** no PAINEL e no COMPARATIVO: *Pendente* + *Erro*.

A coluna que sinaliza que a etapa acabou é a **CONFERÊNCIA** (no Faturamento, é **ENVIADOS**).

---

## 6. Mudanças no cadastro

### O que pode editar na mão

Tudo, **menos o nome da empresa**: CNPJ, regime, FAZ FOLHA?, FAZ SPED?, ATIVA?, perfil, inscrições, senha.

### O que precisa ser pelo menu

| Situação | O que fazer |
|---|---|
| **Trocar o nome da empresa** | 🧮 Modo Contador → **Renomear empresa**. Escolha a empresa na lista e digite o novo nome. Troca em tudo de uma vez (cadastro, 4 abas, comparativo). |
| **Empresa saiu do escritório** | 🧮 Modo Contador → **Ativar / desativar empresa**. Escolha na lista; ela para de entrar nos meses novos e o histórico fica guardado. **Não apague.** |
| **Empresa voltou** | O mesmo menu — ele reativa. |
| **Cadastrei errado, quero sumir com ela** | 🧮 Modo Contador → **Excluir empresa de vez**. Escolha na lista, veja quantas linhas vão embora e confirme. Não tem volta. |
| **Abri o mês errado** | 🧮 Modo Contador → **Excluir um mês inteiro**. Veja a seção 10.1. |
| **Conferir se algum faturamento está estranho** | 🧮 Modo Contador → **Conferir faturamentos suspeitos**. Veja a seção 11.1. |
| **Ver / lançar prazos das obrigações** | 🧮 Modo Contador → **Mostrar / atualizar calendário**. Veja a seção 8. |
| **Anotar lembretes e avisos** | 🧮 Modo Contador → **Bloco de notas** (ou clique no ícone 📝 no painel). Veja a seção 9. |
| **Cadastrar empresa** | 🧮 Modo Contador → **Cadastrar empresa nova**. |

### Por que o nome não pode ser editado na mão

As abas de etapa buscam o CNPJ e o regime **pelo nome**, e o COMPARATIVO soma o faturamento **pelo nome**. Se você trocar só no CADASTRO, as linhas antigas ficam órfãs: o CNPJ some e o faturamento daquela empresa zera no comparativo.

### Três armadilhas

- **Cadastrar empresa direto no CADASTRO, na mão, funciona pela metade**: ela entra nos meses novos, mas **não ganha a linha dela no COMPARATIVO**. Cadastre sempre pelo menu.
- **FAZ FOLHA? e FAZ SPED? só valem para os meses seguintes.** Mudar para "Sim" hoje não faz a empresa aparecer nos meses que já estão abertos — ela entra no próximo "Abrir novo mês". Para incluir num mês já aberto, digite o mês e o nome numa linha vazia da aba (CNPJ e regime aparecem sozinhos).
- **Campo ATIVA? em branco conta como ativa**, para ninguém sumir por descuido.

---

## 7. Acompanhar e comparar

**PAINEL** — escolha o mês na célula azul (ao abrir, já vem no mês atual). Mostra, por etapa: quantas empresas entram naquele mês, quantas estão concluídas, quantas estão pendentes ou com erro, e a barra de progresso. À direita fica o **calendário de vencimentos** do mês (seção 8).

**COMPARATIVO** — duas tabelas:

- *Faturamento por empresa*: cada empresa nas 12 colunas de mês, com total do ano. Dá para ver na hora se uma empresa faturou mais ou menos que no mês passado. A linha TOTAL GERAL puxa direto do Faturamento.
- *Pendências por mês*: quantas declarações ficaram sem concluir em cada etapa, em cada mês. **Vermelho** = tem pendência, **verde** = tudo enviado.

---

## 8. Calendário de vencimentos

No **PAINEL**, à direita (colunas H a N), fica um calendário do mês selecionado com uma **bolinha colorida** no dia de cada obrigação. Ele segue o mês do seletor e destaca o **dia de hoje**.

**Para montar/atualizar:** 🧮 Modo Contador → **Mostrar / atualizar calendário**. Na primeira vez, isso cria a aba **VENCIMENTOS** e desenha o calendário. Depois ele se atualiza sozinho quando você troca o mês ou lança um prazo.

### Lançar os prazos

Os prazos mudam de mês para mês, então são **lançados à mão** na aba **VENCIMENTOS** (colunas A–D):

| Coluna | O que é |
|---|---|
| MÊS | Mês do vencimento (lista suspensa) |
| DIA | Dia do mês (1 a 31) |
| OBRIGAÇÃO | Nome livre (ex.: "Simples", "DCTFWeb") |
| CATEGORIA | A cor da bolinha (lista suspensa) |

Um mesmo dia pode ter vários vencimentos — aparecem várias bolinhas juntas.

### Categorias e cores

Na mesma aba VENCIMENTOS, ao lado (colunas **F e G**), fica a lista de categorias e a bolinha (cor) de cada uma. **Você edita à vontade:** renomeia, adiciona ou troca a cor — o dropdown de CATEGORIA e a legenda do calendário se atualizam sozinhos. As cores saem de uma paleta de bolinhas de emoji (🔴 🟠 🟡 🟢 🔵 🟣 🟤 ⚫ ⚪), então dá para ter até 9 categorias com cores distintas.

> O calendário é "pintado" com valores fixos, sem fórmula viva — **não pesa** na planilha.

---

## 9. Bloco de notas

Um bloco de anotações que abre numa **janela flutuante**. Serve para lembretes soltos, avisos e tarefas.

**Para abrir:** 🧮 Modo Contador → **Bloco de notas**, ou clique no **ícone 📝** no painel (veja abaixo como criar o ícone).

Na janela:
- **+ Nova Nota** cria uma anotação. Você escolhe o **tipo**, a **empresa** (opcional), a **prioridade**, a **data** e a **descrição**.
- Cada nota tem **✎ editar**, **✕ apagar** e, nas tarefas, a caixinha **concluída**.
- Os **filtros** no topo: Todas · Pendentes · Avisos · Concluídas.

### Os três tipos de nota

| Tipo | Para que serve |
|---|---|
| **Tarefa – Pendente** | Algo a fazer. Tem prioridade (🔴 Alta, 🟠 Média, 🔵 Baixa) e a caixinha "concluída". |
| **Tarefa – Concluída** | Uma tarefa marcada como feita (fica riscada, em verde). |
| **Aviso** | 🟣 Um recado que não é tarefa — não fica pendente nem se conclui. |

As anotações ficam guardadas na aba **NOTAS** (oculta) — o script cuida dela, não precisa mexer.

### Criar o ícone 📝 no painel (uma vez)

O menu já abre o bloco. Para ter também um ícone clicável no painel:

1. **Inserir → Desenho**. Escreva um emoji (📝) ou "Notas", **Salvar e fechar**.
2. Arraste o desenho para um canto do PAINEL.
3. Clique no desenho → **⋮ (três pontinhos)** → **Atribuir script** (pode aparecer como *"Transferir script"*).
4. Digite exatamente `abrirBlocoDeNotas` → **OK**.

Pronto: clicar no ícone abre o bloco. Na primeira vez, o Google pede autorização — é só permitir.

---

## 10. Virar o ano

### 10.1 Abri o mês errado — como apagar

**🧮 Modo Contador → Excluir um mês inteiro.** Escolha o mês e clique em **Verificar**: antes de apagar nada, o script mostra quantas linhas existem em cada aba, quantos status já estão marcados e quanto de faturamento foi lançado.

A confirmação muda conforme o risco:

| Situação | O que o script pede |
|---|---|
| **Mês sem trabalho** (recém-aberto, tudo Pendente) | Um clique em *Apagar definitivamente*. |
| **Mês com trabalho lançado** | Alerta vermelho e você precisa **digitar o nome do mês** para confirmar. |

Depois de apagar, o mês some das 4 abas e você pode abri-lo de novo. Se o PAINEL estava naquele mês, ele pula para o último mês que sobrou.

> **Não tem desfazer no script** — mas tem no Google: **Arquivo → Histórico de versões** permite voltar a planilha ao estado anterior. É a sua rede de segurança.

### Virar o ano

A coluna MÊS não guarda o ano — por isso **cada ano tem seu arquivo**.

Em janeiro: **🧮 Modo Contador → Virar o ano** → digite o ano novo (ex.: `2027`).

O que ele faz:

- Cria uma **cópia** chamada `CONTROLE EMPRESAS 2027`.
- Na cópia, **mantém** o CADASTRO e as empresas do COMPARATIVO.
- Na cópia, **zera** as abas de etapa, prontas para o ano novo.
- **Não altera o arquivo atual**, que fica como histórico do ano que passou.

No fim ele mostra o link do arquivo novo. Abra ele e use *Abrir novo mês → JANEIRO*.

> O script vai junto na cópia. Não precisa colar nada de novo.

---

## 11. Alerta de pendências por e-mail (opcional)

Não tem botão no menu — funciona por acionador automático:

1. **Extensões → Apps Script**.
2. Clique no **ícone de relógio** (Acionadores), no menu da esquerda.
3. **Adicionar acionador**:
   - Função: `enviarPendencias`
   - Origem do evento: **Baseado no tempo**
   - Tipo: por exemplo, **Semanal**, segunda de manhã.
4. Salve.

Você recebe um e-mail com as pendências do mês que estiver selecionado no PAINEL, listando as empresas que faltam em cada etapa, mais o link da planilha.

> O e-mail vai para quem criou o acionador. Se outra pessoa quiser receber, ela precisa criar o acionador dela.

---

## 11.1 Conferir faturamentos suspeitos

**🧮 Modo Contador → Conferir faturamentos suspeitos.** Abre uma janela que lista os valores de faturamento que fogem muito do padrão de cada empresa — provável **erro de digitação** (um zero a mais ou a menos) ou **variação real** que vale conferir.

**Como funciona:** para cada empresa, o script calcula a média dos meses dela e compara cada mês com essa média. Se o valor está **±50% ou mais** fora da média, ele aparece na lista. Só analisa empresas com **3 meses ou mais** de faturamento lançado (com menos que isso a média não é confiável).

Na janela: 🔴 **vermelho** = muito acima da média (possível zero a mais); 🟡 **amarelo** = muito abaixo (possível zero a menos ou queda real). É só um *"vale conferir"* — **não altera nada na planilha**. Confira o valor na aba 3. FATURAMENTO.

> No começo do ano ele encontra pouco — precisa de histórico. Fica mais útil conforme os meses avançam.

### Como ajustar a sensibilidade

Dá para mudar a "régua" no topo da seção 9 do `Codigo.gs`, editando **duas linhas**:

```js
const ANOMALIA_LIMITE = 0.50;   // ±50% da média dispara o alerta
const ANOMALIA_MIN_MESES = 3;   // só analisa empresas com 3+ meses lançados
```

| Quer... | Mude `ANOMALIA_LIMITE` para |
|---|---|
| **Pegar mais coisa** (mais sensível, mais alarmes) | `0.30` (±30%) |
| **Equilíbrio** (padrão) | `0.50` (±50%) |
| **Só casos gritantes** (dobrou ou caiu pela metade) | `1.00` (±100%) |

E `ANOMALIA_MIN_MESES`: use `2` para começar a analisar mais cedo no ano, ou deixe `3` (recomendado) para menos alarme falso. Depois de editar, salve o script (💾) e recarregue a planilha (F5).

---

## 12. Solução de problemas

| Problema | Causa e solução |
|---|---|
| **O menu 🧮 Modo Contador não aparece** | Recarregue a página (F5). Se ainda não aparecer, confira se o script foi salvo e se a planilha foi convertida para Planilhas Google. |
| **"Exception: A operação não é aceita em um intervalo com uma linha filtrada"** | Tem filtro escondendo linhas. Limpe os critérios do filtro (funil → Selecionar tudo) e rode de novo. |
| **Abri o mês e ele não aparece no filtro** | O filtro tem faixa fixa e o mês novo caiu fora dela. O script já corrige isso sozinho ao abrir o mês. Se acontecer, reaplique o filtro (Dados → Criar um filtro sobre toda a área). |
| **O CNPJ ficou em branco numa linha** | O nome na coluna EMPRESA não bate com nenhum nome do CADASTRO. Quase sempre é nome editado na mão. Use Renomear empresa. |
| **Erro nas fórmulas depois de mexer no script** | O script usa **ponto e vírgula** (`;`) nas fórmulas, padrão brasileiro. Se sua conta Google estiver em inglês, troque por vírgula (`,`). |
| **A janela abre em branco** | Recarregue a planilha (F5) e tente de novo. Se persistir, confira se o script foi colado inteiro. |
| **Quero saber onde o script travou** | Na mensagem de erro, clique em **Detalhes**: ele mostra a linha exata. |

---

## 13. Limites e manutenção

- **Limite de 1.500 linhas por aba** (`LIMITE_LINHAS` no topo do script). Cada empresa ocupa 12 linhas por aba (uma por mês), e os dados começam na linha 3 — então cabem **124 empresas** com o ano completo nas abas onde todas entram (3. FATURAMENTO e 4. CONSULTAS). As abas 1. FOLHA e 2. SPED têm mais folga, porque só recebem parte das empresas.

  O limite não avisa aos poucos: ele estoura ao **abrir dezembro**, com o ano inteiro já lançado. Se estiver perto de 124, aumente o número antes de virar o ano.

  Para ir além, aumente `LIMITE_LINHAS` **e** estenda as listas suspensas e as cores condicionais até a nova linha. **Atenção:** as fórmulas do COMPARATIVO gravam o limite no próprio texto no momento em que a empresa é cadastrada. As empresas que já estão lá continuam com `$H$1500` e param de somar as linhas acima disso **sem dar erro** — depois de aumentar o limite, recadastre as fórmulas do COMPARATIVO.
- **Histórico de versões**: Arquivo → Histórico de versões. Dá para ver quem mudou o quê e voltar atrás. É a sua rede de segurança.
- **Proteja as fórmulas**: botão direito na aba → Proteger intervalo. Sugestão: colunas CNPJ, REGIME e TOTAL, e as abas PAINEL, COMPARATIVO e LISTAS.

### Trocar o nome do menu

O nome do menu (o que aparece no topo, ao lado de *Extensões*) fica em **uma linha** do `Codigo.gs`, dentro da função `onOpen`:

```js
.createMenu('🧮 Modo Contador')
```

Para trocar: **Extensões → Apps Script**, ache essa linha e mude o texto entre as aspas. O emoji na frente é opcional. Depois salve (💾) e recarregue a planilha (F5).

---

## 14. Compartilhamento

- O script **vai junto** com a planilha. Quem receber vê o menu 🧮 Modo Contador.
- Só quem tem acesso de **Editor** consegue usar as funções (elas escrevem na planilha). Como *Leitor*, a pessoa vê os dados mas não usa o menu.
- **Cada pessoa autoriza por conta própria** na primeira vez, com aquela mesma tela de "app não verificado". Avise antes.
- Quem é Editor também **consegue ver e alterar o código**. Não dá para trancar.
- **Atenção à coluna SENHA (J do CADASTRO)**: quem tem acesso à planilha vê essa coluna. Se for compartilhar, considere ocultar a coluna ou proteger o intervalo — ou manter as senhas num gerenciador de senhas separado.

---

## 15. Resumo de um mês típico

```
1. 🧮 Modo Contador → Abrir novo mês → MAIO
2. Aba 1. FOLHA       → filtra MAIO → marca os status
3. Aba 2. SPED        → filtra MAIO → agenda / baixa / importa → marca
4. Aba 3. FATURAMENTO → filtra MAIO → lança valores → marca ENVIADOS
5. Aba 4. CONSULTAS   → filtra MAIO → marca
6. PAINEL             → confere se ficou tudo verde
7. COMPARATIVO        → compara o faturamento com os meses anteriores
```

**Entrou empresa nova no meio do caminho?**
🧮 Modo Contador → Cadastrar empresa nova, e escolha em "Entra a partir de" o mês certo
(o mais recente para cliente novo; um mês anterior para lançar retroativo).

**Empresa saiu?**
🧮 Modo Contador → Ativar / desativar empresa. Nunca apague.

**Virou o ano?**
🧮 Modo Contador → Virar o ano. O arquivo antigo fica como histórico.
