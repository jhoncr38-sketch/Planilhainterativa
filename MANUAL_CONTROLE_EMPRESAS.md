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
| **PAINEL** | Tela de acompanhamento. Escolha o mês e veja o que falta em cada etapa. |
| **CADASTRO** | Dados fixos da empresa (CNPJ, regime, senha) e em quais etapas ela entra. Preenchido uma vez. |
| **1. FOLHA** | Folhas e encargos. Só empresas com FAZ FOLHA = Sim. |
| **2. SPED** | SPED ICMS/IPI/Contribuições. Empresas com FAZ SPED = Sim ou FAZ EFD CONTRIB.? = Sim. |
| **3. FATURAMENTO** | Faturamento do mês. Todas as empresas ativas. |
| **4. CONSULTAS** | Consultas fiscais. Todas as empresas ativas. A coluna **OBSERVAÇÃO** é texto livre para anotar o que você achou na consulta (débito, certidão positiva, pendência a cobrar...). |
| **COMPARATIVO** | Pendências por mês, por etapa. O faturamento de cada empresa fica na **Ficha da empresa**. |
| **COMO USAR** | Resumo rápido dentro da própria planilha. |
| **LISTAS** | Bastidores (as listas suspensas). Fica oculta — não mexa. |
| **NOTAS** | Guarda as anotações do bloco de notas. Fica oculta — o script cuida dela. (Seção 8) |

Nas abas de etapa, a **coluna MÊS** é o que guarda o histórico: os meses ficam empilhados, nada é apagado.

**A planilha é de um ano.** Em janeiro, use *Virar o ano* (seção 9) para criar o arquivo do ano seguinte.

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
   Ao abrir a planilha, esse seletor já vem no **mês de trabalho**: o mês atual, se ele já foi aberto; senão, o último mês aberto (ex.: em outubro, ainda lançando março → MARÇO). Você troca quando quiser.
2. **🧮 Modo Contador → Cadastrar empresa nova**. Abre um formulário com tudo na mesma tela:

| Campo | Observação |
|---|---|
| Nome da empresa | É a chave que liga tudo. Escolha com cuidado. |
| CNPJ | — |
| Regime | Simples Nacional, Simples Híbrido, Lucro Presumido ou MEI |
| **Entra a partir de** | Em qual mês ela começa. Veja abaixo. |
| Faz a FOLHA? | Define se ela aparece na aba 1. FOLHA |
| Faz o SPED? | SPED fiscal (EFD ICMS/IPI). Com Sim, ela aparece na aba 2. SPED |
| Faz EFD Contrib.? | EFD Contribuições. Já vem marcado pelo regime (Lucro Presumido = Sim), dá para trocar. Com Sim, ela também aparece na aba 2. SPED |

3. Clique em **Cadastrar**. O formulário **continua aberto e já limpo** para a próxima empresa — dá para cadastrar várias seguidas sem reabrir nada. O Enter no campo do nome também cadastra.

O script grava no CADASTRO (marcando ATIVA? = Sim), e insere a empresa nos meses escolhidos.

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

   O script monta o mês novo em todas as abas com as empresas certas (só as **ativas**, e respeitando quem faz folha e quem faz SPED), deixa os status em branco — menos o "Não se aplica" automático da aba 2. SPED, veja abaixo — e muda o PAINEL para o mês novo. **Os meses anteriores continuam intactos.**

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

Na coluna final de cada etapa, **Concluído**, **Sem movimento**, **Não se aplica** e **Retificada** contam como **resolvido**. Todo o resto — *Pendente*, *Erro* e **em branco** — conta como **em aberto**. A regra é a mesma no PAINEL, no COMPARATIVO, na Ficha da empresa e no e-mail de pendências.

A coluna que sinaliza que a etapa acabou é a **CONFERÊNCIA** (no Faturamento, é **ENVIADOS**).

---

## 6. Mudanças no cadastro

### O que pode editar na mão

Tudo, **menos o nome da empresa**: CNPJ, regime, FAZ FOLHA?, FAZ SPED?, FAZ EFD CONTRIB.?, ATIVA?, perfil, inscrições, senha.

### FAZ SPED? e FAZ EFD CONTRIB.? — o "Não se aplica" automático

Na aba 2. SPED, ao abrir um mês, a EFD que a empresa **não faz** já vem marcada como **Não se aplica**:

| FAZ SPED? | FAZ EFD CONTRIB.? | Na aba 2. SPED |
|---|---|---|
| Sim | Sim | Entra; as duas EFDs ficam em branco para você marcar |
| Sim | Não | Entra; **EFD CONTRIBUIÇÕES = Não se aplica** (caso do Simples Nacional) |
| Não | Sim | Entra; **EFD ICMS/IPI = Não se aplica** (empresa que faz só a EFD Contribuições) |
| Não | Não | Não entra na aba 2. SPED |
| Sim | (em branco) | Entra; nada marcado — como era antes da coluna existir |

Da coluna D em diante, o CADASTRO é lido pelo **cabeçalho** (linha 2) — EMPRESA, CNPJ e REGIME ficam sempre em A, B e C. Pode mudar a ordem das outras colunas, mas **não renomeie os cabeçalhos** (ex.: "ATIVA?", "FAZ SPED?"): se o script não achar um deles, ele avisa e para, em vez de ler a coluna errada.

### O que precisa ser pelo menu

| Situação | O que fazer |
|---|---|
| **Trocar o nome da empresa** | 🧮 Modo Contador → **Renomear empresa**. Escolha a empresa na lista e digite o novo nome. Troca em tudo de uma vez (cadastro e as 4 abas). |
| **Empresa saiu do escritório** | 🧮 Modo Contador → **Ativar / desativar empresa**. Escolha na lista; ela para de entrar nos meses novos e o histórico fica guardado. **Não apague.** |
| **Empresa voltou** | O mesmo menu — ele reativa. |
| **Cadastrei errado, quero sumir com ela** | 🧮 Modo Contador → **Excluir empresa de vez**. Escolha na lista, veja quantas linhas vão embora e confirme. Não tem volta. |
| **Abri o mês errado** | 🧮 Modo Contador → **Excluir um mês inteiro**. Veja a seção 9.1. |
| **A empresa passou a fazer SPED/folha (ou voltou) e o mês já está aberto** | Mude no CADASTRO e depois 🧮 Modo Contador → **Incluir empresa num mês aberto**. Escolha a empresa e o mês: a janela mostra em quais abas ela vai entrar e em quais já está, e você confirma. Vale também para os meses abertos depois dele. |
| **Ver o ano inteiro de uma empresa** | 🧮 Modo Contador → **Ficha da empresa**. Veja a seção 7. |
| **Conferir se algum faturamento está estranho** | 🧮 Modo Contador → **Conferir faturamentos suspeitos**. Veja a seção 10.1. |
| **Anotar lembretes e avisos** | 🧮 Modo Contador → **Bloco de notas** (ou clique no ícone 📝 no painel). Veja a seção 8. |
| **Cadastrar empresa** | 🧮 Modo Contador → **Cadastrar empresa nova**. |

### Por que o nome não pode ser editado na mão

As abas de etapa buscam o CNPJ e o regime **pelo nome**, e a Ficha da empresa junta o histórico **pelo nome**. Se você trocar só no CADASTRO, as linhas antigas ficam órfãs: o CNPJ some e o histórico daquela empresa some da ficha.

### Três armadilhas

- **Cadastrar empresa direto no CADASTRO, na mão, funciona pela metade**: ela só entra a partir do próximo mês que você abrir — **não entra nos meses que já estão abertos**. Cadastre sempre pelo menu.
- **FAZ FOLHA?, FAZ SPED? e FAZ EFD CONTRIB.? só valem para os meses seguintes.** Mudar para "Sim" hoje não faz a empresa aparecer nos meses que já estão abertos — ela entra no próximo "Abrir novo mês". Para incluir num mês já aberto, use 🧮 Modo Contador → **Incluir empresa num mês aberto** (se o cursor estiver na linha dela no CADASTRO, ela já vem escolhida).
- **Campo ATIVA? em branco conta como ativa**, para ninguém sumir por descuido.

---

## 7. Acompanhar e comparar

**PAINEL** — escolha o mês na célula azul (ao abrir, já vem no mês de trabalho). Mostra, por etapa: quantas empresas entram naquele mês, quantas estão **concluídas** (Concluído, Sem movimento, Não se aplica ou Retificada), quantas estão **em aberto** (Pendente, Erro ou em branco) e o progresso — que chega a 100% quando tudo está resolvido.

**COMPARATIVO** — *Pendências por mês*: quantas declarações estão em aberto (Pendente, Erro ou em branco) em cada etapa, em cada mês. **Vermelho** = tem pendência, **verde** = tudo resolvido. Para ver o faturamento de uma empresa mês a mês, use a **Ficha da empresa** (abaixo).

**FICHA DA EMPRESA** — 🧮 Modo Contador → **Ficha da empresa**. Mostra o ano inteiro de **uma** empresa numa janela só, sem precisar filtrar as 4 abas:

- **Cadastro**: CNPJ, regime, perfil, IE, IM, se faz Folha/SPED/EFD Contribuições e se está ativa. A **senha não aparece** na ficha, de propósito.
- **Quatro números**: quantas etapas estão em aberto no ano — separado em pendentes, erros e em branco —, faturamento do ano, média por mês e notas pendentes.
- **Status por mês**: uma linha por etapa e uma coluna por mês, com as mesmas cores da planilha. Cada linha mostra a **coluna final** da etapa (a mesma que o PAINEL usa). **Clique no nome da etapa** para ver as outras colunas dela. Status **em branco** num mês aberto aparece como uma caixinha tracejada amarela: é o que falta preencher. Na última linha fica o faturamento total de cada mês (passe o mouse para ver o valor exato). Meses ainda não abertos aparecem com o cabeçalho cinza.
- **Observações**: o que foi anotado na coluna OBSERVAÇÃO da aba 4. CONSULTAS, mês a mês. Na grade, a linha OBSERVAÇÃO (dentro de Consultas) mostra 💬 nos meses com texto — passe o mouse para ler.
- **Notas da empresa**: as notas do Bloco de notas ligadas a ela (pendentes e avisos; as concluídas ficam em "ver / esconder").

**Passar de empresa:** as setas **◀ ▶** ao lado da lista (ou as teclas ← → do teclado) vão para a empresa anterior/seguinte, em ordem alfabética.

**Atalho:** se o cursor estiver na linha de uma empresa (nas abas de etapa ou no CADASTRO) quando você abrir a ficha, ela já abre nessa empresa. A ficha só **lê**: não altera nada na planilha.

---

## 8. Bloco de notas

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

## 9. Virar o ano

### 9.1 Abri o mês errado — como apagar

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
- Na cópia, **mantém** o CADASTRO.
- Na cópia, **zera** as abas de etapa, prontas para o ano novo.
- **Não altera o arquivo atual**, que fica como histórico do ano que passou.

No fim ele mostra o link do arquivo novo. Abra ele e use *Abrir novo mês → JANEIRO*.

> O script vai junto na cópia. Não precisa colar nada de novo.

---

## 10. Alerta de pendências por e-mail (opcional)

Não tem botão no menu — funciona por acionador automático:

1. **Extensões → Apps Script**.
2. Clique no **ícone de relógio** (Acionadores), no menu da esquerda.
3. **Adicionar acionador**:
   - Função: `enviarPendencias`
   - Origem do evento: **Baseado no tempo**
   - Tipo: por exemplo, **Semanal**, segunda de manhã.
4. Salve.

Você recebe um e-mail com as pendências do **mês de trabalho** (o mês atual, se já foi aberto; senão, o último mês aberto), listando as empresas que faltam em cada etapa, mais o link da planilha.

> O e-mail vai para quem criou o acionador. Se outra pessoa quiser receber, ela precisa criar o acionador dela.

---

## 10.1 Conferir faturamentos suspeitos

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

## 11. Solução de problemas

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

## 12. Limites e manutenção

- **Limite de 1.500 linhas por aba** (`LIMITE_LINHAS` no topo do script). Cada empresa ocupa 12 linhas por aba (uma por mês), e os dados começam na linha 3 — então cabem **124 empresas** com o ano completo nas abas onde todas entram (3. FATURAMENTO e 4. CONSULTAS). As abas 1. FOLHA e 2. SPED têm mais folga, porque só recebem parte das empresas.

  O limite não avisa aos poucos: ele estoura ao **abrir dezembro**, com o ano inteiro já lançado. Se estiver perto de 124, aumente o número antes de virar o ano.

  Para ir além, aumente `LIMITE_LINHAS` **e** estenda as listas suspensas e as cores condicionais até a nova linha. **Atenção:** as fórmulas do PAINEL e do COMPARATIVO (PENDÊNCIAS) têm a faixa escrita no próprio texto (ex.: `$H$3:$H$1376`). Estenda essas faixas também — senão elas param de somar as linhas de baixo **sem dar erro**.
- **Histórico de versões**: Arquivo → Histórico de versões. Dá para ver quem mudou o quê e voltar atrás. É a sua rede de segurança.
- **Proteja as fórmulas**: botão direito na aba → Proteger intervalo. Sugestão: colunas CNPJ, REGIME e TOTAL, e as abas PAINEL, COMPARATIVO e LISTAS.

### Trocar o nome do menu

O nome do menu (o que aparece no topo, ao lado de *Extensões*) fica em **uma linha** do `Codigo.gs`, dentro da função `onOpen`:

```js
.createMenu('🧮 Modo Contador')
```

Para trocar: **Extensões → Apps Script**, ache essa linha e mude o texto entre as aspas. O emoji na frente é opcional. Depois salve (💾) e recarregue a planilha (F5).

---

## 13. Compartilhamento

- O script **vai junto** com a planilha. Quem receber vê o menu 🧮 Modo Contador.
- Só quem tem acesso de **Editor** consegue usar as funções (elas escrevem na planilha). Como *Leitor*, a pessoa vê os dados mas não usa o menu.
- **Cada pessoa autoriza por conta própria** na primeira vez, com aquela mesma tela de "app não verificado". Avise antes.
- Quem é Editor também **consegue ver e alterar o código**. Não dá para trancar.
- **Atenção à coluna SENHA do CADASTRO**: quem tem acesso à planilha vê essa coluna. Se for compartilhar, considere ocultar a coluna ou proteger o intervalo — ou manter as senhas num gerenciador de senhas separado.

---

## 14. Resumo de um mês típico

```
1. 🧮 Modo Contador → Abrir novo mês → MAIO
2. Aba 1. FOLHA       → filtra MAIO → marca os status
3. Aba 2. SPED        → filtra MAIO → agenda / baixa / importa → marca
4. Aba 3. FATURAMENTO → filtra MAIO → lança valores → marca ENVIADOS
5. Aba 4. CONSULTAS   → filtra MAIO → marca (e anota em OBSERVAÇÃO o que precisar)
6. PAINEL             → confere se ficou tudo verde
7. COMPARATIVO        → confere as pendências de cada mês
```

**Entrou empresa nova no meio do caminho?**
🧮 Modo Contador → Cadastrar empresa nova, e escolha em "Entra a partir de" o mês certo
(o mais recente para cliente novo; um mês anterior para lançar retroativo).

**Empresa saiu?**
🧮 Modo Contador → Ativar / desativar empresa. Nunca apague.

**Virou o ano?**
🧮 Modo Contador → Virar o ano. O arquivo antigo fica como histórico.
