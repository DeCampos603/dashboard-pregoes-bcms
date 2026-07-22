# Roadmap — Capacidade de Empenho para todas as unidades da Base

Documento de planejamento para evoluir o painel de **uma unidade (BCMS, UASG
160329)** para **todas as unidades da Base**. Escrito a partir dos dados reais
do consolidado de 20/07/2026 (43.680 linhas), não de suposições.

## ✅ Já implementado (2026-07-21) — o SELETOR de OMDS

O painel já tem o **framework multi-OMDS**: uma barra de botões no topo com o
**brasão de cada OM** (7 unidades: BCMS, 1º D Sup, BMSA, D C Mun, ECT, Cia C,
H Cmp). Ao clicar, o painel troca o **emblema do cabeçalho**, a **cor-tema**
(acento derivado do brasão, validado p/ contraste) e o **título**.

- **Só o BCMS tem dados** (é a UASG_ALVO coletada). As demais mostram o painel
  **"Aguardando coleta"** com o brasão e a cor da unidade — prontas para receber
  os dados quando forem coletadas.
- Manifesto em `gerar_dashboard.py` → `UNIDADES` (sigla, nome, UASG, logo,
  accent). Logos em `site/assets/logos/`. Cor de cada uma extraída do brasão
  com Pillow e escurecida até passar 4,5:1.
- Decisão de design: o **acento tinge só o "cromo"** (barra de topo, chip ativo,
  régua do herói, moldura). Gráficos e status ficam na paleta medida — os
  brasões são quase todos vermelhos, e vermelho já significa "Vencida".

**Falta para popular uma 2ª unidade:** coletá-la no robô (ver Fase 3) **e**
migrar os dados embutidos para `data/<uasg>.json` (Fase 2) — hoje o HTML embute
só o BCMS; com N unidades populadas, o embutido não escala.

---

## 1. A distinção que decide tudo

Existem **dois sentidos diferentes** de "unidade" nos dados, e confundi-los é o
maior risco deste projeto:

| | **UASG operando** | **UASG participante** |
|---|---|---|
| Onde vive | `lista_uasgs.txt` do robô | coluna `UASG` da planilha |
| O que é | a unidade em cujo *Gestão de Atas* o robô entra | cada unidade listada **dentro** de uma ata |
| Hoje | **1** (160329 BCMS) | **257** distintas |
| Exige | acesso do certificado àquela UASG | nada — já vem de brinde |

**Consequência prática:** o consolidado atual **já contém** a capacidade de 257
unidades participantes — **R$ 485.040.338,39** no total. Ou seja, boa parte da
expansão **não exige mexer no robô**: basta apontar o painel para outra UASG.

Mas atenção ao alcance: essas 257 unidades aparecem apenas nas atas em que o
**BCMS também participa**. Uma unidade só terá cobertura **completa** (todas as
suas atas) quando for incluída em `lista_uasgs.txt` — e isso depende de acesso.

### Números medidos (capacidade em atas válidas, 20/07/2026)

| UASG | Unidade | Capacidade | Itens |
|---|---|---:|---:|
| 160297 | CMDO 1ª DE | R$ 48.168.286,17 | 937 |
| **160238** | **BA AP LOG EX** | **R$ 37.294.510,57** | 1.017 |
| 160264 | 111ª CIA AP MB | R$ 30.890.608,30 | 208 |
| **160329** | **BCMS** *(painel atual)* | **R$ 23.586.517,53** | 1.503 |
| 160296 | BA ADM BDA INF PQDT | R$ 17.577.146,55 | 1.074 |
| 160304 | BMSA | R$ 5.120.940,73 | 1.216 |
| … | *(mais 251 unidades)* | | |

---

## 2. O que já está pronto para multiunidade

O gerador **já é parametrizado por ambiente** — nenhuma unidade está fixa no código:

```bash
UASG_ALVO=160238 NOME_UNIDADE="Base de Apoio Logístico do Exército" \
NOME_CURTO="Ba Ap Log Ex" py -3 gerar_dashboard.py
```

Isso gera **hoje**, sem nenhuma alteração de código, um painel completo da Base.
É o teste que valida a hipótese antes de investir na arquitetura maior.

---

## 3. Decisão de arquitetura

Três caminhos, com o trade-off honesto de cada um:

### Opção A — Um site, seletor de unidade *(recomendada)*
Um `index.html` com um dropdown de unidade no topo, ao lado dos filtros atuais.

- ✅ Um único link para divulgar; comparação entre unidades no mesmo lugar.
- ✅ Reaproveita 100% dos filtros e do recálculo já implementados.
- ⚠️ **Peso do arquivo é o risco real.** O painel de 1 unidade já tem 2,1 MB com
  1.752 linhas embutidas no HTML. 10 unidades ≈ 20 MB — inviável.
  **Mitigação obrigatória:** trocar as linhas embutidas por um `dados.json`
  carregado sob demanda (`fetch`) e renderizar a tabela em JS. Isso derruba o
  HTML para ~60 KB e o JSON pode ser dividido por unidade
  (`data/160329.json`), carregando só a unidade escolhida.

### Opção B — Uma página por unidade + índice
`site/160329/index.html`, `site/160238/index.html`, … e um índice com o ranking.

- ✅ Simples; sem refatorar a renderização; cada página continua leve.
- ✅ Funciona sem JavaScript.
- ⚠️ Comparar unidades exige trocar de página; N vezes mais HTML gerado.

### Opção C — Visão consolidada da Base + drill-down
Uma página "Base" (soma de todas) com drill para cada unidade.

- ✅ É o que a chefia normalmente quer ver: o total e quem tem o quê.
- ⚠️ **Cuidado com dupla contagem:** uma mesma ata lista várias unidades
  participantes; somar unidades é somar *saldos distintos* (cada linha é o saldo
  daquela UASG) — isso é legítimo, mas **nunca** somar por ata.

**Recomendação:** começar por **A com JSON externo**, que absorve C como um caso
de filtro ("Base inteira" = sem filtro de unidade).

---

## 4. Plano de execução

### Fase 1 — Validação (1 dia, sem código novo)
1. Definir com a chefia **quais UASGs compõem a Base** (lista oficial).
2. Rodar o gerador para 2–3 delas via variável de ambiente e conferir os números
   com quem opera cada unidade.
3. **Ponto de decisão:** os dados de participação bastam, ou é preciso cobertura
   completa (Fase 3)?

### Fase 2 — Refatorar para dados externos (pré-requisito da escala)
1. `gerar_dashboard.py` passa a emitir `data/<uasg>.json` (itens) +
   `data/unidades.json` (índice com nome, capacidade, contagem).
2. `index.html` vira casca (~60 KB): carrega o índice, monta o seletor, faz
   `fetch` do JSON da unidade escolhida e renderiza a tabela em JS.
3. Manter o **fallback sem JS**: um `<noscript>` com link para a versão por
   página (Opção B) da unidade padrão.
4. Regra a preservar: **rótulos vindos da planilha entram no DOM via
   `textContent`**, nunca `innerHTML` (dado não confiável).

### Fase 3 — Cobertura completa (depende de acesso)
1. Adicionar as UASGs da Base em `lista_uasgs.txt` (uma por linha, com apelido).
2. **Verificar o acesso do certificado**: o robô troca de UASG via
   `POST /alteraruasgusuario`, que só funciona para unidades que o usuário
   logado pode operar. Sem permissão, a Fase 1 do robô falha naquela UASG (o log
   já mostra o motivo: `[401/403] sem permissão para operar esta UASG`).
3. Reavaliar o tempo de execução: hoje ~1 unidade leva N minutos; medir antes de
   prometer prazo. A Fase 3 acelerada (fetch + lxml) ajuda, e há espaço para
   concorrência se necessário.

### Fase 4 — Visão de comando
1. Ranking de unidades por capacidade (barras — série única, mesma cor).
2. Alerta consolidado: o que a **Base inteira** perde nos próximos 30/60 dias.
3. Histórico por unidade (`data/history.json` passa a ter chave por UASG).

---

## 5. Riscos e cuidados

| Risco | Mitigação |
|---|---|
| **Peso do arquivo** (2,1 MB × N unidades) | Fase 2 é pré-requisito, não opcional |
| **Certificado sem acesso** a outras UASGs | Validar na Fase 1 antes de prometer |
| **Dupla contagem** ao somar unidades | Somar *linhas por UASG*, nunca por ata |
| **Exposição pública** dos dados | O Pages é público. Com a Base inteira, o volume exposto cresce muito — avaliar repositório privado + Pages privado (exige plano pago) ou publicação interna |
| **ND é sugestão automática** | Já sinalizado no rodapé; não usar como fonte oficial de classificação |
| Unidade sem dados vira painel vazio | O gerador já aborta com `[ERRO]` em vez de publicar vazio — manter esse comportamento por unidade |

---

## 6. Decisões em aberto (para o usuário)

1. **Quais UASGs compõem a Base?** (lista oficial — é o insumo que falta)
2. O certificado atual **tem acesso** para operar essas UASGs no Compras.gov.br?
3. O painel deve ficar **público** (como hoje) ou restrito?
4. A visão principal é **por unidade** ou o **total da Base** com drill?

---

## 7. Referências no código

| Onde | O quê |
|---|---|
| `gerar_dashboard.py` → `UASG_ALVO`, `NOME_UNIDADE`, `NOME_CURTO` | parametrização por unidade (já pronta) |
| `gerar_dashboard.py` → `etl()` | filtra por `UASG` começando com `UASG_ALVO` |
| `gerar_dashboard.py` → `_tabela()` | ponto onde as linhas viram HTML (a trocar por JSON na Fase 2) |
| `RoboExtratorPregoes_V2/lista_uasgs.txt` | UASGs *operando* (Fase 3) |
| `RoboExtratorPregoes_V2/robo/fase1_atas.py` → `trocar_uasg()` | troca de UASG e seus erros de permissão |
