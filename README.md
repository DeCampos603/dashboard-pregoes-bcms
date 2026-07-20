# Dashboard Capacidade de Empenho — Atas BCMS

Painel web estático da **capacidade de empenho** (Qtd. Saldo × Valor Unitário
das atas de registro de preços) da UASG 160329 (BCMS), gerado a partir do
consolidado do **Robô Extrator de Pregões** e publicado no **GitHub Pages** com
atualização diária via **GitHub Actions** — mesmo padrão do
[dashboard-credito-bcms](https://github.com/DeCampos603/dashboard-credito-bcms).

## Como funciona

1. O robô extrator gera `Consolidado_Pregoes.xlsx` (nome fixo, sobrescrito no
   lugar) em `RoboExtratorPregoes_V2/Consultas realizadas/`, que o Google Drive
   sincroniza. Como o arquivo é sobrescrito (não recriado), **o ID/link no
   Drive nunca muda**.
2. `gerar_dashboard.py` baixa esse xlsx pelo link público, filtra as linhas da
   UASG alvo, calcula a capacidade por status da ata (vigente / ≤30 dias /
   vencida) e gera `site/index.html` autocontido (tema claro/escuro, busca,
   ordenação, drill por categoria/pregão, alerta de vencimentos, histórico).
3. O workflow roda todo dia às 08h30 (BRT), commita o histórico e publica no
   Pages.

## Publicar (uma vez)

1. **Compartilhar o consolidado:** no Drive, ache `Consolidado_Pregoes.xlsx`
   → Compartilhar → "Qualquer pessoa com o link". Copie o link e extraia o ID
   (o trecho entre `/d/` e `/view`).
2. **Criar o repositório** (ex.: `dashboard-pregoes-bcms`) e enviar esta pasta:
   ```
   cd Dashboard-Pregoes-BCMS
   git init -b main
   git add . && git commit -m "Dashboard capacidade de empenho v1"
   git remote add origin https://github.com/DeCampos603/dashboard-pregoes-bcms.git
   git push -u origin main
   ```
3. **Configurar o repo no GitHub:**
   - Settings → Secrets and variables → Actions → New secret:
     `DRIVE_FILE_ID` = o ID copiado no passo 1.
   - Settings → Pages → Source: **GitHub Actions**.
4. **Primeira publicação:** Actions → "Atualizar Dashboard Pregões" →
   Run workflow. Depois disso roda sozinho todo dia.

## Testar local

```bat
set SOURCE_XLSX=..\RoboExtratorPregoes_V2\Consultas realizadas\Consolidado_Pregoes.xlsx
py -3 gerar_dashboard.py
```
Abra `site/index.html` no navegador.

## Avisos

- ⚠️ O Pages é **público**: o painel (valores de saldo por item) e o
  `data/history.json` ficam visíveis para quem tiver o link.
- ⚠️ Push **não** dispara deploy (o workflow roda por cron/manual) — para ver
  mudanças na hora, use Actions → Run workflow.
- 🐛 Se o git quebrar com `bad object refs/desktop.ini` (Google Drive injeta
  `desktop.ini` no `.git`): `Get-ChildItem .git -Recurse -Force -Filter
  desktop.ini | Remove-Item -Force` — ou mova o repo para fora da pasta
  sincronizada.
- `UASG_ALVO` (variável de ambiente) muda a unidade alvo sem editar o código.
