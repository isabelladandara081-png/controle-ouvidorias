# controle-ouvidorias

Aplicativo em Python (Streamlit) para o acompanhamento diário de ouvidorias: painel de prioridades,
pendentes, controle de dilatadas, histórico de respondidas, importação das bases e backup.

## Como executar

```bash
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

Os dados ficam em `dados/ouvidorias.json` (ou na pasta indicada pela variável `OUVIDORIAS_DADOS`).

## Fluxo

```
Pendentes -> justificativa da dilação -> Controle de Dilatadas -> acompanhamento e cobrança
          -> importação da base de Respondidas -> Histórico de respondidas
```

| Situação | Significado |
|---|---|
| **Dilatada** | Justificativa preenchida: dilação solicitada. Sai das pendentes automaticamente. |
| **Aguardando devolutiva** | Dilação solicitada e aguardando retorno da área (aparece com o nome do analista em destaque). |
| **Vencida** | Prazo final ultrapassado (vale também para dilatadas). |
| **Unidade vencida** | Sem dilação e com o prazo da unidade ultrapassado. |
| **Respondida** | Código encontrado na planilha de respondidas. É a única forma de concluir. |

Regras de prazo: sem dilação vale o prazo da unidade; com dilação, **só o prazo final** determina o vencimento.
Apagar a justificativa devolve a ouvidoria às pendentes.

## Importações (aba "Importar bases")

1. **Aguardando resposta (diária):** entram códigos novos, os existentes são atualizados sem apagar o que foi preenchido.
2. **Ouvidorias respondidas:** conclui as pendentes/dilatadas encontradas, preservando UC, observações, justificativa, OE e fiscal.
   Lê apenas abas que tenham uma coluna de código (`CODIGO`, `Código`, `Protocolo`…); se existir coluna com "resposta" no nome, usa como data.
3. **Acompanhamento:** vincula UC, OE, equipe, fiscal e observações pelo código. Linhas amarelas ou com "dilação" entram como dilatadas.

Os códigos são normalizados (`267.038.655.920` -> `267038655920`), então não há duplicidade.
Antes de cada importação é criado um ponto de restauração (botão "Desfazer última importação").

Para trazer os dados do aplicativo em HTML, use *Backup e exportação -> Trazer dados do aplicativo em HTML*.

## Testes

```bash
pytest
```

Os testes usam planilhas sintéticas geradas no próprio teste.

## Privacidade

**Não envie planilhas nem a pasta `dados/` ao GitHub** (o `.gitignore` já bloqueia `*.xlsx` e `dados/*`).
Use um repositório privado. Em hospedagens públicas (ex.: Streamlit Community Cloud) o app não tem login
e o disco é temporário; prefira executar localmente ou atrás de autenticação.

## Estrutura

- `controle.py` — regras, importações, armazenamento e exportação (sem interface).
- `app.py` — interface Streamlit.
- `tests/` — testes automatizados.
