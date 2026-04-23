# Base Global - Render Ready

## O que faz
- Remove login de admin
- Qualquer pessoa pode subir uma nova planilha base
- A planilha ativa substitui a anterior
- Todo mundo vê a mesma base imediatamente
- Mostra resumo das abas e prévia das linhas

## Formatos aceitos
- `.xlsx`
- `.xlsm`

## Observação importante sobre o Render
Para teste, pode subir no plano Free.

Para **manter a base salva depois de reinício ou novo deploy**, use um **Persistent Disk** no Render e monte em `/var/data`.

Sem disco persistente, a base pode sumir após restart/redeploy.

## Rodar localmente
```bash
py -3.11 -m venv .venv
.venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt
python app.py
```

## Deploy
O projeto já inclui:
- `requirements.txt`
- `Procfile`
- `runtime.txt`
- `render.yaml`
