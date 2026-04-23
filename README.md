# Créditos por Mês - Render Ready

## O que faz
- Remove login de admin
- Qualquer pessoa pode subir uma nova planilha base
- A planilha ativa substitui a anterior
- O sistema procura `NÃO REALIZADO` na coluna H ou D
- Mostra os créditos por mês
- Mostra:
  - número do recibo
  - empresa
  - CNPJ/CPF extraído do campo Setor (coluna L)
  - quantidade de créditos por recibo

## Regra de agrupamento
Se houver mais de um crédito no mesmo recibo, o sistema mostra apenas **uma linha** para o recibo, com a **quantidade de créditos**.

## Formatos aceitos
- `.xlsx`
- `.xlsm`

## Observação importante sobre o Render
Para teste, pode subir no plano Free.

Para manter a base salva depois de reinício ou novo deploy, use um **Persistent Disk** no Render e monte em `/var/data`.

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
