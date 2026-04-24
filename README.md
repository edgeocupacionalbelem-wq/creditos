# Dashboard de Créditos com Cloudinary - Otimizado

Melhorias desta versão:
- leitura de planilha com read_only=True
- keep_links=False
- timeout do gunicorn aumentado para 120s
- limite de 10MB para planilhas muito pesadas
- comprovantes mantidos por recibo + pix_id
- até 2 comprovantes por crédito


## Correção de PDF no Cloudinary
Nesta versão, PDFs são enviados como `resource_type="raw"` e imagens como `resource_type="image"`.
Isso evita o erro do navegador: "Falha ao carregar documento PDF" quando o arquivo PDF era entregue pela rota `/image/upload`.
