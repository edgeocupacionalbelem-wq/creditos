# Dashboard de Créditos com Cloudinary - Otimizado

Melhorias desta versão:
- leitura de planilha com read_only=True
- keep_links=False
- timeout do gunicorn aumentado para 120s
- limite de 10MB para planilhas muito pesadas
- comprovantes mantidos por recibo + pix_id
- até 2 comprovantes por crédito


## Versão com download direto
- Os dados da planilha ficam no PostgreSQL, então ninguém precisa reenviar planilha ao abrir o site.
- Os comprovantes ficam no Cloudinary, então PDFs e imagens continuam salvos após deploy/restart.
- O botão "Baixar" faz o download pelo próprio site.
- PDFs são enviados para o Cloudinary como `raw`, evitando erro de abrir PDF como imagem.
- Para PDFs antigos enviados antes desta correção, apague e envie novamente.


## Versão profissional
- visual mais profissional
- filtros por nome do mês, status, empresa e busca geral
- status visual dos comprovantes: Sem / Parcial / Completo
- log de atualização
- mantém dados no PostgreSQL e comprovantes no Cloudinary

## Ajuste de usabilidade
- filtro por intervalo removido da tela
- status simplificado: Sem comprovante / Com comprovante
- qualquer comprovante anexado já deixa o crédito verde
- mostra o nome do arquivo selecionado antes de anexar
