'''
Ambiente Repositorio: Onedrive
python 3.8.5 x64
VS Code 
Chrome 89 (atualizar manualmente)
Webdriver para o chrome 89
Disponibilizados no repositorio: https://tpfecombr-my.sharepoint.com/:f:/g/personal/paulo_silva_tpfe_com_br1/EoHQqtEVQC1NohTHW7IsfYMBtnzxDsGvf_0RQ6HxGPBt0w?e=V3TKa7
'''

'''
Padronizacao do codigo:


1- Buscar todos os ID com o status de "Solicitado" (Há variações no nome para cada tipo de solicitação financeira) 
2- Buscar todos aqueles que possuem DATA SOLICITADA PARA PAGAMENTO seguindo a regra: 
Dia de hoje (dia que o robo está fazendo a tramitação) até o dia 15 do mês. Caso o dia 15 já tenha passado ou não haja mais solicitações para o dia 15, ele faz o Dia de hoje até o último dia do mês,
'''