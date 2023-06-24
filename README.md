# web-scraping-aege
## Contexto:
Uma das tarefas que me foi dada logo no inicio do estagio foi o preenchimento de cadastro do site de leilao de energia no site https://aege-empreendedor.epe.gov.br , esse site dava trabalho pra preencher pois ele vive travando , caindo e continha muitas informacoes a serem preenchidas.A empresa em que eu atuava colocava as informacoes a serem preenchidas pra cada empreendimento numa planilha padrao. E eu havia ouvido falar de um amigo na faculdade sobre um metodo de clique automatico usando selenium em python, a qual fiquei curioso pra testar entao testei nesse site o selenium pra web scraping. 

# Implementacao
Foi implementado de forma que ele preenche-se as informacoes do site sozinho , podendo ser retomando do ponto de parada . Por exemplo suponha que o site tenha caido na metade do processo , o codigo continuaria na metade do processo com entrada do usuario.
Com esse metodo o usuario podia entrar em duas , tres contas diferentes e poder preencher o site com mais confiabilidade e mais rapido do que preenchendo manualmente.Sendo necessario para ele apenas superviosonar quando o site caia e entao reiniciar o codigo. 

# Possiveis melhorias

-Treinar tempo de deteccao de carregar itens na pagina 

-Achar metodo de como fazer com que mesmo a pagina caindo o codigo abra uma nova guia e entre no site continuando exatamente de onde paro.Possivelemte salvando os passos num arquivo.

-Implementacao de screenshots para capturar o resultado para verificacao posterior de membros da equipe
