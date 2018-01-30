# Processador de pendências

Este programa lê as planilhas geradas pelos diferentes bancos com as pendencias de contratos e propostas de emprestimos.

### Utilização

Para iniciar o tela gráfica desta aplicação basta iniciar o aplicativo executável que deve estar presente com estes arquivos. Não é necessário fazer o que é descrito na próxima seção.

Caso o aplicativo executável não esteja presente, será necessário instalar o ruby e as gems utilizadas pela aplicação seguindo o passo a passo da próxima seção (Instalação).
Com estes instalados, basta clicar duas vezes no arquivo .rbw para que a tela principal apareça.

### Instalação

1. Instalar o ruby na máquina. Recomendo utilizar o [ruby installer] (https://rubyinstaller.org/downloads/) na versão 2.4.X 64 bits.

2. Abrir o terminal e executar os seguintes comandos na pasta que contém estes arquivos:

'''bash
gem install bundler
bundle install
'''
  
### Geração do aplicativo executável

Caso seja necessário, o usuário pode criar um arquivo executável basta utilizar a gem ocra.

1. Realizar o passo a passo de instalação desta aplicação como descrito na seção anterior (Instalação)
2. Executar o comando 'ocra main.rbw' na pasta que contem estes arquivos utilizando o terminal. A tela principal da aplicação deve abrir
3. Para que o programa ocra encontre as dependencias da aplicação, é necessário processar uma planilha xls, xlsx e clicar no botão de "copiar para a área de transferência". A janela não deve fechar sozinha, senão será necessário recomeçar o procedimento.
4. Por último, basta fechar a tela do programa e esperar o arquivo .exe ser criado.