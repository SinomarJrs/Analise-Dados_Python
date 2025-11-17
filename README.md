# Projeto para leitura de arquivo contendo dados, filtragem e exportação

    O arquivo analise.py lê o arquvo que o usuário selecionar e busca informações nele para finalizar a classificação dos filtros.
    Após executará a análise dos dados mediantes os critérios pré-estabelecidos e exportará os resultados para a pasta pré-determinada ou de escolha do usuário.


# Bibliotecas necessárias para rodar via editor de texto:
    - Utilizar o comando no terminal para instalar as bibliotecas via arquivo.txt

pip install -r requeriments.txt


# Criar um executável do código que possa ser utilizados por outros usuários:

- pip install pyinstaller
- Instalar direto no diretório onde o projeto está alocado:
    D:/Projetos/Programa_Tesouraria/.venv/Scripts/python.exe -m pip install pyinstaller
- Gerar um executável com todas as dependências:
D:/Projetos/Programa_Tesouraria/.venv/Scripts/pyinstaller.exe --onefile --windowed --name "Acerto em Atraso" --icon "Imagens/icon.ico" --add-data "Imagens;Imagens" analise.py
- Comando para limpar os arquivos temporários e recriar o executável:
rm -rf build dist analise.spec
- Comando para criar o arquivo requeriments.txt com todas as dependências do projeto que deverão ser instaladas para rodar via terminal:
pip freeze > requirements.txt
- Comando para instalar todas as dependências do projeto:
pip instalç -r requirements.txt

# Estrutura de pastas:
    - Para que o usuário não precise ficar selecionado o arquivo ou a pasta de exportação, deverá ser mantida essa estrutura de pastas
    - Executar o comando abaixo via terminal. *OBS*-> Se atentar em qual disco de armazenagem as pastas serão criadas
    - Mover o executável para a pasta Programa_Tesouraria
    
- mkdir Programa_tesouraria\Arquivos-Analise
- cp analise.exe para o diretório Programa_Tesouraria
- cp DADOS.xls para o diretório programa-tesouraria
