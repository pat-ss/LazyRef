# LazyRef

Este projeto foi feito para quem não é fã de fazer as referências bibliográficas no final dos trabalhos/documentos. Basta inserir as referências num ficheiro Excel, deicar o código analisar e wham bam, temos referências no final do trabalho, organizadas e documentadas segundo as normas APA.

(sim sim, dá trabalho inserir as referências todas manualmente na mesma mas, a meu ver, através de um ficheiro de Excel é muito mais fácil e organizado - e menos desformatado)

Requisitos:

- python -> https://docs.python.org/3.6/using/windows.html#installing-python
- openpyxl -> ir à linha de comandos > escrever "pip install openpyxl" (sem aspas)
- docx -> ir à linha de comandos > escrever "pip install python-docx" (sem aspas)

Features:

- Se a data não for indicada, o programa escreve (s.d.);
- Se o local da publicação não for indicado, o programa escreve [s.l.] - que corresponde às iniciais da expressão latina sine loco e significa “sem lugar”;
- Se o nome da editora não constar na publicação, o programa escreve [s.n.] - que corresponde às iniciais da expressão latina sine nomine e significa “sem nome”;
- A primeira edição não é suposto ser mencionada e indica-se a edição apenas a partir da 2ª, ou seja, se escreveres "1ª" (desta forma específica, se for "1ª edição" já não), o programa automaticamente altera para que não apareça nada, tendo em conta que não é suposto mencionar a primeira edição;
- Formatação de urls de forma e com cor correta;
- Títulos a itálico, quando assim têm que ser;
- Bibliografia sai justificada;
- As referências são ordenadas alfabeticamente;
- No caso de existirem trabalhos do mesmo autor, o mais antigo aparece em primeiro lugar;
- No caso de existirem trabalhos do mesmo autor e com o mesmo ano, as referências são ordenadas alfabeticamente pelo título e é adicionado uma letra ao ano, para os "separar", de certa forma (ex.: duas obras de Saramago de 2008 vêm como 2008a e 2008b);
- Os trabalhos de um autor precedem trabalhos de múltiplos autores que iniciam pelo mesmo apelido;
- Trabalhos do mesmo primeiro autor e com segundos e terceiros autores diferentes, são organizados por ordem alfabética do segundo autor.

Regras:

- Estes ficheiros têm que estar todos inseridos na pasta do documento onde queremos adicionar as referências;
- Existe um ficheiro Excel, "exemplos_refs" que serve como template. O user deve preencher o ficheiro "refs" com as suas próprias referências, e tem o "exemplos_refs" para perceber como é que os campos devem ser preenchidos;
- Em todas as linhas usadas, é necessário inserir um "." (ponto final) no campo "validação", ou aquela linha não vai ser analisada;
- No caso de ser necessário escrever números ordinais ("1ª edição", "2º volume", etc.), é pedido que o user escreve só "1º", "2º", etc, porque o código já vai inserir o resto automaticamente;
- Antes de correr o programa, é necessário fechar por completo tanto o documento Word utilizado como o ficheiro Excel que contém as referências a inserir;
- O nome dos autores pode ser escrito por inteiro ou só a primeira letra do nome (excepto apelidos, onde é necessário o nome inteiro);
- O código não contempla legislação estrangeira, apenas portuguesa;
- No campo do número de autores, não é necessário repetir em todas as linhas se houver mais que um autor. Por exemplo, um livro tem 3 autores - no primeiro, é necessário preencher tudo (nº de autores, nomes, título, etc), mas nos outros dois só é preciso preencher os nomes (e a validação, claro).

Instruções:

- Preencher o ficheiro Excel "refs";
- Fechar o documento Excel e o documento Word onde se vai inserir as referências;
- Abrir a linha de comandos (símbolo do windows > "cmd" ou "command line" ou "linha de comandos";
- Escrever "cd" (sem aspas) e depois o caminho das pastas até à pasta onde está o documento Word - ex.: cd Dektop\trabalho - e carregar no enter;
- Escrever "python exe.py" (sem aspas) e carregar no enter;
- Escrever o nome do documento Word - ex.: trabalho.docx - e carregar no enter.
