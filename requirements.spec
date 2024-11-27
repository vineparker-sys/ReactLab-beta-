pandas
xlsxwriter
pyserial
rich
Pillow

A UI do Programa de Análise de Reatividade deve seguir as heuristicas de Nielsen, e deve ter um design moderno.

A janela deve ser feita em proporção que caiba em uma tela 1920x1080, e também telas 1280x720. Também deve suportar resoluções 4:3 além das resoluções 16:9. Ficando em tela cheia, sempre.
Na pasta assets, temos a logo da aplicação, logo.png, e a logo da empresa, lhoist_logo.png.

A logo da aplicação deve aparecer centralizada, na parte de cima da aplicação. E Abaixo dela, o texto: "Software de Análise de Reatividade"
A logo da Lhoist, deve aparecer no canto inferior direito da janela. Num tamanho menor, mas visível.

A tela inicial deverá conter menus como:

Arquivo
    - Exportar -> ultima análise (pensando numa funcionalidade futura de guardar os dados da ultima analise)
                -> relatório (pensando numa funcionalidade futura para exportação de várias análises passadas, puxando do banco de dados)
    - Preferencias -> Tela onde será possível editar o idioma do software (pensando em implementar multi-linguagem futuramente) e também o tema do software (tambem implementado futuramente), e também editar o layout de exportação (também implementado futuramente)
    - Fechar -> Fecha o programa


Ajuda
    - Sobre -> menu de informações do software e seu desenvolvimento
    - Suporte -> informações e botão que redirecionará para o outlook para envio de emails para vinicius.alves@lhoist.com para suporte do software


Na tela, abaixo da logo, teremos o texto "Para começar, selecione uma opção de tempo de análise."

E abaixo organizado de forma simétrica, teremos as 7 opções de tempo de análise.
A primeira, precisa estar no centro da grade 9x9, na primeira fileira, onde ao invés de ter 3 botões, terá apenas 1 botão, o de 30 segundos, centralizado.
abaixo os outros 6 botoes, 3 na fileira do meio, e 3 na fileira de baixo.

Ao selecionar uma opção, ao invés de criar uma nova janela livre, e apartada para inserção dos codigos das amostras baseados nos termopares ativos, quero que ele gere uma nova janela, menor, no centro da janela principal, e que fique travada ali, o usuário não poderá clicar na tela principal enquanto não fechar, ou colocar os codios da amostra para seguir com a analise.
Após inserir o código das amostras, e clicar para continuar, a janela menor para inserção dos codigos das amostras é fechada, e é aberta uma nova janela com um gráfico tempo x temperatura, onde o gráfico mostrará em tempo real a curva de reatividade da amostra.
Abaixo do gráfico, teremos uma barra de progresso indicando o progresso da análise, quanto tempo falta, e quanto tempo já foi, além do porcentual, assim como era na interface CLI.
Ao lado esquerdo, uma tabela, com apenas 2 colunas, coluna de tempo, e coluna de temperatura.
A coluna de tempo, assim como o export em excel, com as linhas tempo 0, -00:30, 1:00, e por ai vai até 35 minutos. Isso independente do tempo escolhido. Caso as analises escolhidas sejam menor que 35 minutos, as linhas após o tempo escolhido ficaram vazias.
A ideia aqui é que, conforme a analise for correndo, a tabela vai sendo preenchida visualmente nessa tela, assim o tecnico pode acompanhar.
Gostaria que ficasse um campo de texto abaixo do gráfico, Para que quando tivermos análises mais longas, como 30 e 35 minutos, essa caixa mostrar, em verde, a palavra "Estável" quando a temperatura da amostra estabilizar.
A logica para estabilização pode ser algo como, ao manter a mesma temperatura por determinado tempo, e depois cair, a temperatura em que a amostra estabilizou deverá ser usada para o calculo do DELTA T, ao invés da temperatura final. Assim, ao ver que a temperatura se manteve e não vai mais subir, apresentar o texto estável.
E também, abaixo da tabela, um campo de texto informando o Delta, após a conclusão da análise, usando a temperatura final, para analises mais curtas, e usando a temperatura em que a amostra estabilizou para as analises mais longas.
Após a análise, precisará aparecer dois botões em algum lugar na tela, de preferencia abaixo da barra de progresso, um botão para exportar para excel, e outro para realizar outra análise, e voltar a tela principal.
 
