DOCUMENTAÇÃO DO BOT INADIMPLENCIA.

Pra que serve?

I -  O codigo em si, serve para verificar se a pasta BASE foi atualizada, caso nao tenha sido atualizada, apresenta erro e pede para atualizar a mesma.
II - Caso contrario, o codigo verifica quais pastas do diretorio estao presentes no arquivo .env e se estiverem, quais tem emails definidos.
III - Caso tenham emails definidos, ele abre a pasta, abre a planilha, atualiza os dados dela, salva em um dataframe, envia pelo email definido, 
exclui o arquivo do dataframe e fecha o excel.

IV - O codigo tambem gera tres relatorios:
-> O primeiro relatorio é gerado na planilha excel do google onde voce pode ver a data de execução e tempo de execução.
-> O segundo relatorio é um log de erro das pastas em geral.
-> E o terceiro é um log de erro da pasta GERAL.


Como o codigo funciona?

1) Deve se verificar se o caminho, emails de destinatario, email e senha de remetente do arquivo .env estão corretos.
2) E então ver se a planilha base foi atualizada com sucesso.

OBS* lembrando que a planilha base atualiza as 9:20 e o codigo é executado automaticamente as 9:30 - logo apos a planilha base ser atualizada.
