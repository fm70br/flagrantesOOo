# flagrantesOOo
OpenOffice.org Macros for mail merge documents

## AVISO! O texto a seguir também é antigo e pode estar desatualizado. s/OpenOffice/LibreOffice/g

## FLAGRANTES OpenOffice.org

Formulários, banco de dados e biblioteca de macros do OpenOffice.org para confecção automática de Inquéritos Policiais. Os modelos de documentos pertencem a "Flagrantes Jatinhos" para MS Word, cujo autor é o Escrivão Jorge. 

## RECURSOS

Um arquivo modelo para cada tipo de documento. Fácil e rápida edição / correção / adaptação dos modelos. Formulários separados para cada assunto/parte do documento. Acolhimento dos depoimentos em formulários. Possibilidade de registrar qualquer número de envolvidos (conduzidos / vítimas). Tratamento automatizado de palavras no singular / plural. Tratamento automatizado de listas de nomes (conduzidos / vítimas). Geração automática de Inquérito Policial com um único clique.


## INSTRUÇÕES PARA INSTALAÇÃO

Pré-requisitos: Java instalado e funcionando em conjunto com o OpenOffice.org.

Descompate o arquivo .zip em sua pasta de trabalho.  Preserve a estrutura dos diretórios descompactados, caso contrário as macros não vão funcionar.

Abra o OpenOffice e registe o Banco de Dados "FLAGRANTES.odb":
  - Menu "Ferramentas" - "Opcoes" - "OpenOffice.org BASE" - "Banco de Dados". Clique em "NOVO", escolha o arquivo "FLAGRANTES.odb" que está na mesma pasta do LeiaMe.odt (este documento).

Configure o diretório descompactado como "seguro" para execução de macros:
  - Menu "Ferramentas" - "Opções" - "OpenOffice.org" – "Segurança".
      
Clique em "SEGURANÇA DE MACRO", depois na segunda aba clique em "Adicionar" e selecione a pasta descompactada. O código fonte das macros pode ser acessado via "Alt+F11" ou no menu "Ferramentas" – "Macro". Observação: Esta configuração é opcional, caso queira, poderá confirmar a execução da macro sempre que necessário, basta ajustar as opções de “Segurança de Macro” para o nível “médio”. Se a macro não puder ser executada, não será possível gerar nenhum documento.
      
Após o registro do Banco de Dados e das Macros, o programa estará pronto para funcionar.

## INSTRUÇÕES DE USO

1. Abra o menu principal: arquivo "FlagrantesOpenOffice.odt" (para facilitar o acesso ao documento, crie um atalho para este arquivo em sua área de trabalho). Cada botão abre um formulário diferente e o último botão gera o documento final. 
2. Preencha os formulários "DP" (dados da Delegacia) e "Data" (hora, dia, mês e ano). Clique em  “Salvar Dados” após o preenchimento dos campos. O formulário “DP” pode ser preenchido somente quando necessário. Altere o campo “hora” do formulário “Data” a cada flagrante. Para apagar todos os campos, clique em “Limpar Campos”. O OpenOffice vai perguntar se deseja realmente exluir o registro. Confirme a exclusão. 
3. Preencha os formulários "Testemunhas" (condutor e testemunhas), "Delito" (tipo de delito e objetos), “Especialistas" (peritos, constatadores e avaliadores) e "Oficios" (número dos ofícios) a cada flagrante.  Salve os dados após o preenchimento dos formulários.
4. Preencha os formulários "Conduzidos" e "Vitimas". Registre quantos conduzidos ou vítimas forem necessários. As vítimas e os conduzidos devem ser registrados sequencialmente, começando pelo número '1'. Para navegar, incluir ou apagar registros, use a Barra de Navegação (figura abaixo). Lembre-se de sempre salvar os dados após o preenchimento dos formulários.
5. Clique no último botão para gerar o Inquérito Policial. Aguarde alguns instantes (1 ou 2 minutos).
6. O arquivo será salvo no diretório "Saida" e depois será aberto para edição. 
7. Leia e verifique o arquivo, faça as correções necessárias e salve antes de imprimir. Dica: use o corretor ortográfico, use o Ctrl+F para encontrar/substituir palavras. 

## IMPORTANTE!

A estrutura das pastas é importante para o funcionamento das Macros. Qualquer alteração na estrutura dos diretório impedirá o correto funcionamento da automatização. Não apague e não renomeie  os arquivos da pasta “Modelos”.

Se não houver conduzidos, não será possível a geração do IP.

O registro de vítimas é opcional. 

Limites do programa: 7200 caracteres no campo “depoimento” (vitimas, conduzidos, condutor).
65535 caracteres de depoimentos de conduzidos ou vítimas (todos os campos somados)

Após a geração do documento, será necessário efetuar ajustes manuais na formatação.

## HISTÓRICO

20100212 versão inicial – banco de dados, formulários, mala direta

20100301 versão Beta – macros para gerar mala direta e para montar o documento.

## LICENÇA DE USO

Direitos Autorais Reservados (c) 2010 Fábio Minami <fminami@gmail.com>

Esta biblioteca é Software Livre; você pode redistribuí-la e/ou modificá-la sob os termos da Licença Pública Geral Menor do GNU conforme publicada pela Free Software Foundation; tanto a versão 2.1 da Licença, ou (a seu critério) qualquer versão posterior.

Esta biblioteca é distribuída na expectativa de que seja útil, porém, SEM NENHUMA GARANTIA; nem mesmo a garantia implícita de COMERCIABILIDADE OU ADEQUAÇÃO A UMA FINALIDADE ESPECÍFICA. 
Consulte a Licença Pública Geral Menor do GNU para mais detalhes.


Você deve ter recebido uma cópia da Licença Pública Geral Menor do GNU junto com esta biblioteca; se não, escreva para a Free Software Foundation, Inc., no endereço 59 Temple Street, Suite 330, Boston, MA 02111-1307 USA. 
