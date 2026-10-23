# Auto Email 4.0

Sistema de automação de e-mails personalizados para Google Sheets, com templates, envio por Gmail, painel analítico e relatório de rastreamento mais consultivo.

**Projeto original:** [ademirvarjao/autoemail](https://github.com/ademirvarjao/autoemail) · **Autor original:** Ademir Varjão · **Versão:** 4.0.0

> O repositório `autoemail-4.0` preserva os recursos disponíveis na versão 3.2 e evolui a experiência, a persistência e os relatórios. A aplicação continua sendo um projeto vinculado ao Google Apps Script/Sheets; não é um serviço de disparo independente.

## O que mudou na versão 4.0

- **Dashboard executivo:** mensagens aceitas pelo Gmail, aberturas únicas estimadas, taxa de abertura rastreável, falhas, tendência de sete dias e campanhas recentes.
- **Relatório consultivo:** recomendações variam conforme falhas, sinais de abertura e configuração do pixel; inclui limites e cuidados para interpretar os números.
- **Rastreamento detalhado:** busca, filtros por status/período e exportação CSV; exibe primeira abertura e carregamentos do pixel.
- **Persistência estruturada:** o Google Sheets funciona como banco de dados do projeto, com abas técnicas para campanhas, destinatários e eventos. As abas técnicas ficam ocultas para reduzir ruído; os dados continuam na planilha.
- **Migração do histórico v3.2:** os relatórios e registros antigos são importados uma vez, quando o banco é inicializado. Como a versão anterior não relacionava cada destinatário ao disparo com segurança, o histórico é identificado e mantido separado para evitar taxas ou comparações enganosas.
- **Mais segurança no envio:** validação de e-mails e CC global, estimativa da cota incluindo CCs, registro de falhas por linha e aviso antes de processar.
- **Editor renovado:** interface responsiva, visualização por destinatário, edição HTML, variáveis, assinatura Gmail, imagens inline e anexos.
- **Templates mantidos:** salvar, atualizar, renomear, excluir, importar/exportar JSON seletivamente; conteúdo extenso pode continuar armazenado no Google Drive.
- **Compatibilidade:** nomes das abas de relatório e nomes das principais funções públicas da versão anterior foram mantidos.

## Funcionalidades preservadas

### Criação e personalização

- Substituição de variáveis da planilha no assunto e no corpo: `Olá {{ Nome }}, o valor é {{ Valor }}.`
- Editor HTML com negrito, itálico, sublinhado, alinhamento, listas, links, cores e imagens.
- Pré-visualização com navegação entre as linhas visíveis.
- Imagens incorporadas como conteúdo inline (CID) no envio.
- Assinatura principal do Gmail, quando a API avançada do Gmail está habilitada.

### Envio

- Destinatário principal e endereços adicionais na célula separados por vírgula ou ponto e vírgula; os endereços adicionais e o CC global são enviados em cópia.
- Envio de teste personalizado com a primeira linha visível.
- Envio em massa limitado às linhas visíveis na planilha.
- Verificação prévia de cota diária considerando destinatário principal e CCs.
- Anexos de até 10 MB por arquivo e 20 MB no total; imagens inline de até 5 MB cada e 10 MB somadas. Os limites do Google/Gmail também se aplicam.

### Templates

- Criar, carregar, atualizar, renomear e excluir modelos.
- Importar e exportar modelos em JSON, incluindo seleção dos modelos a exportar.
- Templates grandes armazenados em uma pasta do Google Drive chamada `Templates AutoEmail (Sistema)`.

## Banco de dados e relatórios

O banco usa a própria planilha vinculada ao Apps Script — não há servidor ou banco externo para provisionar. As abas técnicas são:

| Aba | Conteúdo |
| --- | --- |
| `AE • Campanhas` | Assunto, origem, status, totais, duração, estimativa da cota do Gmail e métricas agregadas. |
| `AE • Destinatários` | Um registro por linha/mensagem processada, status, destinatários, falha e primeira/última abertura. |
| `AE • Eventos` | Eventos de carregamento do pixel associados ao token da mensagem. |

As abas visíveis `📊 Relatório de Envios` e `📊 Rastreamento Detalhado` continuam disponíveis e foram ampliadas para a leitura operacional. O sistema usa `LockService` para reduzir conflitos de gravação simultânea.

**Importante:** ocultar uma aba não é controle de acesso. Quem já tem permissão de edição na planilha pode reexibir e consultar as abas técnicas. Restrinja o compartilhamento da planilha aos usuários autorizados e defina uma política de retenção adequada.

## Instalação no Google Apps Script

1. Abra a planilha que contém a lista de contatos e escolha **Extensões → Apps Script**.
2. Copie os arquivos `.gs` deste repositório para o projeto Apps Script. Crie também um arquivo HTML chamado exatamente `Painel` e copie `Painel.html`.
3. Ative a opção **Mostrar o arquivo de manifesto `appsscript.json`** nas configurações do projeto e substitua o manifesto pelo arquivo deste repositório. Ele declara o runtime V8, os escopos e o serviço avançado do Gmail. Se a importação da assinatura falhar, confirme também se o Gmail API está disponível em **Serviços** no Apps Script/Cloud Project.
4. Em `Config.gs`, ajuste `ABA_DADOS` (por padrão, `Contatos`), `NOME_REMETENTE` e, se necessário, `DELAY_ENVIO_MS`.
5. Salve o projeto, volte à planilha e recarregue a página. No menu **Auto Email 4.0**, selecione **Inicializar banco de dados** ou abra o dashboard. Autorize os acessos solicitados.
6. O banco é criado e o histórico disponível é migrado. O menu antigo da planilha pode continuar aparecendo como **Auto Email** se o código original ainda estiver implantado; use o novo menu depois de salvar e recarregar.
7. Abra **Auto Email 4.0 → Abrir dashboard e painel**. O editor, os relatórios e os templates ficam no painel lateral.

### Habilitar rastreamento de abertura

1. No Apps Script, selecione **Implantar → Nova implantação → App da Web**.
2. Escolha **Executar como: eu**. O pixel precisa ser acessível sem login; selecione a política de acesso permitida pela sua organização (em muitas contas, “Qualquer pessoa”).
3. Implante, copie a URL de produção terminada em `/exec` e cole em **Configurações → Rastreamento de abertura** no painel.
4. Verifique o status no dashboard e envie um teste antes de uma campanha real.
5. Se publicar uma nova versão/deployment, confira se a URL ainda é a esperada.

O rastreamento é opcional. Sem URL configurada, os envios, falhas e campanhas continuam sendo armazenados; somente as aberturas deixam de ser mensuráveis.

## Fluxo de uso

1. Na aba de contatos, mantenha os cabeçalhos na linha 1 e inclua uma coluna chamada `Email`, `E-mail` ou `E mail`. Configure `ABA_DADOS` se a aba não se chamar `Contatos`.
2. Acesse **Auto Email 4.0 → Abrir dashboard e painel**. No editor, selecione ou crie um template, insira variáveis e revise anexos/CC.
3. Use **Enviar teste** e **Pré-visualizar** para conferir assunto, destinatários e formatação.
4. Inicie o disparo. O sistema considera somente as linhas visíveis e confirma a cota estimada antes de enviar.
5. Acompanhe campanhas no dashboard e consulte **Rastreamento** para buscar destinatários, revisar falhas e exportar CSV.

## Estrutura do repositório

- `Config.gs` — configurações, menu e inicialização da interface.
- `Email.gs` — leitura das linhas visíveis, editor, pré-visualização, anexos e envio.
- `Templates.gs` — ciclo de vida de templates e armazenamento no Drive.
- `Database.gs` — modelo de dados, migração, relatórios, dashboard e CSV.
- `WebApp.gs` — endpoint do pixel e registro de eventos.
- `Utils.gs` — normalização, variáveis, e-mails e utilitários puros.
- `Painel.html` — dashboard, editor e relatório de rastreamento.
- `appsscript.json` — runtime, escopos de autorização e serviço Gmail avançado.
- `tests/` — testes locais de lógica pura e validação sintática.

## Desenvolvimento e testes locais

O código executado em produção é o Google Apps Script. Para validar utilitários e sintaxe localmente, com Node.js 18 ou superior:

```bash
npm test
```

Os testes usam apenas módulos nativos do Node; não é necessário instalar dependências.

## Privacidade e leitura dos indicadores

- O rastreamento grava um token aleatório por mensagem, horários e quantidade de carregamentos do pixel. Não grava IP ou user-agent.
- Uma imagem pode ser carregada automaticamente por proxy/recursos de privacidade ou bloqueada pelo cliente de e-mail. Portanto, a taxa é **estimativa**, pode ficar acima ou abaixo do comportamento humano e não comprova leitura, entrega na caixa de entrada, intenção ou conversão.
- O token é por mensagem/linha da planilha. Se a linha inclui CCs, o pixel não revela qual dos destinatários carregou a imagem.
- “Mensagens aceitas pelo Gmail” significa que a chamada de envio foi processada pelo Gmail; não é confirmação de entrega final.
- Use respostas, cliques, conversões e feedback da equipe como sinais complementares antes de decidir reenviar ou alterar uma campanha.
- Respeite as regras aplicáveis de privacidade, consentimento, comunicações comerciais e retenção de dados.
- As cotas do Gmail e os limites de execução do Apps Script são definidos pelo Google e podem variar por conta.

## Licença e uso

O repositório original informa que o uso é restrito e sujeito à autorização do desenvolvedor. Esta atualização mantém esse aviso: disponibilização, redistribuição e uso comercial devem ser autorizados pelo titular. **Desenvolvido originalmente por Ademir Varjão.**
