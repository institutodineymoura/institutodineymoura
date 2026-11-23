# Lista de avisos — como configurar

Quem clica em **"Quero ser avisado"** no site fica guardado numa Planilha Google.
Quando o Instituto publica uma data na planilha, a data aparece no site e todo mundo da lista recebe um e-mail.

Por enquanto só o **Exame de vista** está ligado.

## 1. Criar a planilha (uma vez só, uns 10 minutos)

1. Entre no Google com **institutodineymoura@gmail.com**.
2. Crie uma planilha nova em <https://sheets.new> e dê o nome **Avisos do site**.
3. No menu, clique em **Extensões → Apps Script**.
4. Apague tudo o que estiver escrito, cole o conteúdo do arquivo `Codigo.gs` e clique no disquete (**Salvar**).
5. Na barra de cima, escolha a função **configurar** e clique em **Executar**.
   - O Google vai pedir autorização. Escolha a conta do Instituto.
   - Se aparecer "O Google não verificou este app", clique em **Avançado → Acessar Avisos do site (não seguro)** → **Permitir**. É normal: o script é do próprio Instituto.
6. Volte à planilha: as abas **Interessados** e **Agenda** foram criadas.

## 2. Ligar a planilha ao site

1. No Apps Script, clique em **Implantar → Nova implantação**.
2. Na engrenagem, escolha **App da Web**.
3. Em **Executar como**, escolha **Eu (institutodineymoura@gmail.com)**. Em **Quem pode acessar**, escolha **Qualquer pessoa**.
4. Clique em **Implantar** e copie o endereço que termina em `/exec`.
5. Cole esse endereço no `index.html`, na linha `const AVISOS_URL='';` (entre as aspas), e publique o site.
   A partir daí, o botão **"Quero ser avisado"** aparece no card de Exame de vista.

## 3. Como usar no dia a dia

**Quando sair uma data de exame de vista**, preencha uma linha na aba **Agenda**:

| Serviço | Data | Horário | Local | Vagas | Observações | Publicar e avisar |
|---|---|---|---|---|---|---|
| Exame de vista | 13/12/2026 | 8h às 12h | *(em branco = sede)* | 30 | Leve documento com foto. | ☑ |

- Ao marcar **Publicar e avisar**, a data aparece no site na hora.
- Os e-mails saem sozinhos **em até 1 hora**. Para mandar na hora, use o menu **Instituto → Enviar avisos agora**.
- Cada pessoa recebe o aviso **uma vez só**. Depois a situação dela muda para **Avisado** e aparece um link **"Enviar no WhatsApp"**, para quem deixou o número.
- Quem se cadastrar depois que a data já foi avisada entra na lista e recebe o aviso da **próxima** data.
- Depois que o dia passa, a data sai do site sozinha.

**Se alguém responder SAIR**, mude a **Situação** dessa pessoa para **Cancelado**.

**Limite:** uma conta Gmail comum manda cerca de 100 e-mails por dia por aqui. Se a lista passar disso, o resto é enviado no dia seguinte, sem precisar fazer nada.

## Se precisar mudar o script depois

Cole o código novo, salve e vá em **Implantar → Gerenciar implantações → lápis (editar) → Versão: Nova versão → Implantar**.
O endereço `/exec` continua o mesmo, então o site não precisa mudar.
