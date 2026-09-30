# Hermes MSGraph

Uma biblioteca Python para interagir com a Microsoft Graph API de forma simples, cobrindo e-mails, pastas de caixa de correio, Planner, usuários e Microsoft Teams.

Usa autenticação **app-only** (client credentials), ideal para automações e integrações de backend que não dependem de um usuário logado.

## Instalação

Via pip, direto do GitHub:

```bash
pip install git+https://github.com/celoramalho/hermes_msgraph.git
```

Ou clonando o repositório para desenvolvimento local:

```bash
git clone https://github.com/celoramalho/hermes_msgraph.git
cd hermes_msgraph
pip install -e .
```

## Pré-requisitos no Azure AD / Entra ID

Registre uma aplicação no [Azure Portal](https://portal.azure.com) e conceda as permissões de aplicativo (application permissions, com consentimento do admin) necessárias para os recursos que for usar, por exemplo:

- `Mail.Send`, `Mail.ReadWrite` — envio e leitura de e-mails
- `User.Read.All` — listar/buscar usuários e licenças
- `Sites.Read.All` — listar sites do SharePoint
- `Group.Read.All`, `Tasks.Read.All` — Planner
- `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `ChannelMessage.Read.All`, `Chat.Read.All`, `ChatMember.Read.All` — Teams

Você vai precisar de `client_id`, `client_secret` e `tenant_id` dessa aplicação.

## Uso básico

```python
from hermes_msgraph import HermesMSGraph

hermes = HermesMSGraph(
    client_id="SEU_CLIENT_ID",
    client_secret="SEU_CLIENT_SECRET",
    tenant_id="SEU_TENANT_ID",
)

# Enviar um e-mail
hermes.send_email(
    sender_mail="caixa@suaempresa.com",
    subject="Teste",
    body="Enviado via Hermes MSGraph",
    to_address="destinatario@suaempresa.com",
)

# Listar usuários do tenant
usuarios = hermes.get_all_users(data="simple")
```

Veja o notebook [`examples/hermes_msgraph_examples.ipynb`](examples/hermes_msgraph_examples.ipynb) para exemplos de todas as funções, organizados por serviço.

## Serviços disponíveis

### E-mail (`EmailService`)

| Método | Descrição |
|---|---|
| `send_email` | Envia um e-mail, com suporte a CC, anexos e agendamento (`delay`). |
| `get_emails` | Lista e-mails de uma caixa, com filtros por assunto, pasta, remetente, data, etc. |
| `get_email_by_id` | Busca um e-mail específico pelo ID. |
| `list_email_attachments` | Lista os anexos de um e-mail. |
| `download_attachment` | Baixa um anexo específico para um arquivo local. |
| `forward_email_by_id` | Encaminha um e-mail existente. |
| `list_sharepoint_sites` | Lista sites raiz do SharePoint do tenant. |

### Pastas de caixa de correio (`MailboxFolderService`)

| Método | Descrição |
|---|---|
| `list_mailbox_folders` | Lista todas as pastas de uma caixa de correio. |
| `get_mailbox_folders` | Mesmo resultado, como `pandas.DataFrame`. |
| `get_folder_id` | Retorna o ID de uma pasta a partir do nome. |
| `validate_folder_id` | Verifica se um ID de pasta existe na caixa. |

### Planner (`PlannerService`)

| Método | Descrição |
|---|---|
| `list_plans_by_group_id` | Lista planos de um grupo do Microsoft 365. |
| `list_visible_plans_by_user_id` | Lista planos visíveis para um usuário. |
| `list_tasks_by_user_id` | Lista tarefas atribuídas a um usuário. |
| `list_tasks_by_plan_id` | Lista tarefas de um plano específico. |

### Usuários (`UsersService`)

| Método | Descrição |
|---|---|
| `get_user_id_by_email` | Retorna o ID (GUID) de um usuário a partir do e-mail. |
| `get_all_users` | Lista todos os usuários do tenant (`data="all"` ou `data="simple"`). |
| `search_from_mailboxes` | Busca usuários por nome/e-mail (`$search`). |
| `get_tenant_licenses` | Lista as licenças (SKUs) disponíveis no tenant, com nome amigável. |

> `add_user_to_shared_mailbox` é referenciada em `HermesMSGraph` mas **ainda não está implementada**: a Microsoft Graph API não expõe um endpoint equivalente a `Add-MailboxPermission`/`Add-RecipientPermission` do Exchange Online PowerShell. Conceder acesso a uma shared mailbox hoje só é possível via Exchange Online PowerShell ou pelo Admin Center.

### Teams (`TeamsService`)

| Método | Descrição |
|---|---|
| `list_all_teams` | Lista todos os times do tenant. |
| `list_joined_teams_by_user_id` | Lista os times de que um usuário participa. |
| `list_channels_by_team_id` | Lista canais de um time (`include_private=True` inclui privados/compartilhados). |
| `get_channel_by_id` | Busca um canal específico. |
| `list_channel_messages` | Lista mensagens de um canal (`include_replies=True` aninha as respostas). |
| `list_channel_message_replies` | Lista as respostas de uma mensagem de canal. |
| `delta_channel_messages` | Sincronização incremental de mensagens de canal via delta query. |
| `list_chats_by_user_id` | Lista chats (1:1, grupo, reunião) de um usuário. |
| `get_chat_by_id` | Busca um chat específico. |
| `list_chat_members` | Lista membros de um chat. |
| `list_chat_messages` | Lista mensagens de um chat. |
| `delta_chat_messages` | Sincronização incremental de mensagens de chat via delta query. |
| `download_hosted_content` | Baixa conteúdo hospedado referenciado numa mensagem (ex.: imagens inline). |

## Tratamento de erros

Todos os métodos lançam `HermesMSGraphError` (em `hermes_msgraph.exceptions`) quando a API retorna um erro:

```python
from hermes_msgraph.exceptions import HermesMSGraphError

try:
    hermes.get_user_id_by_email("naoexiste@suaempresa.com")
except HermesMSGraphError as e:
    print(f"Erro: {e}")
```

## Requisitos

- Python 3.7+
- `requests`, `pandas`, `pyyaml`, `tqdm`

## Licença

Este projeto é licenciado sob a licença MIT.
