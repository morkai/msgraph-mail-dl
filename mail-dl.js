require('isomorphic-fetch');

const os = require('os');
const path = require('path');
const fs = require('fs-extra');
const {Client} = require("@microsoft/microsoft-graph-client");
const {TokenCredentialAuthenticationProvider} = require("@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials");
const {ClientSecretCredential} = require("@azure/identity");

const targetDir    = process.env.WMES_MAIL_DL_TARGET_DIR;
const userId       = process.env.WMES_MAIL_DL_USER_ID;
const tenantId     = process.env.WMES_MAIL_DL_TENANT_ID;
const clientId     = process.env.WMES_MAIL_DL_CLIENT_ID;
const clientSecret = process.env.WMES_MAIL_DL_CLIENT_SECRET;

const credential = new ClientSecretCredential(tenantId, clientId, clientSecret, {

});

const authProvider = new TokenCredentialAuthenticationProvider(credential, {
  scopes: ['https://graph.microsoft.com/.default']
});

const client = Client.initWithMiddleware({
  debugLogging: false,
  authProvider
});

(async function processMessages()
{
  console.log(`${now()} Fetching messages...`);

  const messagesRes = await client.api(`/users/${userId}/mailFolders('Inbox')/messages?$top=11&$select=from,toRecipients,ccRecipients,bccRecipients,subject,body,receivedDateTime&$orderby=receivedDateTime&expand=attachments($select=name,contentType,size,microsoft.graph.fileAttachment/contentId)`).get();

  if (!messagesRes.value.length)
  {
    console.log(`${now()} No messages found!`);

    return;
  }

  console.log(`${now()} Found at least ${messagesRes.value.length} messages!`);

  const hasMoreMessages = messagesRes.value.length > 10;
  const messages = messagesRes.value.slice(0, 10);

  for (const message of messages)
  {
    console.log(`${now()} Processing message...`, {
      id: message.id,
      subject: message.subject,
      from: message.from.emailAddress
    });

    const tmpDir = path.join(os.tmpdir(), `WMES_EMAIL`);

    console.log(`${now()} Preparing TEMP directory...`);
    await fs.ensureDir(tmpDir);

    if (matchMessage(message))
    {
      if (message.attachments)
      {
        for (const attachment of message.attachments)
        {
          if (!attachment['@odata.type'].includes('fileAttachment'))
          {
            console.log(`${now()} Skipping attachment...`, {
              id: attachment.id,
              name: attachment.name,
              size: attachment.size,
              type: attachment['@odata.type']
            });

            continue;
          }

          console.log(`${now()} Downloading attachment...`, {
            id: attachment.id,
            name: attachment.name,
            size: attachment.size
          });

          await downloadAttachment(
            tmpDir,
            attachment,
            client.api(`/users/${userId}/messages/${message.id}/attachments/${attachment.id}/$value`).getStream()
          );
        }
      }

      const receivedAt = Math.floor(Date.parse(message.receivedDateTime) / 1000);
      const email = {
        id: message.id,
        receivedAt: message.receivedDateTime,
        subject: message.subject,
        from: formatEmailAddress(message.from),
        to: message.toRecipients.map(formatEmailAddress),
        cc: message.ccRecipients.map(formatEmailAddress),
        bcc: message.bccRecipients.map(formatEmailAddress),
        body: message.body.content,
        attachments: message.attachments.map(attachment =>
        {
          return {
            id: attachment.id,
            contentId: attachment.contentId,
            contentType: attachment.contentType,
            name: attachment.name,
            size: attachment.size
          };
        })
      };

      console.log(`${now()} Writing email.json...`);
      await fs.writeFile(path.join(tmpDir, 'email.json'), JSON.stringify(email, null, 2));

      console.log(`${now()} Moving to target directory...`);
      await fs.move(tmpDir, path.join(targetDir, `${receivedAt}@EMAIL_${Date.now()}`));
    }
    else
    {
      console.log(`${now()} Message skipped!`);
    }

    console.log(`${now()} Deleting message...`);
    await client.api(`/users/${userId}/messages/${message.id}`).delete();
  }

  if (hasMoreMessages)
  {
    console.log(`${now()} Waiting for more messages...`);
    setTimeout(processMessages, 5000);
  }
})();

function downloadAttachment(tmpDir, attachment, stream)
{
  return new Promise((resolve, reject) =>
  {
    stream
      .then(stream =>
      {
        const writeStream = fs.createWriteStream(path.join(tmpDir, attachment.name));

        stream.pipe(writeStream).on('error', reject);

        writeStream.on('error', reject);

        writeStream.on('finish', resolve);
      })
      .catch(reject);
  });
}

function formatEmailAddress({emailAddress})
{
  if (emailAddress.name)
  {
    return `${emailAddress.name} <${emailAddress.address}>`;
  }

  return emailAddress.address;
}

function matchMessage(message)
{
  if (/ETO.*?([0-9]{12}|[0-9A-Z]{7})/i.test(message.subject))
  {
    return true;
  }

  if (/(zpw|near\s*miss|kaizen|uspraw).*?[0-9]+/i.test(message.subject))
  {
    return message.attachments.some(a => a.contentType.startsWith('image/'));
  }

  return false;
}

function now()
{
  return new Date().toISOString();
}
