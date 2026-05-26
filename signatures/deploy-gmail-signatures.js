const fs = require("node:fs/promises");
const path = require("node:path");
const crypto = require("node:crypto");

const GMAIL_SETTINGS_SCOPE = "https://www.googleapis.com/auth/gmail.settings.basic";
const TOKEN_URL = "https://oauth2.googleapis.com/token";
const GMAIL_SEND_AS_BASE = "https://gmail.googleapis.com/gmail/v1/users/me/settings/sendAs";

function base64Url(input) {
  return Buffer.from(input)
    .toString("base64")
    .replace(/=/g, "")
    .replace(/\+/g, "-")
    .replace(/\//g, "_");
}

function signJwt(payload, privateKey) {
  const header = { alg: "RS256", typ: "JWT" };
  const encodedHeader = base64Url(JSON.stringify(header));
  const encodedPayload = base64Url(JSON.stringify(payload));
  const data = `${encodedHeader}.${encodedPayload}`;
  const signature = crypto.createSign("RSA-SHA256").update(data).sign(privateKey);
  return `${data}.${base64Url(signature)}`;
}

function htmlEscape(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function renderTemplate(template, user) {
  return template.replace(/\{\{(\w+)\}\}/g, (_match, key) => htmlEscape(user[key]));
}

async function getAccessToken(serviceAccount, subjectEmail) {
  const now = Math.floor(Date.now() / 1000);
  const assertion = signJwt(
    {
      iss: serviceAccount.client_email,
      scope: GMAIL_SETTINGS_SCOPE,
      aud: TOKEN_URL,
      sub: subjectEmail,
      iat: now,
      exp: now + 3600,
    },
    serviceAccount.private_key
  );

  const body = new URLSearchParams({
    grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
    assertion,
  });

  const response = await fetch(TOKEN_URL, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok || !payload.access_token) {
    throw new Error(`Token request failed for ${subjectEmail}: ${payload.error_description || payload.error || response.status}`);
  }
  return payload.access_token;
}

async function updateSignature(accessToken, userEmail, signatureHtml) {
  const sendAsEmail = encodeURIComponent(userEmail);
  const response = await fetch(`${GMAIL_SEND_AS_BASE}/${sendAsEmail}`, {
    method: "PATCH",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ signature: signatureHtml }),
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(`Gmail signature update failed for ${userEmail}: ${payload.error?.message || response.status}`);
  }
  return payload;
}

function parseArgs(argv) {
  const options = {
    dryRun: argv.includes("--dry-run"),
    credentialsPath: "",
  };
  const credentialsIndex = argv.indexOf("--credentials");
  if (credentialsIndex !== -1) {
    options.credentialsPath = argv[credentialsIndex + 1] || "";
  }
  return options;
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  const root = __dirname;
  const template = await fs.readFile(path.join(root, "shipide-signature-template.html"), "utf8");
  const users = JSON.parse(await fs.readFile(path.join(root, "users.json"), "utf8"));
  const rendered = users.map((user) => ({
    user,
    html: renderTemplate(template, user),
  }));

  if (options.dryRun) {
    const outDir = path.join(root, "dist");
    await fs.mkdir(outDir, { recursive: true });
    await Promise.all(
      rendered.map(({ user, html }) =>
        fs.writeFile(path.join(outDir, `${user.email.replace(/[^a-z0-9]+/gi, "_")}.html`), html)
      )
    );
    console.log(`Rendered ${rendered.length} signature preview file(s) to ${outDir}`);
    return;
  }

  if (!options.credentialsPath) {
    throw new Error("Missing --credentials path/to/service-account.json. Run with --dry-run to render HTML only.");
  }

  const serviceAccount = JSON.parse(await fs.readFile(options.credentialsPath, "utf8"));
  for (const { user, html } of rendered) {
    const accessToken = await getAccessToken(serviceAccount, user.email);
    await updateSignature(accessToken, user.email, html);
    console.log(`Updated Gmail signature for ${user.email}`);
  }
}

main().catch((error) => {
  console.error(error.message || error);
  process.exitCode = 1;
});
