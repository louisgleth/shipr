const fs = require("node:fs/promises");
const path = require("node:path");
const crypto = require("node:crypto");
const { execFile } = require("node:child_process");
const { promisify } = require("node:util");

const GMAIL_SETTINGS_SCOPE = "https://www.googleapis.com/auth/gmail.settings.basic";
const TOKEN_URL = "https://oauth2.googleapis.com/token";
const GMAIL_SEND_AS_BASE = "https://gmail.googleapis.com/gmail/v1/users/me/settings/sendAs";
const IAM_CREDENTIALS_SIGN_JWT_BASE = "https://iamcredentials.googleapis.com/v1/projects/-/serviceAccounts";
const execFileAsync = promisify(execFile);

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

async function getGcloudAccessToken() {
  const gcloudBin = process.env.GCLOUD_BIN || "gcloud";
  const { stdout } = await execFileAsync(gcloudBin, ["auth", "print-access-token"]);
  const token = String(stdout || "").trim();
  if (!token) {
    throw new Error("gcloud did not return an access token. Run `gcloud auth login` first.");
  }
  return token;
}

async function signJwtWithIamCredentials(payload, serviceAccountEmail, sourceAccessToken) {
  const response = await fetch(
    `${IAM_CREDENTIALS_SIGN_JWT_BASE}/${encodeURIComponent(serviceAccountEmail)}:signJwt`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${sourceAccessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ payload: JSON.stringify(payload) }),
    }
  );
  const result = await response.json().catch(() => ({}));
  if (!response.ok || !result.signedJwt) {
    throw new Error(`IAM signJwt failed: ${result.error?.message || result.error || response.status}`);
  }
  return result.signedJwt;
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

async function exchangeJwtForAccessToken(assertion, subjectEmail) {
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

async function getAccessTokenWithKey(serviceAccount, subjectEmail) {
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

  return exchangeJwtForAccessToken(assertion, subjectEmail);
}

async function getAccessTokenKeyless(serviceAccountEmail, subjectEmail, sourceAccessToken) {
  const now = Math.floor(Date.now() / 1000);
  const assertion = await signJwtWithIamCredentials(
    {
      iss: serviceAccountEmail,
      scope: GMAIL_SETTINGS_SCOPE,
      aud: TOKEN_URL,
      sub: subjectEmail,
      iat: now,
      exp: now + 3600,
    },
    serviceAccountEmail,
    sourceAccessToken
  );
  return exchangeJwtForAccessToken(assertion, subjectEmail);
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
    serviceAccountEmail: "",
    sourceAccessToken: "",
  };
  const credentialsIndex = argv.indexOf("--credentials");
  if (credentialsIndex !== -1) {
    options.credentialsPath = argv[credentialsIndex + 1] || "";
  }
  const serviceAccountEmailIndex = argv.indexOf("--service-account-email");
  if (serviceAccountEmailIndex !== -1) {
    options.serviceAccountEmail = argv[serviceAccountEmailIndex + 1] || "";
  }
  const sourceAccessTokenIndex = argv.indexOf("--source-access-token");
  if (sourceAccessTokenIndex !== -1) {
    options.sourceAccessToken = argv[sourceAccessTokenIndex + 1] || "";
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

  if (!options.credentialsPath && !options.serviceAccountEmail) {
    throw new Error("Missing --credentials path/to/service-account.json or --service-account-email name@project.iam.gserviceaccount.com. Run with --dry-run to render HTML only.");
  }

  const serviceAccount = options.credentialsPath
    ? JSON.parse(await fs.readFile(options.credentialsPath, "utf8"))
    : null;
  const sourceAccessToken = options.serviceAccountEmail
    ? options.sourceAccessToken || await getGcloudAccessToken()
    : "";

  for (const { user, html } of rendered) {
    const accessToken = serviceAccount
      ? await getAccessTokenWithKey(serviceAccount, user.email)
      : await getAccessTokenKeyless(options.serviceAccountEmail, user.email, sourceAccessToken);
    await updateSignature(accessToken, user.email, html);
    console.log(`Updated Gmail signature for ${user.email}`);
  }
}

main().catch((error) => {
  console.error(error.message || error);
  process.exitCode = 1;
});
