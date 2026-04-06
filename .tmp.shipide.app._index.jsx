import { useLoaderData } from "react-router";
import { boundary } from "@shopify/shopify-app-react-router/server";
import { authenticate } from "../shopify.server";
import styles from "../shipide-embedded-home.module.css";

function buildPortalUrl(shop = "") {
  const rawBase = String(process.env.SHIPIDE_PORTAL_URL || "https://portal.shipide.com").trim();
  const portalUrl = new URL(rawBase || "https://portal.shipide.com");
  portalUrl.searchParams.set("source", "shopify-embedded");
  if (shop) {
    portalUrl.searchParams.set("shop", shop);
  }
  return portalUrl.toString();
}

export const loader = async ({ request }) => {
  const { session } = await authenticate.admin(request);
  const shop = String(session?.shop || "").trim();

  return {
    shop,
    portalUrl: buildPortalUrl(shop),
  };
};

export default function Index() {
  const { shop, portalUrl } = useLoaderData();

  return (
    <div className={styles.shell}>
      <div className={styles.backdrop} aria-hidden="true">
        <span className={styles.glowPrimary} />
        <span className={styles.glowAccent} />
        <span className={styles.grid} />
      </div>

      <main className={styles.frame}>
        <section className={styles.card}>
          <div className={styles.brandRow}>
            <span className={styles.brandMark} aria-hidden="true">
              <span />
              <span />
            </span>
            <span className={styles.brandText}>Shipide</span>
            <span className={styles.brandDivider}>|</span>
            <span className={styles.brandSubtext}>Portal</span>
          </div>

          <div className={styles.copyBlock}>
            <p className={styles.eyebrow}>Shopify Embedded App</p>
            <h1 className={styles.title}>Connected to Shipide</h1>
            <p className={styles.lead}>
              Your Shopify store is linked to Shipide. Open the portal to import orders,
              manage fulfillment settings, and run your shipping workflow.
            </p>
          </div>

          <dl className={styles.metaGrid}>
            <div className={styles.metaCard}>
              <dt>Store</dt>
              <dd>{shop || "Connected Shopify store"}</dd>
            </div>
            <div className={styles.metaCard}>
              <dt>Status</dt>
              <dd>Embedded session active</dd>
            </div>
            <div className={styles.metaCard}>
              <dt>Workspace</dt>
              <dd>Shipide Portal</dd>
            </div>
          </dl>

          <div className={styles.actions}>
            <a
              className={styles.primaryAction}
              href={portalUrl}
              target="_blank"
              rel="noreferrer"
            >
              Open Shipide Portal
            </a>
            <p className={styles.actionHint}>
              Keep this app installed so Shipide can maintain your Shopify connection.
            </p>
          </div>
        </section>
      </main>
    </div>
  );
}

export const headers = (headersArgs) => {
  return boundary.headers(headersArgs);
};
