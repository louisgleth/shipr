create extension if not exists pgcrypto;

create table if not exists public.clickwrap_contract_versions (
  id uuid primary key default gen_random_uuid(),
  version text not null unique,
  title text not null,
  body_text text not null,
  hash_sha256 text not null check (char_length(hash_sha256) = 64),
  pdf_url text,
  effective_at timestamptz not null default now(),
  is_active boolean not null default false,
  created_at timestamptz not null default now(),
  created_by uuid references auth.users(id) on delete set null,
  notes text
);

alter table public.clickwrap_contract_versions
  add column if not exists pdf_url text;

create unique index if not exists clickwrap_contract_single_active_idx
  on public.clickwrap_contract_versions ((is_active))
  where is_active;

create index if not exists clickwrap_contract_effective_idx
  on public.clickwrap_contract_versions (effective_at desc, created_at desc);

create table if not exists public.clickwrap_acceptances (
  id uuid primary key default gen_random_uuid(),
  user_id uuid references auth.users(id) on delete set null,
  invite_id uuid references public.registration_invites(id) on delete set null,
  accepted_email text not null,
  invited_email text,
  contract_id uuid references public.clickwrap_contract_versions(id) on delete set null,
  contract_version text not null,
  contract_hash_sha256 text not null check (char_length(contract_hash_sha256) = 64),
  agreement_method text not null default 'scroll_clickwrap',
  scrolled_to_end boolean not null default false,
  scrolled_to_end_at timestamptz,
  agreed_at timestamptz not null,
  ip_address text,
  user_agent text,
  client_timezone text,
  client_locale text,
  proof_digest_sha256 text not null check (char_length(proof_digest_sha256) = 64),
  proof_signature_hmac text not null,
  proof_payload jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now()
);

create index if not exists clickwrap_acceptances_user_agreed_idx
  on public.clickwrap_acceptances (user_id, agreed_at desc, created_at desc);

create index if not exists clickwrap_acceptances_invite_idx
  on public.clickwrap_acceptances (invite_id, agreed_at desc);

create index if not exists clickwrap_acceptances_contract_idx
  on public.clickwrap_acceptances (contract_version, agreed_at desc);

alter table public.clickwrap_contract_versions enable row level security;
alter table public.clickwrap_acceptances enable row level security;

drop policy if exists clickwrap_contract_versions_select_authenticated on public.clickwrap_contract_versions;
create policy clickwrap_contract_versions_select_authenticated
  on public.clickwrap_contract_versions
  for select
  to authenticated
  using (true);

drop policy if exists clickwrap_acceptances_select_own on public.clickwrap_acceptances;
create policy clickwrap_acceptances_select_own
  on public.clickwrap_acceptances
  for select
  to authenticated
  using (auth.uid() = user_id);

insert into public.clickwrap_contract_versions (
  version,
  title,
  body_text,
  hash_sha256,
  pdf_url,
  effective_at,
  is_active,
  notes
) values (
  'v2.0',
  'Commercial Agreement',
  $contract$
Commercial Agreement

By clicking "I Agree", creating an account, or using the Platform, the Merchant agrees to be bound by the terms of this Agreement.

By accepting, you confirm you are at least 18 years old and have the legal authority to bind the Merchant entity you represent to this Agreement.

Shipide records the date, time, IP address, account identity, and version of this Agreement accepted at the time of account creation or acceptance.

This Shipping Platform Services Agreement ("Agreement") governs access to and use of the Shipide platform and related services. Shipide is a commercial brand operated by Cryvelin LLC, a company incorporated under the laws of the State of Delaware, United States.

1. Definitions

For the purposes of this Agreement, the following definitions apply.

"Cryvelin LLC" means Cryvelin LLC, a limited liability company organized under the laws of the State of Delaware, United States, which operates the Shipide platform.

"Shipide" refers to the commercial brand and platform operated by Cryvelin LLC.

"Platform" means the Shipide online platform enabling Merchants to generate shipping labels and access shipping services.

"Merchant" means the business entity or individual registering for and using the Shipide Platform.

"Carrier" means any third-party shipping or logistics provider whose services are accessible through the Platform.

"Shipping Label" means any shipping label generated via the Platform for the transport of a parcel.

"Account Balance" means the prepaid funds deposited by the Merchant into its Shipide account.

"Invoice" means a billing document issued by Shipide for shipping labels generated during a billing period.

"Aggregated Data" means anonymised or pseudonymised shipment data, rate benchmarks, and logistics metrics derived from Merchant activity and used by Shipide for platform optimization, carrier negotiations, and analytics, as further described in Section 13.

"GDPR" means the EU General Data Protection Regulation (Regulation (EU) 2016/679).

2. Scope of Services

Shipide provides a technology platform that allows Merchants to access shipping services and generate shipping labels with third-party carriers at negotiated rates.

Shipide aggregates shipping volumes across multiple Merchants in order to obtain preferential shipping rates from carriers.

Shipide does not operate as a carrier and does not physically transport parcels. All shipments generated through the Platform are transported and handled by independent third-party Carriers.

3. Freedom of Use and Termination

The Merchant remains free at all times to decide whether or not to use the Shipide Platform.

This Agreement does not impose any minimum usage obligations, exclusivity requirements, or long-term commitments on the Merchant. The Merchant may stop using the Platform at any time without notice and without penalty.

The Merchant may terminate its account at any time by providing written notice to Shipide or by ceasing use of the Platform.

Termination or non-use of the Platform does not create any fees or penalties, provided that all outstanding amounts related to shipping labels previously generated through the Platform have been fully paid. Any unused prepaid Account Balance remaining at the time of account closure will be refunded to the Merchant within thirty (30) days, subject to deduction of any outstanding amounts owed.

4. Merchant Account

To access the Platform, the Merchant must create an account and provide accurate, complete, and up-to-date information during registration. The Merchant is responsible for maintaining the confidentiality of its account credentials and remains fully responsible for all activities conducted through its account. Shipide reserves the right to suspend or terminate accounts that contain inaccurate, misleading, or fraudulent information, or that are used in violation of this Agreement or applicable laws.

5. Generation and Purchase of Shipping Labels

Shipping labels may be generated through the Platform using available Carrier services. Each shipping label generated through the Platform constitutes a binding purchase by the Merchant. All shipping labels generated remain payable by the Merchant regardless of whether the shipment is ultimately used or cancelled after generation. The Merchant is responsible for ensuring the accuracy of all shipment information, including but not limited to weight, dimensions, destination, and shipment contents. Any additional charges imposed by Carriers due to inaccurate shipment information, adjustments, or re-classification of parcels will be charged to the Merchant.

6. Payment Methods

Shipide provides several payment methods for the purchase of shipping labels, including prepaid account balances, card payments, and monthly invoicing for selected Merchants.

6.1 Prepaid Account Balance — Merchants may fund their Shipide account by transferring funds via bank transfer to the IBAN designated by Shipide. Funds received will be credited to the Merchant's Account Balance and used to pay for shipping labels generated on the Platform. The cost of each label will be deducted from the Account Balance at the time of generation. Shipide reserves the right to suspend the ability to generate labels if the Account Balance is insufficient to cover the cost of the requested label.

6.2 Card Payments — Merchants may also register a debit or credit card within their account. Shipide may charge the registered payment method automatically for shipping labels generated or to replenish the Merchant's Account Balance. By providing a payment card, the Merchant authorizes Shipide to charge the payment method on file for all amounts due under this Agreement.

6.3 Monthly Invoicing — Monthly invoicing may be granted to selected Merchants at the sole discretion of Shipide. Eligibility for invoicing may depend on shipping volume, payment history, or other creditworthiness criteria determined by Shipide. Merchants granted monthly invoicing will receive invoices summarizing shipping labels generated during the billing period. Unless otherwise specified, invoices are payable within fourteen (14) days from the invoice date. Shipide reserves the right to revoke invoicing privileges at any time, with reasonable prior notice where commercially practicable.

7. Billing, Late Payment, and Penalties

Prompt and complete payment is essential to the operation of the Platform. Shipide operates on narrow margins with carrier partners, and unpaid balances directly affect its ability to maintain negotiated rates for all Merchants. The following provisions are strictly enforced.

7.1 Payment Deadlines — All invoices issued under this Agreement are due and payable within fourteen (14) calendar days from the invoice date, unless a different payment term has been agreed in writing between Shipide and the Merchant. Prepaid balance topups are due immediately upon request or as required to maintain a positive Account Balance.

7.2 Late Payment Interest — Any amount not received by Shipide on or before the due date will automatically and without prior notice accrue late payment interest at a rate of one and a half percent (1.5%) per month (equivalent to eighteen percent (18%) per annum), calculated from the due date until the date of full payment. This rate applies both to invoiced amounts and to carrier adjustment charges. Late payment interest will be added to the Merchant's outstanding balance and will appear on the next invoice issued.

7.3 Administrative Recovery Fee — In addition to accrued interest, any invoice that remains unpaid for more than thirty (30) calendar days past its due date will be subject to an administrative recovery fee of EUR 75 (or the equivalent in the applicable currency) per overdue invoice. This fee reflects the administrative burden of debt follow-up and is in addition to any late interest accrued under Section 7.2.

7.4 Suspension for Non-Payment — If any amount owed by the Merchant remains unpaid for more than seven (7) calendar days past its due date, Shipide may, without prior notice and without liability, immediately suspend the Merchant's access to the Platform, including the ability to generate new shipping labels. Suspension shall remain in effect until all outstanding amounts, including accrued interest and fees, have been paid in full.

7.5 Termination for Non-Payment — If payment remains outstanding for more than sixty (60) calendar days past its due date, Shipide reserves the right to permanently terminate the Merchant's account. Termination for non-payment does not extinguish the Merchant's obligation to pay all outstanding amounts, accrued interest, fees, and any collection costs incurred by Shipide, including reasonable legal and debt recovery expenses.

7.6 Debt Recovery Costs — The Merchant shall be liable for all reasonable costs incurred by Shipide in recovering any unpaid amounts, including but not limited to legal fees, court costs, and the fees of third-party debt collection agencies. Shipide may report persistent non-payment to credit reference agencies or commercial credit bureaux, which may affect the Merchant's credit profile.

7.7 Disputed Invoices — If the Merchant disputes any item on an invoice, it must notify Shipide in writing within seven (7) calendar days of receipt of the invoice, specifying the disputed amount and the grounds for the dispute. Undisputed portions of the invoice remain due and payable on the original due date. Failure to dispute an invoice within the prescribed period constitutes acceptance of the invoice in full.

7.8 Set-Off — The Merchant may not withhold or set off any amounts owed to Shipide against any claimed liability of Shipide to the Merchant, whether under this Agreement or otherwise, without the prior written consent of Shipide.

8. Credit Limits

Shipide may establish credit limits for Merchants who are granted monthly invoicing privileges. If a Merchant exceeds its assigned credit limit, Shipide may suspend the ability to generate new shipping labels until outstanding balances are settled. Shipide reserves the right to adjust or modify credit limits at its discretion based on usage, payment behavior, or risk considerations, with reasonable prior notice where commercially practicable.

9. Carrier Terms

All shipments generated through the Platform remain subject to the terms and conditions of the selected Carrier. The Merchant agrees to comply with all Carrier rules, including but not limited to restrictions on prohibited goods, packaging requirements, customs declarations, weight limits, and dimensional restrictions. In the event of any conflict between this Agreement and a Carrier's shipping conditions regarding the handling of shipments, the Carrier's terms shall prevail with respect to transport services.

10. Merchant Responsibilities

The Merchant is responsible for providing accurate shipment information and ensuring that all shipments comply with applicable laws and Carrier requirements. The Merchant must ensure that parcels are properly packaged and labeled in accordance with Carrier guidelines.

The Merchant further agrees not to ship any of the following categories of goods through the Platform: explosives, ammunition, weapons, or items regulated under firearms legislation; narcotics, controlled substances, or illegal drugs; flammable, corrosive, radioactive, toxic, or otherwise hazardous materials; lithium batteries or dangerous goods not properly declared and packaged per IATA/IMDG regulations; perishable items not authorized by the relevant Carrier; counterfeit or intellectual property-infringing goods; human remains, live animals, or biological specimens not authorized by law; and any items prohibited under applicable local, national, or international laws.

The Merchant remains solely responsible for the contents of all shipments generated through its account. Shipide assumes no responsibility for the legality or nature of items shipped by the Merchant.

11. Indemnification

The Merchant agrees to indemnify, defend, and hold harmless Cryvelin LLC, its officers, employees, agents, and partners from and against any and all claims, damages, fines, penalties, liabilities, costs, and expenses (including reasonable legal fees) arising from or related to the Merchant's shipments including any illegal, dangerous, or prohibited goods transported through the Platform.

12. Platform Role, Liability Cap, and Warranties

12.1 Platform Role — Shipide acts solely as a technology platform facilitating access to shipping services offered by third-party Carriers. Shipide is not a Carrier and does not physically transport shipments. Accordingly, Shipide shall not be liable for shipment delays, loss or damage to parcels, Carrier service disruptions, customs issues, delivery failures, or any other events arising from the transportation process. Any claims related to shipment loss or damage must be handled according to the claims procedures established by the relevant Carrier.

12.2 No Warranty — The Platform is provided "as is" and "as available", without warranties of any kind, express or implied, including but not limited to warranties of merchantability, fitness for a particular purpose, or uninterrupted availability. Shipide does not warrant that the Platform will be error-free or that it will be available at all times.

12.3 Liability Cap — To the maximum extent permitted by applicable law, Cryvelin LLC's total aggregate liability to the Merchant arising out of or in connection with this Agreement, whether in contract, tort (including negligence), or otherwise, shall not exceed the total amounts paid by the Merchant to Cryvelin LLC in the three (3) calendar months immediately preceding the event giving rise to the claim.

This cap applies regardless of the nature of the claim or the number of incidents giving rise to liability.

12.4 Exclusion of Consequential Loss — In no event shall Cryvelin LLC be liable for any indirect, incidental, special, consequential, or punitive damages, including but not limited to loss of revenue, loss of profits, loss of data, or loss of business, even if Cryvelin LLC has been advised of the possibility of such damages.

13. Data Governance, Privacy, and Shipment Data Pooling

13.1 Personal Data and GDPR — Cryvelin LLC processes personal data in connection with the provision of the Platform in accordance with the GDPR and all other applicable data protection laws.

The categories of personal data processed may include Merchant business contact details, recipient names and addresses, and transactional data related to shipments.

Merchants established in the EEA, or who process personal data of EEA-resident individuals through the Platform, acknowledge that Cryvelin LLC acts as a data processor in respect of shipment recipient data, and as a data controller in respect of Merchant account data.

The terms governing such processing, including the technical and organisational measures applied by Cryvelin LLC, the lawful bases for processing, sub-processor relationships, data subject rights procedures, and breach notification commitments, are set out in full in the Shipide Privacy Policy available at [Privacy Policy URL].

The Merchant's acceptance of this Agreement constitutes acceptance of those processing terms.

Merchants who require additional information about how their data is handled may contact Cryvelin LLC at legal@shipide.com.

13.2 Shipment Data Pooling and Platform Analytics — By using the Platform, the Merchant expressly acknowledges and consents to Shipide's right to collect, retain, and use Aggregated Data derived from the Merchant's shipment activity for the following purposes: negotiating and maintaining preferential rates with Carriers on behalf of all Platform users; optimizing routing, carrier selection, and delivery performance across the Platform; generating anonymised market intelligence and logistics analytics for internal use; improving the Platform's features, pricing models, and service quality; and benchmarking carrier performance and service levels.

Aggregated Data used for these purposes will be anonymised or pseudonymised in accordance with GDPR standards and will not be attributed to any individual Merchant or used to identify the Merchant's commercial relationships, customer identities, or business strategies. The Merchant retains ownership of its own raw shipment records as maintained in its account.

13.3 Data Retention — Shipide will retain Merchant shipment data for a period of five (5) years from the date of the relevant transaction, or for such longer period as may be required by applicable tax, customs, or regulatory law. After expiry of the retention period, data will be securely deleted or irreversibly anonymised.

13.4 Data Security — Shipide implements appropriate technical and organisational security measures to protect Merchant data against unauthorised access, accidental loss, disclosure, or destruction. These measures include encrypted data transmission, access controls, and regular security assessments. In the event of a personal data breach affecting Merchant data, Shipide will notify the Merchant without undue delay and, where required, notify the relevant supervisory authority in accordance with GDPR Article 33.

13.5 Merchant Obligations Regarding Data — The Merchant warrants that it has obtained all necessary consents and has a lawful basis under applicable data protection law for providing recipient personal data to Shipide for the purpose of generating shipping labels. The Merchant shall indemnify Shipide for any regulatory fines, penalties, or third-party claims arising from the Merchant's failure to comply with this obligation.

13.6 Data Portability and Deletion — Upon written request, Shipide will provide the Merchant with an export of its own account and shipment data in a standard machine-readable format. Upon account termination, Shipide will, upon request, delete or anonymise the Merchant's personally identifiable account data within ninety (90) days, subject to any legal retention obligations.

14. Intellectual Property

The Shipide Platform, software, brand, trademarks, API, documentation, and all related intellectual property rights are and shall remain the exclusive property of Cryvelin LLC. This Agreement does not transfer any intellectual property rights to the Merchant. The Merchant is granted a limited, non-exclusive, non-transferable, revocable licence to access and use the Platform solely for the purposes set out in this Agreement. The Merchant shall not reproduce, reverse engineer, decompile, or create derivative works of any part of the Platform without the prior written consent of Cryvelin LLC.

15. API, Integrations, and Automation

Where Shipide makes an API, Shopify integration, CSV import functionality, or other automated tools available to the Merchant, such tools are provided "as is" and subject to the terms of this Agreement. Shipide shall not be liable for errors, interruptions, or data losses arising from the use of such integrations. The Merchant is responsible for ensuring that any data transmitted through automated integrations is accurate and compliant with this Agreement. Shipide may impose rate limits, access controls, or technical restrictions on API usage to maintain Platform stability.

16. Suspension and Termination

Shipide may suspend or terminate access to the Platform immediately and without prior notice if the Merchant fails to pay any amounts due within the timeframes specified in Section 7; exceeds its assigned credit limit and fails to remedy within 48 hours of notification; engages in fraudulent, abusive, or unlawful activity; ships or attempts to ship prohibited or illegal goods; or materially breaches any term of this Agreement and, where the breach is remediable, fails to remedy it within seven (7) days of written notice.

The Merchant may terminate its account at any time by providing written notice to Shipide. Termination of the account does not release the Merchant from its obligation to pay all outstanding amounts owed for shipping labels previously generated, accrued interest, or fees.

17. Account Balance and Refund Policy

Account Balances are non-interest-bearing and may only be used to purchase shipping labels through the Platform. In the event of account closure, any remaining positive Account Balance will be refunded to the Merchant within thirty (30) days, subject to verification of the Merchant's identity and bank details, and after deduction of any amounts owed by the Merchant to Shipide. Refunds will only be issued to a bank account held in the name of the Merchant.

18. Fees and Pricing

Shipping rates displayed on the Platform may change from time to time depending on Carrier pricing or commercial conditions. The price applicable to a shipping label is the price displayed at the moment the label is generated through the Platform. For material pricing changes affecting Merchants on monthly invoicing, Shipide will provide at least fourteen (14) days prior written notice. Continued use of the Platform after such notice constitutes acceptance of the revised pricing.

19. Force Majeure

Shipide shall not be liable for any failure or delay in performance resulting from events beyond its reasonable control, including but not limited to natural disasters, strikes, pandemic-related disruptions, war, government restrictions, Carrier outages, system failures, or network disruptions. Shipide will make reasonable efforts to notify affected Merchants and to restore normal service as soon as practicable.

20. Agreement Updates

Shipide may update this Agreement from time to time to reflect changes in the law, regulatory requirements, or platform developments.

Each version of this Agreement carries a unique version identifier displayed at the top of the document.

Shipide will provide at least fourteen (14) days prior notice of any material changes, by email to the Merchant's registered address or by posting a notice on the Platform. Continued use of the Platform after the effective date of the updated Agreement constitutes acceptance of the revised terms. If the Merchant does not accept the revised terms, it may terminate its account prior to the effective date in accordance with Section 3.

21. Governing Law and Jurisdiction

This Agreement shall be governed by and interpreted in accordance with the laws of Belgium. Any disputes arising out of or in connection with this Agreement shall be submitted to the exclusive jurisdiction of the courts of Brussels, Belgium.

22. Miscellaneous

If any provision of this Agreement is found to be invalid, illegal, or unenforceable, such provision shall be modified to the minimum extent necessary to make it enforceable, and the remaining provisions shall remain in full force and effect.

The failure of Shipide to enforce any provision of this Agreement shall not be construed as a waiver of its right to enforce such provision in the future.

This Agreement constitutes the entire agreement between the Parties with respect to its subject matter and supersedes all prior agreements, representations, or understandings relating to the same subject matter.

23. Contact and Notices

All notices, requests, or communications under this Agreement shall be sent to Cryvelin LLC via email at legal@shipide.com or by post to the registered address of Cryvelin LLC as notified to the Merchant. Notices to the Merchant will be sent to the email address registered on the Merchant's Shipide account.

This agreement was accepted electronically by the following party on the date and time recorded below. This record constitutes a legally binding acceptance of the Shipide Shipping Platform Services Agreement and serves as confirmation of the terms in force at the time of account creation.

Merchant Legal Name:
Merchant VAT/Tax ID:
Merchant Billing Address:
Merchant Contact Name (Authorised Representative):
Merchant Email Address:
Merchant IP Address:
Merchant Internal Customer ID:
Record ID:
Date and Time of Acceptance:
Agreement Version Accepted:
Acceptance Method: Clickwrap - explicit checkbox confirmation at account registration

This document was automatically generated and delivered to the Merchant immediately following acceptance and registration. No handwritten signature was required; acceptance was recorded electronically. Acceptance of this Agreement was recorded by Cryvelin LLC systems at the date and time stated above and constitutes a binding agreement between the Merchant and Cryvelin LLC under the laws of Belgium.
$contract$,
  encode(digest($contract$
Commercial Agreement

By clicking "I Agree", creating an account, or using the Platform, the Merchant agrees to be bound by the terms of this Agreement.

By accepting, you confirm you are at least 18 years old and have the legal authority to bind the Merchant entity you represent to this Agreement.

Shipide records the date, time, IP address, account identity, and version of this Agreement accepted at the time of account creation or acceptance.

This Shipping Platform Services Agreement ("Agreement") governs access to and use of the Shipide platform and related services. Shipide is a commercial brand operated by Cryvelin LLC, a company incorporated under the laws of the State of Delaware, United States.

1. Definitions

For the purposes of this Agreement, the following definitions apply.

"Cryvelin LLC" means Cryvelin LLC, a limited liability company organized under the laws of the State of Delaware, United States, which operates the Shipide platform.

"Shipide" refers to the commercial brand and platform operated by Cryvelin LLC.

"Platform" means the Shipide online platform enabling Merchants to generate shipping labels and access shipping services.

"Merchant" means the business entity or individual registering for and using the Shipide Platform.

"Carrier" means any third-party shipping or logistics provider whose services are accessible through the Platform.

"Shipping Label" means any shipping label generated via the Platform for the transport of a parcel.

"Account Balance" means the prepaid funds deposited by the Merchant into its Shipide account.

"Invoice" means a billing document issued by Shipide for shipping labels generated during a billing period.

"Aggregated Data" means anonymised or pseudonymised shipment data, rate benchmarks, and logistics metrics derived from Merchant activity and used by Shipide for platform optimization, carrier negotiations, and analytics, as further described in Section 13.

"GDPR" means the EU General Data Protection Regulation (Regulation (EU) 2016/679).

2. Scope of Services

Shipide provides a technology platform that allows Merchants to access shipping services and generate shipping labels with third-party carriers at negotiated rates.

Shipide aggregates shipping volumes across multiple Merchants in order to obtain preferential shipping rates from carriers.

Shipide does not operate as a carrier and does not physically transport parcels. All shipments generated through the Platform are transported and handled by independent third-party Carriers.

3. Freedom of Use and Termination

The Merchant remains free at all times to decide whether or not to use the Shipide Platform.

This Agreement does not impose any minimum usage obligations, exclusivity requirements, or long-term commitments on the Merchant. The Merchant may stop using the Platform at any time without notice and without penalty.

The Merchant may terminate its account at any time by providing written notice to Shipide or by ceasing use of the Platform.

Termination or non-use of the Platform does not create any fees or penalties, provided that all outstanding amounts related to shipping labels previously generated through the Platform have been fully paid. Any unused prepaid Account Balance remaining at the time of account closure will be refunded to the Merchant within thirty (30) days, subject to deduction of any outstanding amounts owed.

4. Merchant Account

To access the Platform, the Merchant must create an account and provide accurate, complete, and up-to-date information during registration. The Merchant is responsible for maintaining the confidentiality of its account credentials and remains fully responsible for all activities conducted through its account. Shipide reserves the right to suspend or terminate accounts that contain inaccurate, misleading, or fraudulent information, or that are used in violation of this Agreement or applicable laws.

5. Generation and Purchase of Shipping Labels

Shipping labels may be generated through the Platform using available Carrier services. Each shipping label generated through the Platform constitutes a binding purchase by the Merchant. All shipping labels generated remain payable by the Merchant regardless of whether the shipment is ultimately used or cancelled after generation. The Merchant is responsible for ensuring the accuracy of all shipment information, including but not limited to weight, dimensions, destination, and shipment contents. Any additional charges imposed by Carriers due to inaccurate shipment information, adjustments, or re-classification of parcels will be charged to the Merchant.

6. Payment Methods

Shipide provides several payment methods for the purchase of shipping labels, including prepaid account balances, card payments, and monthly invoicing for selected Merchants.

6.1 Prepaid Account Balance — Merchants may fund their Shipide account by transferring funds via bank transfer to the IBAN designated by Shipide. Funds received will be credited to the Merchant's Account Balance and used to pay for shipping labels generated on the Platform. The cost of each label will be deducted from the Account Balance at the time of generation. Shipide reserves the right to suspend the ability to generate labels if the Account Balance is insufficient to cover the cost of the requested label.

6.2 Card Payments — Merchants may also register a debit or credit card within their account. Shipide may charge the registered payment method automatically for shipping labels generated or to replenish the Merchant's Account Balance. By providing a payment card, the Merchant authorizes Shipide to charge the payment method on file for all amounts due under this Agreement.

6.3 Monthly Invoicing — Monthly invoicing may be granted to selected Merchants at the sole discretion of Shipide. Eligibility for invoicing may depend on shipping volume, payment history, or other creditworthiness criteria determined by Shipide. Merchants granted monthly invoicing will receive invoices summarizing shipping labels generated during the billing period. Unless otherwise specified, invoices are payable within fourteen (14) days from the invoice date. Shipide reserves the right to revoke invoicing privileges at any time, with reasonable prior notice where commercially practicable.

7. Billing, Late Payment, and Penalties

Prompt and complete payment is essential to the operation of the Platform. Shipide operates on narrow margins with carrier partners, and unpaid balances directly affect its ability to maintain negotiated rates for all Merchants. The following provisions are strictly enforced.

7.1 Payment Deadlines — All invoices issued under this Agreement are due and payable within fourteen (14) calendar days from the invoice date, unless a different payment term has been agreed in writing between Shipide and the Merchant. Prepaid balance topups are due immediately upon request or as required to maintain a positive Account Balance.

7.2 Late Payment Interest — Any amount not received by Shipide on or before the due date will automatically and without prior notice accrue late payment interest at a rate of one and a half percent (1.5%) per month (equivalent to eighteen percent (18%) per annum), calculated from the due date until the date of full payment. This rate applies both to invoiced amounts and to carrier adjustment charges. Late payment interest will be added to the Merchant's outstanding balance and will appear on the next invoice issued.

7.3 Administrative Recovery Fee — In addition to accrued interest, any invoice that remains unpaid for more than thirty (30) calendar days past its due date will be subject to an administrative recovery fee of EUR 75 (or the equivalent in the applicable currency) per overdue invoice. This fee reflects the administrative burden of debt follow-up and is in addition to any late interest accrued under Section 7.2.

7.4 Suspension for Non-Payment — If any amount owed by the Merchant remains unpaid for more than seven (7) calendar days past its due date, Shipide may, without prior notice and without liability, immediately suspend the Merchant's access to the Platform, including the ability to generate new shipping labels. Suspension shall remain in effect until all outstanding amounts, including accrued interest and fees, have been paid in full.

7.5 Termination for Non-Payment — If payment remains outstanding for more than sixty (60) calendar days past its due date, Shipide reserves the right to permanently terminate the Merchant's account. Termination for non-payment does not extinguish the Merchant's obligation to pay all outstanding amounts, accrued interest, fees, and any collection costs incurred by Shipide, including reasonable legal and debt recovery expenses.

7.6 Debt Recovery Costs — The Merchant shall be liable for all reasonable costs incurred by Shipide in recovering any unpaid amounts, including but not limited to legal fees, court costs, and the fees of third-party debt collection agencies. Shipide may report persistent non-payment to credit reference agencies or commercial credit bureaux, which may affect the Merchant's credit profile.

7.7 Disputed Invoices — If the Merchant disputes any item on an invoice, it must notify Shipide in writing within seven (7) calendar days of receipt of the invoice, specifying the disputed amount and the grounds for the dispute. Undisputed portions of the invoice remain due and payable on the original due date. Failure to dispute an invoice within the prescribed period constitutes acceptance of the invoice in full.

7.8 Set-Off — The Merchant may not withhold or set off any amounts owed to Shipide against any claimed liability of Shipide to the Merchant, whether under this Agreement or otherwise, without the prior written consent of Shipide.

8. Credit Limits

Shipide may establish credit limits for Merchants who are granted monthly invoicing privileges. If a Merchant exceeds its assigned credit limit, Shipide may suspend the ability to generate new shipping labels until outstanding balances are settled. Shipide reserves the right to adjust or modify credit limits at its discretion based on usage, payment behavior, or risk considerations, with reasonable prior notice where commercially practicable.

9. Carrier Terms

All shipments generated through the Platform remain subject to the terms and conditions of the selected Carrier. The Merchant agrees to comply with all Carrier rules, including but not limited to restrictions on prohibited goods, packaging requirements, customs declarations, weight limits, and dimensional restrictions. In the event of any conflict between this Agreement and a Carrier's shipping conditions regarding the handling of shipments, the Carrier's terms shall prevail with respect to transport services.

10. Merchant Responsibilities

The Merchant is responsible for providing accurate shipment information and ensuring that all shipments comply with applicable laws and Carrier requirements. The Merchant must ensure that parcels are properly packaged and labeled in accordance with Carrier guidelines.

The Merchant further agrees not to ship any of the following categories of goods through the Platform: explosives, ammunition, weapons, or items regulated under firearms legislation; narcotics, controlled substances, or illegal drugs; flammable, corrosive, radioactive, toxic, or otherwise hazardous materials; lithium batteries or dangerous goods not properly declared and packaged per IATA/IMDG regulations; perishable items not authorized by the relevant Carrier; counterfeit or intellectual property-infringing goods; human remains, live animals, or biological specimens not authorized by law; and any items prohibited under applicable local, national, or international laws.

The Merchant remains solely responsible for the contents of all shipments generated through its account. Shipide assumes no responsibility for the legality or nature of items shipped by the Merchant.

11. Indemnification

The Merchant agrees to indemnify, defend, and hold harmless Cryvelin LLC, its officers, employees, agents, and partners from and against any and all claims, damages, fines, penalties, liabilities, costs, and expenses (including reasonable legal fees) arising from or related to the Merchant's shipments including any illegal, dangerous, or prohibited goods transported through the Platform.

12. Platform Role, Liability Cap, and Warranties

12.1 Platform Role — Shipide acts solely as a technology platform facilitating access to shipping services offered by third-party Carriers. Shipide is not a Carrier and does not physically transport shipments. Accordingly, Shipide shall not be liable for shipment delays, loss or damage to parcels, Carrier service disruptions, customs issues, delivery failures, or any other events arising from the transportation process. Any claims related to shipment loss or damage must be handled according to the claims procedures established by the relevant Carrier.

12.2 No Warranty — The Platform is provided "as is" and "as available", without warranties of any kind, express or implied, including but not limited to warranties of merchantability, fitness for a particular purpose, or uninterrupted availability. Shipide does not warrant that the Platform will be error-free or that it will be available at all times.

12.3 Liability Cap — To the maximum extent permitted by applicable law, Cryvelin LLC's total aggregate liability to the Merchant arising out of or in connection with this Agreement, whether in contract, tort (including negligence), or otherwise, shall not exceed the total amounts paid by the Merchant to Cryvelin LLC in the three (3) calendar months immediately preceding the event giving rise to the claim.

This cap applies regardless of the nature of the claim or the number of incidents giving rise to liability.

12.4 Exclusion of Consequential Loss — In no event shall Cryvelin LLC be liable for any indirect, incidental, special, consequential, or punitive damages, including but not limited to loss of revenue, loss of profits, loss of data, or loss of business, even if Cryvelin LLC has been advised of the possibility of such damages.

13. Data Governance, Privacy, and Shipment Data Pooling

13.1 Personal Data and GDPR — Cryvelin LLC processes personal data in connection with the provision of the Platform in accordance with the GDPR and all other applicable data protection laws.

The categories of personal data processed may include Merchant business contact details, recipient names and addresses, and transactional data related to shipments.

Merchants established in the EEA, or who process personal data of EEA-resident individuals through the Platform, acknowledge that Cryvelin LLC acts as a data processor in respect of shipment recipient data, and as a data controller in respect of Merchant account data.

The terms governing such processing, including the technical and organisational measures applied by Cryvelin LLC, the lawful bases for processing, sub-processor relationships, data subject rights procedures, and breach notification commitments, are set out in full in the Shipide Privacy Policy available at [Privacy Policy URL].

The Merchant's acceptance of this Agreement constitutes acceptance of those processing terms.

Merchants who require additional information about how their data is handled may contact Cryvelin LLC at legal@shipide.com.

13.2 Shipment Data Pooling and Platform Analytics — By using the Platform, the Merchant expressly acknowledges and consents to Shipide's right to collect, retain, and use Aggregated Data derived from the Merchant's shipment activity for the following purposes: negotiating and maintaining preferential rates with Carriers on behalf of all Platform users; optimizing routing, carrier selection, and delivery performance across the Platform; generating anonymised market intelligence and logistics analytics for internal use; improving the Platform's features, pricing models, and service quality; and benchmarking carrier performance and service levels.

Aggregated Data used for these purposes will be anonymised or pseudonymised in accordance with GDPR standards and will not be attributed to any individual Merchant or used to identify the Merchant's commercial relationships, customer identities, or business strategies. The Merchant retains ownership of its own raw shipment records as maintained in its account.

13.3 Data Retention — Shipide will retain Merchant shipment data for a period of five (5) years from the date of the relevant transaction, or for such longer period as may be required by applicable tax, customs, or regulatory law. After expiry of the retention period, data will be securely deleted or irreversibly anonymised.

13.4 Data Security — Shipide implements appropriate technical and organisational security measures to protect Merchant data against unauthorised access, accidental loss, disclosure, or destruction. These measures include encrypted data transmission, access controls, and regular security assessments. In the event of a personal data breach affecting Merchant data, Shipide will notify the Merchant without undue delay and, where required, notify the relevant supervisory authority in accordance with GDPR Article 33.

13.5 Merchant Obligations Regarding Data — The Merchant warrants that it has obtained all necessary consents and has a lawful basis under applicable data protection law for providing recipient personal data to Shipide for the purpose of generating shipping labels. The Merchant shall indemnify Shipide for any regulatory fines, penalties, or third-party claims arising from the Merchant's failure to comply with this obligation.

13.6 Data Portability and Deletion — Upon written request, Shipide will provide the Merchant with an export of its own account and shipment data in a standard machine-readable format. Upon account termination, Shipide will, upon request, delete or anonymise the Merchant's personally identifiable account data within ninety (90) days, subject to any legal retention obligations.

14. Intellectual Property

The Shipide Platform, software, brand, trademarks, API, documentation, and all related intellectual property rights are and shall remain the exclusive property of Cryvelin LLC. This Agreement does not transfer any intellectual property rights to the Merchant. The Merchant is granted a limited, non-exclusive, non-transferable, revocable licence to access and use the Platform solely for the purposes set out in this Agreement. The Merchant shall not reproduce, reverse engineer, decompile, or create derivative works of any part of the Platform without the prior written consent of Cryvelin LLC.

15. API, Integrations, and Automation

Where Shipide makes an API, Shopify integration, CSV import functionality, or other automated tools available to the Merchant, such tools are provided "as is" and subject to the terms of this Agreement. Shipide shall not be liable for errors, interruptions, or data losses arising from the use of such integrations. The Merchant is responsible for ensuring that any data transmitted through automated integrations is accurate and compliant with this Agreement. Shipide may impose rate limits, access controls, or technical restrictions on API usage to maintain Platform stability.

16. Suspension and Termination

Shipide may suspend or terminate access to the Platform immediately and without prior notice if the Merchant fails to pay any amounts due within the timeframes specified in Section 7; exceeds its assigned credit limit and fails to remedy within 48 hours of notification; engages in fraudulent, abusive, or unlawful activity; ships or attempts to ship prohibited or illegal goods; or materially breaches any term of this Agreement and, where the breach is remediable, fails to remedy it within seven (7) days of written notice.

The Merchant may terminate its account at any time by providing written notice to Shipide. Termination of the account does not release the Merchant from its obligation to pay all outstanding amounts owed for shipping labels previously generated, accrued interest, or fees.

17. Account Balance and Refund Policy

Account Balances are non-interest-bearing and may only be used to purchase shipping labels through the Platform. In the event of account closure, any remaining positive Account Balance will be refunded to the Merchant within thirty (30) days, subject to verification of the Merchant's identity and bank details, and after deduction of any amounts owed by the Merchant to Shipide. Refunds will only be issued to a bank account held in the name of the Merchant.

18. Fees and Pricing

Shipping rates displayed on the Platform may change from time to time depending on Carrier pricing or commercial conditions. The price applicable to a shipping label is the price displayed at the moment the label is generated through the Platform. For material pricing changes affecting Merchants on monthly invoicing, Shipide will provide at least fourteen (14) days prior written notice. Continued use of the Platform after such notice constitutes acceptance of the revised pricing.

19. Force Majeure

Shipide shall not be liable for any failure or delay in performance resulting from events beyond its reasonable control, including but not limited to natural disasters, strikes, pandemic-related disruptions, war, government restrictions, Carrier outages, system failures, or network disruptions. Shipide will make reasonable efforts to notify affected Merchants and to restore normal service as soon as practicable.

20. Agreement Updates

Shipide may update this Agreement from time to time to reflect changes in the law, regulatory requirements, or platform developments.

Each version of this Agreement carries a unique version identifier displayed at the top of the document.

Shipide will provide at least fourteen (14) days prior notice of any material changes, by email to the Merchant's registered address or by posting a notice on the Platform. Continued use of the Platform after the effective date of the updated Agreement constitutes acceptance of the revised terms. If the Merchant does not accept the revised terms, it may terminate its account prior to the effective date in accordance with Section 3.

21. Governing Law and Jurisdiction

This Agreement shall be governed by and interpreted in accordance with the laws of Belgium. Any disputes arising out of or in connection with this Agreement shall be submitted to the exclusive jurisdiction of the courts of Brussels, Belgium.

22. Miscellaneous

If any provision of this Agreement is found to be invalid, illegal, or unenforceable, such provision shall be modified to the minimum extent necessary to make it enforceable, and the remaining provisions shall remain in full force and effect.

The failure of Shipide to enforce any provision of this Agreement shall not be construed as a waiver of its right to enforce such provision in the future.

This Agreement constitutes the entire agreement between the Parties with respect to its subject matter and supersedes all prior agreements, representations, or understandings relating to the same subject matter.

23. Contact and Notices

All notices, requests, or communications under this Agreement shall be sent to Cryvelin LLC via email at legal@shipide.com or by post to the registered address of Cryvelin LLC as notified to the Merchant. Notices to the Merchant will be sent to the email address registered on the Merchant's Shipide account.

This agreement was accepted electronically by the following party on the date and time recorded below. This record constitutes a legally binding acceptance of the Shipide Shipping Platform Services Agreement and serves as confirmation of the terms in force at the time of account creation.

Merchant Legal Name:
Merchant VAT/Tax ID:
Merchant Billing Address:
Merchant Contact Name (Authorised Representative):
Merchant Email Address:
Merchant IP Address:
Merchant Internal Customer ID:
Record ID:
Date and Time of Acceptance:
Agreement Version Accepted:
Acceptance Method: Clickwrap - explicit checkbox confirmation at account registration

This document was automatically generated and delivered to the Merchant immediately following acceptance and registration. No handwritten signature was required; acceptance was recorded electronically. Acceptance of this Agreement was recorded by Cryvelin LLC systems at the date and time stated above and constitutes a binding agreement between the Merchant and Cryvelin LLC under the laws of Belgium.
$contract$, 'sha256'), 'hex'),
  '/assets/contracts/commAgreement-v1.pdf',
  now(),
  false,
  'Seeded by supabase_clickwrap.sql'
)
on conflict (version) do update
set
  title = excluded.title,
  body_text = excluded.body_text,
  hash_sha256 = excluded.hash_sha256,
  pdf_url = excluded.pdf_url,
  effective_at = excluded.effective_at,
  notes = excluded.notes;

update public.clickwrap_contract_versions
set pdf_url = '/assets/contracts/commAgreement-v1.pdf'
where coalesce(pdf_url, '') = '';

update public.clickwrap_contract_versions
set is_active = false
where is_active;

update public.clickwrap_contract_versions
set is_active = true
where version = 'v2.0';
