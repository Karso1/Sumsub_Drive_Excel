# Security and sensitive-data handling

This project can process KYC, screening, and identity-report PDFs. Treat all inputs and outputs as confidential.

Never commit or upload:

- `credentials.json`, `token.json`, or browser profiles;
- Excel input/output containing personal data;
- PDFs, manifests, logs, Drive file IDs, or production folder IDs;
- company-specific URLs, client IDs, or internal paths.

Use a private repository unless legal, security, and data-protection teams approve public release. Keep the Drive folder restricted to the minimum necessary audience. Do not use `--share-anyone-with-link` without explicit authorization.
