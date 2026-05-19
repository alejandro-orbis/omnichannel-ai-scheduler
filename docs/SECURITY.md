# Security

Never commit real access tokens, credential IDs, webhook verification tokens, patient data, exported executions, pin data or production calendar IDs.

Before publishing a workflow export, search for:

- `access_token`
- `Bearer `
- `credentials`
- `pinData`
- real email addresses
- real phone numbers
- real patient names
- private webhook URLs
