# GoHighLevel / LeadConnector Integration

This project already includes a GHL branch in n8n.

## Current GHL flow

1. `Enviar GHL if` checks whether the item should be sent to GHL.
2. `Preparar Payload GHL` builds contact, opportunity and note payloads.
3. `GHL Upsert Contact` calls `https://services.leadconnectorhq.com/contacts/upsert`.
4. `Preparar Opportunity GHL` takes the returned contact ID and prepares the opportunity payload.
5. `GHL Create Opportunity` creates the opportunity in the selected pipeline.
6. `GHL Add Note` adds the conversation note to the contact.

## Required GHL variables

```bash
GHL_ACCESS_TOKEN=
GHL_LOCATION_ID=
GHL_PIPELINE_ID=
GHL_STAGE_NUEVO_LEAD=
GHL_STAGE_AGENDADO=
GHL_STAGE_ESCALADO=
GHL_STAGE_CANCELADO=
```

## Recommended pipeline for an aesthetic clinic

- New lead
- Qualified
- Appointment requested
- Appointment booked
- Human follow-up
- No-show
- Cancelled
- Converted client

## Recommended custom fields

- Preferred treatment
- Source channel
- Conversation ID
- Last AI stage
- Appointment date
- Calendar event ID
- Human takeover reason
