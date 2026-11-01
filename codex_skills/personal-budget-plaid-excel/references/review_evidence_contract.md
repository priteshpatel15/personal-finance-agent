# Review Evidence Contract

Purchase research is optional, read-only, local, and applied only when the evidence confidently identifies one transaction. Do not log manifest contents.

Use a JSON list or an object with an `evidence` list. Each record has exact transaction match fields plus an evidence disposition:

```json
{
  "evidence": [
    {
      "match": {
        "date": "2026-08-04",
        "merchant": "Example Merchant",
        "amount": "12.34",
        "account": "Optional visible account name",
        "occurrence": 1
      },
      "status": "confirmed",
      "source": "gmail",
      "summary": "Concise receipt purpose",
      "items": ["Item one", "Item two"],
      "gmail_thread_id": "actual-connector-returned-id"
    }
  ]
}
```

- `date`, `merchant`, and `amount` are required and must match the transaction exactly. `account` narrows a match. Use `occurrence` only to disambiguate otherwise identical visible transactions after reviewing them in order.
- Supported status values are `confirmed`, `ambiguous`, `unmatched`, and `review_needed`. Only confirmed evidence can produce item detail.
- Supported sources are `gmail` and `amazon`. Confirmed Gmail evidence requires the actual thread/message identifier returned by the connector; the workflow builds the Gmail URL from that identifier. Never fabricate or repair an identifier.
- Confirmed evidence needs a useful summary or item list. Invalid, ambiguous, conflicting, missing, or non-unique evidence becomes `Review needed` without a hyperlink.
- Keep the manifest on the local machine with restricted access and delete it after successful workbook validation unless the user explicitly asks to retain it.
