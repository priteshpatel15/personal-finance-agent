# Automatic categorization with owner review

**Feature ID:** `FEATURE-CATEGORIZATION`

Incoming transactions are normalized and classified using merchant, Plaid category, amount, cadence, account, and the imported labeled examples. Exact owner rules and historical merchant matches outrank the local supervised model. The output is a bridged budget category, confidence, reason codes, and a review status.

Only high-confidence predictions enter the budget aggregate automatically. Low confidence, conflicting evidence, split/transfer/credit-card-payment candidates, and `Other / Needs Review` remain visible in the review queue. An owner correction writes an auditable rule and retrains only the local model. Imported `Cancel?`, `Refund?`, and `Won't Recur?` flags inform history but never automatically suppress a new live transaction.
