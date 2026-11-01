# Daily budget position and pre-purchase advice

**Feature ID:** `FEATURE-ADVICE`

The owner sees each calendar-month category budget, posted spend, pending spend, unclassified amount, remaining amount, and data-as-of time. Budget spend includes personal positive outflows and category-matched refunds; it excludes transfers, credit-card payments, and work reimbursements to prevent double counting.

For a question such as "Can I spend $75 on dinner?", the product proposes a category and then reports before/after remaining, current snapshot time, confidence, pending impact, and a clear **Yes**, **Caution**, or **No**. Yes requires a fresh snapshot, confident category, and sufficient category remaining. Caution covers a stale snapshot, uncertain category, pending/unclassified impact, or less than ten percent of the budget remaining after the purchase. No means the proposed purchase exceeds the category budget. Missing amount or category yields a concise clarification, never a fabricated decision.
