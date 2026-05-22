# LabPlus Invoice System — V1

A procedural VBA prototype that replaced a fully manual Word-based invoicing
process at a laboratory LIMS. The first working version of a system that
would be rebuilt twice before reaching its current architecture.

---

## The Problem It Solved

The lab was generating invoices by opening an old Word document, saving it
as a new file, and editing every field by hand. Time-consuming and
error-prone. Several other lab systems had already been automated, so an
invoice automation was a natural next step — though the owner was skeptical
it could be done.

V1 proved it could.

---

## How It Worked

The codebase was built around array manipulation — a technique I had just
learned and was keen to apply. The approach was:

1. Loop through the Excel testing-request forms and load data into arrays
2. Loop back through the arrays and write output to a Word document from a template

It worked. Significantly faster and more accurate than manual entry.

---

## Where It Broke Down

The owner was impressed enough to request support for a specific customer
that used a completely different invoice format. Without the tools to model
that variation cleanly, the solution was more loops with `If...Then` forks
branching between the standard path and the special case.

That worked too — but the cost was visible immediately. Any modification
required disentangling the loops and forks before touching the actual logic.
The code had no seams.

---

## Code Example

```vba
' Data processing and output mixed together
For i = 2 To LastRow
    InvoiceTotal = InvoiceTotal + Cells(i, "D").Value
Next i

' Direct output to Word — no separation from calculation
Set doc = CreateObject("Word.Application")
doc.Documents.Add
doc.Content.Text = "Invoice Total: " & InvoiceTotal
```

Any change to the output format required touching the same code that
calculated the totals. There was no boundary between the two concerns.

---

## Characteristics

- **Procedural** — loops and arrays handled all data processing
- **Output-coupled** — arrays were shaped around what the Word document needed
- **No separation of concerns** — data loading, calculation, and output lived
  in the same place
- **Fragile to change** — adding the special-case customer format required
  forking core logic rather than extending it

---

## What It Led To

V1 demonstrated that automation was viable and worth investing in. It also
made the maintenance problem concrete — not as a theoretical concern, but as
something felt every time a change came in.

V2 introduced classes to address this. That turned out to be necessary but
not sufficient. See the [V2 README](../V2/README.md) for what happened next.
