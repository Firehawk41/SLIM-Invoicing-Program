# SLIM Invoice System — V2

A class-based VBA invoice system supporting three distinct invoice formats
via a shared base class and interface. A meaningful step forward from V1 —
and the version whose failure directly motivated V3.

\---

## What Changed From V1

V1 was procedural: loops, arrays, no structure. It worked but had no seams.
Any change required disentangling everything.

V2 introduced proper OOP. The three invoice formats — a standard individual
report invoice, a weekly summary for Customer A, and a monthly summary for
Customer B — each had unique layout requirements. The solution was a shared
`clsInvoiceBase` class handling common logic, with `IInvoice` as the
interface contract that each format implemented independently.

This was the right instinct. Extracting shared behaviour into a base class
(DRY), defining a contract via interface, separating the three output
variants — all sound design moves. The system worked well and handled
real production load.

\---

## The Interface

```vba
' IInvoice — implemented by clsInvoiceDefaultIndividual,
'            clsInvoiceCustomerASummary, clsInvoiceCustomerBSummary
Public Property Get CustomerName() As String
End Property
Public Property Get WordInvoice() As Word.Document
End Property
Public Property Get FileName() As String
End Property
Public Sub Initialize(Submissions As Collection, SequenceNumber As Long, \_
    Logger As clsLoggingSystem, AccessDB As clsAccessDatabase, \_
    Factory As clsInvoiceFactory, PricingRetriever As clsInvoicePriceRetriever, \_
    SLIMDict As clsSLIMDictionaryManager)
End Sub
Public Sub CreateWordInvoice(WordApp As Word.Application, \_
    InvoiceWriter As clsInvoiceWriter)
End Sub
Public Sub AddInvoiceToDatabase()
End Sub
```

The problem is visible in the interface itself: `WordInvoice As Word.Document`
is part of the contract. Word output wasn't an implementation detail — it was
baked into the definition of what an invoice *was*.

\---

## Where It Broke Down

The system ran in production until the output requirement changed: Word
documents out, CSV in, with a different field structure.

Because `Word.Document` was embedded in the interface, there was no clean
path to a new output format. The domain objects had been shaped around their
writer from the start — customer data, pricing logic, and payment terms were
all organised around what the Word template needed. Swapping the writer meant
refactoring the domain.

There were no pure domains. Business logic was distributed across the class
tree with no clear layer boundaries, which made the required changes harder
than they needed to be.

\---

## Characteristics

* **Interface-based** — `IInvoice` implemented by three format-specific classes
* **Shared base class** — `clsInvoiceBase` extracted common logic (DRY)
* **Output-coupled** — `Word.Document` embedded in the interface contract
* **No layer separation** — pricing, customer, and output logic interleaved
throughout the class tree
* **Fragile to output change** — swapping the writer required refactoring
the domain

\---

## What It Led To

Rather than patch V2 a second time, the CSV requirement became the forcing
function for a principled rebuild. V3 was designed from the start with a
clean boundary between domain objects and output — the writer reads from
domain objects, it has no influence on their structure.

Each version of this system was built at the edge of what I knew at the
time, and pushed past it. V2's failure wasn't a mistake so much as the
natural limit of applying OOP without yet having the vocabulary for
separation of concerns. V3 is the direct response to that experience.

See the [V3 README](../V3/README.md) for the architecture that replaced it.

