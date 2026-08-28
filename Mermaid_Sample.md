# Mermaid Sample

This sample shows the three most common Mermaid diagram kinds as fenced
code blocks. Press Alt+V in EdSharp to insert any of them from the
shipped snippets, answer the prompts, and preview the document in the
web browser to see the drawn diagrams. The block text itself is the
accessible view: each diagram reads as plain structured text.

## Flowchart

```mermaid
flowchart TD
    A[Start] --> B{Is it ready?}
    B -->|Yes| C[Ship it]
    B -->|No| D[Fix it]
    D --> A
```

## Sequence diagram

```mermaid
sequenceDiagram
    participant U as User
    participant P as Program
    U->>P: Request
    P-->>U: Response
```

## Pie chart

```mermaid
pie title Share by category
    "First" : 45
    "Second" : 30
    "Third" : 25
```
