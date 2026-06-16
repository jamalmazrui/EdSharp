# Camel Type: C# Coding Guidelines

Camel Type is a coding style designed for systematic readability, optimized for
efficient navigation and review, especially by screen-reader users. It carries
type information in the identifier itself, so a reader hears `sPath`, `bFound`,
`iCount`, `lsFiles` and knows the type without seeking back to a declaration.
The rules below apply to C# (.NET Framework 4.8 and later). Equivalent
documents exist for JavaScript, VBScript, and JAWS scripting.

## 1. Hungarian type prefixes

Prefix every variable, parameter, field, and constant with a short type tag:

- `a` — array (`string[]`, `int[]`)
- `b` — boolean
- `bin` — binary buffer (`byte[]`, `Stream`)
- `dt` — date-time (`DateTime`, `TimeSpan`)
- `f` — file object (`FileStream`, `FileInfo`)
- `h` — window or OS handle (`IntPtr` from P/Invoke)
- `i` — integer (`int`, `long`, `short`, `byte`)
- `ls` — `List<T>` (use `ls`, not `l`)
- `n` — real number (`float`, `double`, `decimal`)
- `s` — string
- `d` — dictionary (`Dictionary<K,V>`)
- `hs` — hash set (`HashSet<T>`)
- `o` — **COM objects only** (`oWord`, `oExcel`, `oWorkbook`). Never use `o`
  as a generic "object" prefix.
- `v` — variant: `object` or `dynamic`, where the real type is unknown or varies

For managed class instances (your own classes, framework objects) use the
lowercase class name as the prefix, or a universally understood abbreviation
(`sb` for `StringBuilder`, `ex` for `Exception`). With one instance in scope the
class-name prefix is the whole name:

```csharp
StreamWriter writer = new StreamWriter(sPath);   // not oWriter
Form form = new Form();                           // not oForm
OpenFileDialog dialog = new OpenFileDialog();     // not oDlg
dynamic oWord = com.createApp("Word.Application"); // COM -> o is correct
```

## 2. Constants

Constants follow the **same naming rules as variables** — the Hungarian type
prefix and lower camel case, with **no special marker**. There is no `c_`
prefix (that earlier convention is obsolete). A constant is distinguished from a
variable only by the `const` or `static readonly` keyword, never by its name.

```csharp
const string sDefaultEncoding = "utf-8";
const int iTimeoutMs = 60000;
static readonly HashSet<string> hsSupportedExts =
    new HashSet<string> { ".docx", ".xlsx", ".pdf" };
```

Two further points:

- **Constants are declared and initialized on their own lines, separate from
  variables** — in their own group, immediately above the variable
  declarations of the enclosing scope. Do not interleave constant and variable
  declarations.
- **Third-party / API constants are used verbatim.** If a library or the
  framework supplies a constant (`int.MaxValue`, `Keys.Control`,
  `Environment.NewLine`, a COM enum value), use it exactly as given regardless
  of Camel Type. Camel Type governs only the constants your own code defines.

## 3. Capitalization

Lower camel case for everything the language allows: locals, parameters,
constants, fields, and — in strict Camel Type — methods too. PascalCase is used
only where the platform requires or strongly expects it:

- Class names: PascalCase (required in practice).
- Namespaces: PascalCase (required for interoperability).
- `public` methods/properties/events: a judgment call. Strict Camel Type uses
  lower camel everywhere; pragmatic Camel Type may use PascalCase on public
  surface area that external (non-Camel-Type) callers see. Pick one per file and
  be consistent.

Avoid `snake_case` unless an external API forces it.

## 4. Declarations

- Declare variables and constants at the **top of the enclosing scope**, before
  the logic.
- Group by type, one declaration statement per type, names listed
  **alphabetically** within the line, and the declaration lines themselves in
  alphabetical order by prefix. Constants form their own group above the
  variables (see section 2).

## 5. Methods, loops, resources

- Define routines as **methods** (no bare subprocedures). Prefer returning a
  value — `bool`, a count, a status — so callers can inspect or chain, even when
  the value is often ignored. `void` is fine for event handlers and clearly
  effectful operations.
- Use **single-line** form for a simple one-consequence conditional:
  `if (sPath == null) return "";`
- Prefer **foreach** over index `for` unless the index itself is needed. Avoid
  LINQ where a plain `foreach` reads more clearly aloud.
- Use `using` for any `IDisposable`, declared near the top of its scope.
- Prefer explicit types over `var` unless the right-hand side makes the type
  obvious (`new T(...)`, anonymous types, unwieldy generics).

## 6. Text files

Our own text-file writes get the "Windows treatment": UTF-8 **with BOM** and
CRLF line endings. Reads tolerate a BOM or none (a BOM-detecting reader, UTF-8
fallback). This does not apply to a user's own document saves, which honor the
encoding and line endings the user chose.

## 7. Strings and double quotes

Use double-quoted string literals. Reserve single quotes for `char` literals.
