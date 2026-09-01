---
title: Feature Showcase
author: UAT
date: 2026-09-01
---

# Feature Showcase

This document exercises every supported Markdown feature so the DOCX output can
be verified by hand. Use the table of contents below to jump around, then check
each section against what Word renders.

## Table of Contents

- [Headings](#headings)
- [Inline formatting](#inline-formatting)
- [Backslash escapes](#backslash-escapes)
- [Links and autolinks](#links-and-autolinks)
- [Lists](#lists)
- [Task list](#task-list)
- [Code](#code)
- [Table](#table)
- [Blockquote](#blockquote)
- [Images](#images)
- [Line breaks](#line-breaks)

---

## Headings

The converter supports six heading levels. The five below should each map to a
distinct Word heading style (levels 5 and 6 fall back to a bold paragraph if the
template lacks those styles).

## Heading level 2

### Heading level 3

#### Heading level 4

##### Heading level 5

###### Heading level 6

## Inline formatting

Plain text with **bold**, *italic*, and ***bold italic*** runs. Underscores work
too: _italic_ and __bold__. Inline `code` uses a monospace font. Strikethrough
renders as ~~struck-through text~~.

Formatting can combine: **bold with `code` inside**, *italic with ~~strike~~*,
and ***everything at once***.

Intra-word underscores are left literal so identifiers like `date_updated` and
`snake_case_name` are not italicized.

## Backslash escapes

These should render as literal punctuation, not formatting:

- Literal asterisks: \*not italic\* and \*\*not bold\*\*
- Literal backticks: \`not code\`
- Literal tildes: \~\~not struck\~\~
- Literal brackets: \[not a link\]\(nowhere\)

## Links and autolinks

- Internal link back to [Inline formatting](#inline-formatting).
- External link: [Kiro docs](https://kiro.dev).
- Bare URL in a sentence: visit https://example.com for details.
- Angle autolink: <https://example.org>.
- A URL inside code stays literal: `https://example.net/not/a/link`.
- Mailto link: [email us](mailto:hello@example.com).

## Lists

Unordered list with a nested continuation line:

- First bullet
- Second bullet
  continuation text on the next line
- Third bullet with **bold** and a [link](#headings)

Ordered list with sub-bullets (numbering should restart at 1):

1. First step
2. Second step
   - sub-item A
   - sub-item B
3. Third step

## Task list

- [ ] Unchecked task
- [x] Checked task
- [ ] Task with **formatting** and `code`
- A normal bullet mixed into the same list

## Code

Fenced block with a language info string (the language should appear as a small
caption above the code):

```python
def greet(name: str) -> str:
    # a comment
    return f"Hello, {name}!"
```

Fenced block with no language (should render as plain monospace, no caption):

```
plain code block
  indented line
```

## Table

| Feature        | Supported | Notes                          |
|----------------|-----------|--------------------------------|
| Headings 1-6   | Yes       | Levels 5-6 may fall back       |
| Strikethrough  | Yes       | GFM `~~text~~`                 |
| Images         | Yes       | Local embed; remote is opt-in  |
| Autolinks      | Yes       | Bare and `<angle>` URLs        |

## Blockquote

> This is a blockquote. It should render with a gray vertical bar on the left.
>
> It can span multiple paragraphs and contain **bold** and *italic* text.

## Images

Embedded local image (should appear inline, scaled to fit the page width):

![A blue sample banner](assets/sample-image.png)

Missing image (should fall back to its alt text):

![this alt text should appear instead](assets/does-not-exist.png)

Remote image. By default it is not fetched, so the alt text should appear. Run
with `--fetch-remote-images` to download and embed it instead (the URL below is
a placeholder, so it will still fall back to alt text):

![remote alt fallback](https://example.com/remote.png)

## Line breaks

This line ends with two trailing spaces to force a hard break,  
so this continues on the next line within the same paragraph.

---

End of showcase.
