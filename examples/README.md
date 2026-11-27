# Examples

Each folder has a template, the data it's rendered with, and a `render.py` script.

| Example | Shows |
|---|---|
| [invoice](invoice) | Variables, a table row loop with `{%tr %}`, paragraph conditions with `{%p %}`, and formatting carried over from tags |
| [multilevel list](multilevel_list) | A numbered subdoc whose child levels restart under each new parent |

Run one from the repo root:

```bash
python examples/invoice/render.py
python examples/multilevel_list/render.py
```
