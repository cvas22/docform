# Security Policy

## Supported versions

Security fixes go into the latest release on PyPI.

| Version | Supported |
|---|---|
| 0.21.x | Yes |
| < 0.21 | No |

## Reporting a vulnerability

Please report vulnerabilities privately through GitHub: go to the [Security tab](https://github.com/cvas22/docform/security) and click "Report a vulnerability". Don't open a public issue.

Include the Docform version, a description of the problem, and a minimal template and script that reproduce it if you can. You should hear back within a week. Once a fix is released, the advisory is published with credit to the reporter unless you'd rather stay anonymous.

## Templates are code

Docform renders templates with Jinja2, and by default it uses Jinja2's `SandboxedEnvironment`, so a template can't reach Python internals. If you pass your own `jinja_env`, Docform uses it as is, so make it a `SandboxedEnvironment` when templates come from people you don't trust:

```python
from jinja2.sandbox import SandboxedEnvironment

doc.render(context, jinja_env=SandboxedEnvironment(autoescape=True))
```

Rendering with a plain `jinja2.Environment()` or the `--unsafe` command line flag turns the sandbox off. Code run by a template in that mode is expected behavior, not a vulnerability.

A way to escape the default sandbox through Docform's own code would be a vulnerability, and should be reported as above.
