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

Docform renders templates with Jinja2, so a template can run expressions in your Python process. That's expected behavior, not a vulnerability. If you render templates written by people you don't trust, use a sandboxed environment:

```python
from jinja2.sandbox import SandboxedEnvironment

doc.render(context, jinja_env=SandboxedEnvironment(autoescape=True))
```

A way to escape the sandbox through Docform's own code would be a vulnerability, and should be reported as above.
