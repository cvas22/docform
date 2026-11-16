# Contributing to Docform

Thanks for wanting to help. Bug reports, fixes, docs and new features are all welcome.

## Reporting a bug

Open an issue with the bug report template. The most useful thing you can include is a small runnable script plus the `.docx` template that shows the problem. Strip anything private out of the template first.

For security problems, please don't open a public issue. See [SECURITY.md](SECURITY.md).

## Setting up

```bash
git clone https://github.com/cvas22/docform.git
cd docform
python -m venv .venv
source .venv/bin/activate   # .venv\Scripts\activate on Windows
pip install -e ".[subdoc]" pytest
```

## Running the tests

```bash
pytest
```

The unit tests in `tests/unit` build small templates on the fly and check the rendered output. The scripts in `tests/*.py` render the sample templates in `tests/templates` into `tests/output`, which is handy for opening the result in Word:

```bash
cd tests
python runtests.py
```

CI runs both on Python 3.10 to 3.14.

## Making a change

1. Open an issue first for anything bigger than a small fix, so we can agree on the approach.
2. Create a branch from `main` and keep the change focused on one thing.
3. Add or update a test in `tests/unit` that fails without your change.
4. Update the docs in `docs/` and add a line to `CHANGES.rst` if users will notice the change.
5. Open a pull request. CI has to pass before it can be merged.

Rendering output should stay compatible with existing templates. If a change alters what an existing template renders, call it out in the pull request.

## Porting fixes from python-docx-template

Docform started as a fork of python-docx-template, and fixes that are still open there are welcome here. Keep the original author on the commit (`git cherry-pick`, or `git commit --author=...`) and link the upstream pull request.

## Code of conduct

Everyone taking part is expected to follow the [code of conduct](CODE_OF_CONDUCT.md).
