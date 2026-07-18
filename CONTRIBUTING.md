# Contributing to FormaFlow

Thanks for helping improve FormaFlow.

## Important privacy rule

Never upload or commit real student records, names, phone numbers, addresses,
birth dates, identity numbers, grades, or parent details. Bug reports should use
synthetic or thoroughly anonymized documents.

## Local setup

```bash
python -m pip install -r requirements-dev.txt
python app.py
```

Open `http://127.0.0.1:8000`.

## Tests

```bash
pytest -q
```

## Pull requests

- Keep changes focused.
- Add regression tests for importer fixes.
- Describe the input structure without sharing private files.
- Confirm all tests pass.
- Explain any new automatic decision and how the user can review or override it.
