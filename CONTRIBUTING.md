# Contributing

Thanks for considering a contribution! This is a small, focused tool, so
the process is intentionally lightweight.

## Reporting a bug

Open an issue with:
- What you expected to happen
- What actually happened (exact error message/screenshot if there is one)
- Your `config.json`'s `hsn` section **with any real emails/passwords
  removed** — the `key_col`/`compare_cols`/`sheets` structure is usually
  the relevant part
- Whether it's a first-ever run (baseline) or a later check

## Suggesting a feature

Open an issue describing the use case, not just the feature — it helps
to know *why* you need it, since there's sometimes already a way to do
it (e.g. the difference between the "delta only" and "full list"
validation scans covers a lot of ground).

## Submitting a change

1. Fork the repo, create a branch off `main`.
2. Keep changes focused — one logical change per PR is much easier to
   review than several unrelated ones bundled together.
3. Test manually before submitting: run through at least a baseline
   check and a second check that produces a real delta, and confirm the
   dashboard shows what you'd expect.
4. Open a PR describing what changed and why.

## Code style

- Plain, readable Python — this project intentionally avoids heavy
  frameworks or abstractions beyond what Flask/pandas already provide.
- Comments explain *why*, not just *what*, especially around anything
  non-obvious (e.g. why HSN and SAC are never cross-compared, why
  invisible characters are checked in two different scopes).
- No new dependencies without a good reason — keeping `requirements.txt`
  minimal is a deliberate goal (see the "$0 cost" framing in the README).

## Things to be careful about

- **Never commit `config.json` or `subscribers.json`** — both are
  gitignored for a reason (credentials and personal data respectively).
- If you change the GST portal URL or scraping logic, be mindful this
  hits a real government website — don't introduce aggressive polling or
  remove the retry/backoff behavior in `hsn_core.py`.
