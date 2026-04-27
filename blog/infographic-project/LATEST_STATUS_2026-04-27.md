# Latest Status: 2026-04-27

This file supersedes older operational assumptions in `RUNBOOK.md` where they conflict.

## Verified current state

- Git remote status: local `main` and `origin/main` are identical.
- Checked on: 2026-04-27
- Result: no pending upstream commits to pull into this workspace.

## Dependency status

The primary dependencies used by this workflow are already at their current versions as of 2026-04-27:

- `openai==2.32.0`
- `pillow==12.2.0`
- `pytest==9.0.3`
- `pytest-cov==7.1.0`
- `pptxgenjs==4.0.1`

## GPT-5.5 requirement

When using the Codex or ChatGPT-assisted route for prompt generation, review, or image-direction work, use `GPT-5.5` as the default model.

Operationally, this means:

- Any older references to `ChatGPT Plus` in `RUNBOOK.md` should be read as `use GPT-5.5`.
- The Codex/manual route remains the preferred route.
- The OpenAI API route in this project remains the fallback route for image generation and still uses `gpt-image-1`.

## Practical guidance

- Use `python -m scripts.export_for_codex --date <today>` for the preferred Codex/GPT-5.5 route.
- Use `python -m scripts.generate_images --date <today>` only when the API fallback route is needed.

