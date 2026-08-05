# assaf-agent-skills

A marketplace of Agent Skills, each packaged as an individually installable
plugin. Top-level layout: one directory per skill, holding both the plugin
manifest and the skill itself.

```
<skill>/
  .claude-plugin/plugin.json          plugin manifest
  skills/<skill>/SKILL.md             the skill
  skills/<skill>/...                  scripts, templates, requirements.txt
```

## Versioning: change all four places together

A skill's version appears in **four** places. They must always agree. Changing
one and not the others produces releases that ship the wrong code, because CI
derives the release tag from `plugin.json` while the installed skill reports
`SKILL.md`.

For a skill named `<skill>` at version `X.Y.Z`:

1. `<skill>/skills/<skill>/SKILL.md` — `metadata.version: X.Y.Z`
2. `<skill>/.claude-plugin/plugin.json` — `"version": "X.Y.Z"`
3. `.claude-plugin/marketplace.json` — the `"version"` inside that skill's entry
   in the `plugins` array (**not** the top-level catalog version, see below)
4. the git tag — `<skill>-vX.Y.Z`, annotated (`git tag -a`), on the commit that
   makes the change

`.github/workflows/package-skills.yml` reads the version from `plugin.json`
(`jq -r '.version'`) and builds the release tag as `<skill>-v<version>`. If an
older release with that tag already exists it does `gh release upload --clobber`,
so a stale `plugin.json` silently republishes new code onto an old release
instead of failing. Treat `plugin.json` as the number CI trusts, and keep the
other three in step with it.

The **top-level `version` in `marketplace.json` is separate** — it versions the
catalog, not any skill. Bump it when the catalog itself changes (a skill added
or removed, an entry's description or tags rewritten), not on every skill
release.

Use semver: patch for bug fixes and documentation, minor for new capability,
major for a breaking change to a skill's inputs or output format. Ask before
bumping, and confirm which component.

## Skill metadata

Every skill's `SKILL.md` frontmatter carries:

```yaml
metadata:
  version: X.Y.Z
  author: Assaf Tzur-El
```

## Dependencies

Each skill declares its Python dependencies in
`<skill>/skills/<skill>/requirements.txt`, and its `SKILL.md` "How to run it"
section repeats them as a `pip install ... --break-system-packages` line.
Declare a package the script imports directly even when it would arrive
transitively (for example `lxml`, which `python-pptx` pulls in but
`extract.py` imports itself). `uv` was considered and rejected: keeping a
`requirements.txt` fallback for environments without `uv` means maintaining two
mechanisms for one job. Revisit if the skills standard starts assuming `uv`.

## Line endings

`.gitattributes` sets `* text=auto`, so the repository stores LF while a Windows
working tree keeps CRLF. When editing files programmatically, preserve the
existing line endings (in Python: open with `newline=''`); otherwise a two-line
change shows up as a whole-file rewrite.

## Branch

The default branch is `master`, not `main`. Workflow triggers must say `master`.
