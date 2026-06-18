# assaf-agent-skills

A personal, vendor-neutral collection of [Agent Skills](https://docs.claude.com/en/docs/agents-and-tools/agent-skills/overview) built on the open `SKILL.md` standard.

A skill is just a folder: a `SKILL.md` file describing when and how to use it, plus any helper scripts. That makes these portable - usable with any AI assistant or agent that can follow the instructions and run the scripts, not tied to a single vendor. They happen to also be packaged as a [Claude Code plugin marketplace](https://code.claude.com/docs/en/plugin-marketplaces) for one-command installs, but that's only one of several ways to use them.

This is an umbrella catalog I add to over time. Each skill is cataloged as its own independently installable unit, so you can take only the ones you want.

## Skills in this collection

| Skill | What it does | Language |
| --- | --- | --- |
| [`hebrew-md-to-docx`](./hebrew-md-to-docx) | Converts Hebrew or mixed Hebrew/English Markdown into a properly formatted, RTL-aware Word `.docx`, setting per-paragraph text direction and fixing the bidi bugs you get from pasting Markdown into Word (reordered inline code, detached numbered markers, punctuation drifting to the wrong side). | Hebrew-focused; works on English too |
| [`hebrew-md-to-pptx`](./hebrew-md-to-pptx) | Converts a [slide-authoring Markdown deck](./docs/slide-markdown-format.md) (blocks separated by `---`, body/notes by `===`, presenter instructions in `[brackets]`) into an RTL-aware PowerPoint `.pptx` skeleton with real title/body placeholders, native sections, and speaker notes. The slide counterpart of `hebrew-md-to-docx`. | Hebrew-focused; works on English too |
| [`pptx-to-slide-md`](./pptx-to-slide-md) | Extracts an existing `.pptx` back into the same [slide-structured Markdown format](./docs/slide-markdown-format.md) `hebrew-md-to-pptx` consumes – so a finished deck becomes editable source again. The inverse of `hebrew-md-to-pptx`; the two round-trip (`PPTX → slide-Markdown → PPTX`). | Language-agnostic |

The two `hebrew-*` skills run on English Markdown as well, but their reason to exist is the RTL handling. `pptx-to-slide-md` is language-agnostic.

`hebrew-md-to-pptx` and `pptx-to-slide-md` share one plain-text slide-authoring format, documented in [docs/slide-markdown-format.md](./docs/slide-markdown-format.md).

All three are portable: pure Python (`python-docx` / `python-pptx` / `mistune` / `Pillow`), with no agent-specific tool dependencies, so they work in any runtime that supports the `SKILL.md` standard.

## Install / use

Pick whichever fits your setup – none is more "official" than the others.

### Manual copy (any agent runtime)

Each skill lives at `<plugin-name>/skills/<skill-name>/`. Copy that inner skill folder into wherever your runtime looks for skills, then point your assistant at it:

```bash
# from a clone of this repo
cp -R hebrew-md-to-docx/skills/hebrew-md-to-docx ~/.claude/skills/
```

Because a skill is just instructions plus Python, you can also hand the folder to any other assistant (ChatGPT, a local agent, etc.) and have it follow `SKILL.md` and run the script directly.

### claude.ai / Claude Desktop (Settings → Capabilities)

Every push that changes a skill publishes a `<skill-name>.skill` archive as a GitHub Release asset (see [`.github/workflows/package-skills.yml`](./.github/workflows/package-skills.yml)). Download the one you want from the [Releases](https://github.com/AssafTzurEl/assaf-agent-skills/releases) page and upload it under **Settings → Capabilities**.

### Claude Code plugin marketplace

Add the marketplace once, then install whichever skills you want:

```text
/plugin marketplace add AssafTzurEl/assaf-agent-skills
/plugin install hebrew-md-to-docx@assaf
/plugin install hebrew-md-to-pptx@assaf
/plugin install pptx-to-slide-md@assaf
```

`/plugin marketplace add` also accepts the full repository URL. Once installed, Claude activates each skill automatically when a task matches its description.

## Repository layout

```text
assaf-agent-skills/
├── .claude-plugin/
│   └── marketplace.json          # catalogs each skill as a plugin (marketplace name: assaf)
├── docs/
│   └── slide-markdown-format.md  # shared slide-authoring format spec
├── hebrew-md-to-docx/            # one plugin per skill
│   ├── .claude-plugin/plugin.json
│   └── skills/hebrew-md-to-docx/ # the actual skill (SKILL.md + scripts)
├── hebrew-md-to-pptx/
│   ├── .claude-plugin/plugin.json
│   └── skills/hebrew-md-to-pptx/
├── pptx-to-slide-md/
│   ├── .claude-plugin/plugin.json
│   └── skills/pptx-to-slide-md/
└── .github/workflows/package-skills.yml
```

The nested `skills/<skill-name>/` layout is what Claude Code requires to discover a skill inside a plugin; the inner folder is also the unit that gets zipped into a portable `.skill` and the unit you copy for manual install.

## License

[MIT](./LICENSE) © Assaf Tzur-El
