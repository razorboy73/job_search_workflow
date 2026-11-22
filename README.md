# WAT: Workflows, Agents, Tools

A starter template for agentic workflows with Claude Code. The design splits the work in two: the AI agent reasons and coordinates, and deterministic Python scripts carry out each step.

| Layer | Where it lives | What it does |
|---|---|---|
| **Workflows** | `workflows/*.md` | Plain-language instructions: the goal, the inputs, which tools to run and what to do when something goes wrong |
| **Agent** | Claude Code, guided by `CLAUDE.md` | Reads the right workflow, runs the tools in order, recovers from errors and asks questions when it needs to |
| **Tools** | `tools/*.py` | Small, single-purpose scripts that make the API calls, process data and handle files |

Why split it this way: if each step an AI does by hand is 90% reliable, a five-step job succeeds only about 59% of the time. Handing each step to a tested script keeps the end-to-end result reliable.

## Start a new project from this template

Click **Use this template** on GitHub, or run:

```
gh repo create <new-name> --template razorboy73/WAT --private --clone
cd <new-name>
```

Then set up the environment (Windows):

```
<full-path-to-python.exe> -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
copy .env.example .env
```

On Windows, typing `python` may open the Microsoft Store instead of running Python. If it does, give the full path to a real Python install, such as your Anaconda `python.exe`.

## Project layout

```
CLAUDE.md          # How the agent should operate in this project
workflows/         # Markdown instructions, one per task
tools/             # Python scripts; start new ones from _template.py
.tmp/              # Throwaway intermediate files (not committed)
.env               # API keys and settings (not committed)
.env.example       # The keys a project needs, with no real values
requirements.txt   # Python packages
```

The `.gitignore` keeps `.env`, `credentials.json` and `token.json` (Google sign-in files), `.tmp/` and `.venv/` out of git.

Finished work goes to cloud services such as Google Sheets or Slides. Local files are only for processing along the way.

## Adding a tool

1. Copy `tools/_template.py` to `tools/<name>.py`.
2. Take inputs as command-line arguments, read keys from `.env`, and write any scratch files to `.tmp/`.
3. Print a short result and exit with a non-zero code on failure, so the agent can tell something went wrong.
4. Add any new packages to `requirements.txt` and any new keys to `.env.example`.

Run a tool with:

```
.venv\Scripts\python.exe tools\<name>.py --help
```

## Adding a workflow

Write a Markdown file in `workflows/` that covers the goal, the inputs, which tools to run and in what order, what it should produce, and the known edge cases. Update the workflow whenever you learn something new, such as a rate limit or a better method, so later runs benefit.

## Using it with Claude Code

Open the project in Claude Code and describe what you want done. Claude reads `CLAUDE.md`, finds the matching workflow and runs the tools it names. When a tool fails, Claude fixes the tool, tests it again and records what it learned in the workflow. It checks with you before re-running anything that costs money or credits.
