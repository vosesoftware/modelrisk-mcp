# Working preferences for Claude on this repo

Read this before doing work on `modelrisk-mcp`.

## Git push policy

**Always do `git push` yourself.** Don't ask the user to switch to PowerShell and run it. The credential dialog works reliably from the bash/PowerShell shell Claude uses, provided the repo-local credential helper is set to `manager`:

```bash
git config --local credential.https://github.com.helper "manager"
```

That's already configured for `modelrisk-mcp` and shouldn't need to be re-set. If a push hangs, it's almost always Git Credential Manager waiting on the OS chooser — wait it out or check the user's screen for the dialog. **Don't fall back to telling the user to push manually.**

The standard sequence after a tagged commit:

```
git push origin main
git push origin <new-tag>
gh run watch <run-id> --repo vosesoftware/modelrisk-mcp --exit-status
```

## Release pipeline (automated)

Pushing a tag `v*.*.*` or `v*.*.*-*` triggers `.github/workflows/release.yml`. Five jobs run, all green-or-fail:

1. `build-wheel` — wheel + sdist via `uv build`
2. `build-windows-exe` — PyInstaller exe + obfuscation scan (release-blocker)
3. `publish-pypi` — wheel + sdist via OIDC trusted publishing
4. `publish-mcp-registry` — `mcp-publisher` via GitHub OIDC against `registry.modelcontextprotocol.io`
5. `github-release` — assets attached to a GitHub Release page

After a tag pushes, watch the run with `gh run watch <id> --exit-status` and confirm all five jobs green. Then confirm the version appears on PyPI (`curl pypi.org/pypi/modelrisk-mcp/json | jq .info.version`) and on the MCP Registry.

## Quality gate before every commit

Always run all three before committing. No exceptions:

```
.venv/Scripts/python.exe -m pytest tests/unit -q
.venv/Scripts/python.exe -m ruff check src tests scripts
.venv/Scripts/python.exe -m mypy src
```

`mypy` runs on `src` only — running it on `tests/` surfaces a lot of test-time-only `type: ignore` lint noise that's not actionable.

## Version bumping

Bump in five places per release. They must all match or the registry verifier rejects:

- `src/modelrisk_mcp/__init__.py::__version__`
- `pyproject.toml::version`
- `server.json::version` AND `packages[0].version`
- `tests/unit/test_server_boot.py::test_version_is_set` assertion
- `uv.lock` (regenerates automatically on `uv sync --extra dev`)

The CHANGELOG entry under the new version header gets the user-facing summary.

## File-virtualisation gotcha (Microsoft Store Python + Claude Desktop)

The user has Microsoft Store Python at `C:\Users\timou\AppData\Local\Python\pythoncore-3.14-64\python.exe`. The `python` command alias is intercepted by Windows App Execution Aliases and goes to the Store stub instead of the real install. **Always use `py -m modelrisk_mcp ...` or the full path, never bare `python -m ...`.**

The user also has two Claude Desktop installs (regular + MS Store packaged). Both share `%APPDATA%\Claude\` via Windows file virtualisation, so a single config write reaches both.

## Stale-bug-report pattern

When the user (or Claude in a fresh session) reports a bug that references tools we deleted (`use_vba_helper_for_simulation`, `ensure_modelrisk_active` and the bitness-mismatch hypothesis), that's training-data drift — those tools were removed in v0.3.0-alpha.1 when we pivoted from ATL COM dispatch to the MRService.dll + XLL command surface. The MCP server's actual `tools/list` returns the current 38 tools; nothing named that way is registered.

## Architectural anchors

- Simulation kickoff: `Application.Run("VoseStartSimulCustom12", options_array)` then `Application.Run("VoseGetDataSZ12", session_name, save_path)`. The session-name format is `h<hwndExcel>_SaveResultsToFile_<book_name>`.
- Results read: ctypes against `MRService.dll`. Bundled activation key in `bridge/_keymat.py`; rotation script in `scripts/encode_activation_key.py`.
- Excel must be launched **interactively** (Start menu / taskbar) before the MCP server tries `run_simulation` — when launched programmatically, the XLL skips `xlAutoOpen` and the simulation commands aren't in Excel's `Application.Run` table.
- Real Excel returns `Range.formula` as **tuples**, not lists. `_as_2d` must accept both at every nesting level — a regression here silently collapses list-scans into single records (the alpha.11 fix). Don't regress this.
