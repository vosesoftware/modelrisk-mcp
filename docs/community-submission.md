# Discovery: the MCP Registry

ModelRisk MCP is registered with the official [MCP Registry](https://registry.modelcontextprotocol.io/) under the server name `io.github.vosesoftware/modelrisk-mcp`. Listing there is the canonical way users and aggregator clients discover MCP servers.

## How registration works (already automated)

Publishing happens automatically when a version tag (`v*.*.*` or `v*.*.*-*`) is pushed. The pipeline in `.github/workflows/release.yml` runs four jobs in order:

1. `build-wheel` — wheel + sdist via `uv build`.
2. `build-windows-exe` — single-file PyInstaller exe; the obfuscation scan blocks the release on any plain-key hit.
3. `publish-pypi` — wheel + sdist to PyPI via trusted publishing (OIDC).
4. **`publish-mcp-registry`** — `mcp-publisher` authenticates via GitHub OIDC, the registry verifies ownership by fetching the new PyPI version's README and looking for `<!-- mcp-name: io.github.vosesoftware/modelrisk-mcp -->`, then registers the entry from `server.json`.

The verification gate (step 4) is why the marker in `README.md` and the `name` field in `server.json` must always match exactly.

## Ownership verification

The MCP Registry verifies that the GitHub identity publishing a server actually owns the underlying package. For PyPI packages, this is done by:

1. Fetching the package's current README from `pypi.org` (which is the `long_description` baked into the wheel + sdist).
2. Scanning that README for the literal string `mcp-name: <server-name>`.
3. Confirming `<server-name>` matches the `name` field in `server.json`.

Our marker lives in the README under the H1 title:

```markdown
# ModelRisk MCP

<!-- mcp-name: io.github.vosesoftware/modelrisk-mcp -->
```

It's a Markdown comment so users don't see it when reading the rendered page, but it's still present in the raw `long_description` PyPI serves to the registry's verifier.

## server.json

The registry entry's schema lives in `server.json` at repo root. Fields:

| Field | Meaning |
|---|---|
| `$schema` | The MCP Registry JSON Schema version |
| `name` | The canonical server identifier — `io.github.<github-org>/<repo>` for GitHub-OIDC-authed publishes |
| `title` | Human-readable display name |
| `description` | One-line summary (≤300 chars) shown in registry search |
| `version` | The server's version (must match the PyPI package version being verified) |
| `packages[]` | Where the registry tells clients to fetch the actual code from. For us: `registryType: pypi`, `identifier: modelrisk-mcp`, the current version, and `transport: stdio` |

When bumping versions, update `server.json`'s `version` and `packages[0].version` to match `pyproject.toml`. (A future improvement would derive these from the package version automatically during the release build — tracked in an issue.)

## Authentication

We use GitHub OIDC because we already use it for PyPI publishing — same trust model, no tokens to manage. The flow:

1. GitHub Actions runs as the `vosesoftware` org identity.
2. The job has `permissions: id-token: write`, which lets it mint a short-lived OIDC token.
3. `mcp-publisher login github-oidc` exchanges that token for a registry session.
4. Because our server name starts with `io.github.vosesoftware/`, the registry accepts the OIDC identity as the owner.

DNS-based authentication (where the server name would be `com.vosesoftware/modelrisk` instead) is a possible upgrade: cleaner branding, but requires adding a TXT record to `vosesoftware.com` DNS and switching the `mcp-publisher login` command. Not pursued yet — the GitHub-OIDC path works and rebranding the registry name would orphan installed configs.

## Manual publish (for emergencies)

If the automated pipeline can't run (e.g. registry is down during a CI window), publish from a maintainer's machine:

```powershell
# Install mcp-publisher (one-time)
$arch = if ([System.Runtime.InteropServices.RuntimeInformation]::ProcessArchitecture -eq "Arm64") { "arm64" } else { "amd64" }
Invoke-WebRequest "https://github.com/modelcontextprotocol/registry/releases/latest/download/mcp-publisher_windows_$arch.tar.gz" -OutFile mcp-publisher.tar.gz
tar xf mcp-publisher.tar.gz mcp-publisher.exe

# Authenticate (browser flow)
.\mcp-publisher.exe login github

# Publish (uses server.json in cwd)
.\mcp-publisher.exe publish
```

Versions must already be live on PyPI before publishing to the registry, or the verifier won't find the marker.

## Other places worth posting (separate channels)

Once the registry entry is live, parallel announcements still amplify reach:

- **Hacker News** — *Show HN: ModelRisk MCP — Open MCP server for Monte Carlo simulation in Excel*. Lead with the strategic narrative and a demo GIF. Weekday morning US Eastern is the best window.
- **r/excel + r/RiskManagement** — short post linking to the README and demo.
- **MCP community Discord / forums** — drop the link with one or two sentences.
- **LinkedIn** — Vose Software corporate + personal posts. The strategic narrative ("open vs proprietary connector") plays better on LinkedIn than the technical depth does.

Avoid being adversarial about closed alternatives. The README's comparison table does that work without naming names; piling on adds nothing and risks looking petty.
