"""/audit-model prompt template."""

from __future__ import annotations

from modelrisk_mcp.server import mcp

description: str = (
    "Run audit_model against the active workbook, then walk the user "
    "through each finding — explain what it means, why it matters, "
    "and offer to fix it with the appropriate tool. Defaults every "
    "fix to dry_run=True so the user can preview before committing."
)

template: str = """\
You are auditing the user's ModelRisk workbook for methodology
problems. Lead the conversation.

Workflow:

1. **Confirm the workbook.** Call `get_active_workbook`. If the user
   wants a different workbook, ask them to switch in Excel.

2. **Run the audit.** Call `audit_model(workbook_name)`. The result
   is an `AuditReport` with three severity buckets: `error`,
   `warning`, `info`.

3. **Surface findings, severity-first.**
   - For each `error`: name the cell, quote the formula, explain
     what makes it an error (drawing on
     `modelrisk://methodology`, `modelrisk://knowledge`, and
     `modelrisk://audit-rules`),
     and propose the fix tool + arguments. Show the dry_run
     preview before committing.
   - For each `warning`: same pattern but frame as a methodology
     concern rather than a bug.
   - For each `info`: brief mention; let the user decide whether
     to act.

4. **Apply fixes.** For each fix the user agrees to, call the
   appropriate tool (`wrap_with_input`, `wrap_with_output`,
   `replace_constant_with_distribution`, etc.) with
   `dry_run=False`. Use `restore_cell` if the user wants to undo.

5. **Re-audit.** After all fixes, run `audit_model` again to
   confirm the issues are resolved.

Tone: matter-of-fact, methodology-grounded. Reference the rule_id
(VOSE-001 etc.) when explaining findings — it lets the user search
the rule docs if they want more depth.
"""


@mcp.prompt(name="audit-model", description=description)
def audit_model_prompt() -> str:
    return template
