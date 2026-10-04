from mcp.server.mcpserver.exceptions import ToolError


class ModelRiskMCPError(ToolError):
    """Base class for all ModelRisk MCP server errors.

    A `ToolError` because current mcp 2.x releases (2.1.1 and 2.3.0 checked)
    treat any other exception a tool raises as a crash: the client then sees
    only `Error executing tool <name>`, and the message written for the user
    is lost."""


class ExcelNotRunningError(ModelRiskMCPError):
    pass


class ModelRiskNotLoadedError(ModelRiskMCPError):
    pass


class ModelRiskNotFunctionalError(ModelRiskNotLoadedError):
    """The ModelRisk add-in is not live in the running Excel — Vose
    functions don't resolve (cells show #NAME?) and simulations can't
    run. Distinct from ModelRiskNotLoadedError's broader sense: this is
    specifically 'Excel is here but the add-in didn't load', which the
    bridge tries to auto-correct before raising this."""

    pass


class WorkbookNotFoundError(ModelRiskMCPError):
    pass


class CellReferenceError(ModelRiskMCPError):
    pass


class UnknownFunctionError(ModelRiskMCPError):
    pass


class ParameterMismatchError(ModelRiskMCPError):
    pass


class SimulationNotAvailableError(ModelRiskMCPError):
    """Raised when a simulation-control endpoint isn't exposed by the installed ModelRisk."""


class SimulationFailedError(ModelRiskMCPError):
    pass


class ConcurrentWriterError(ModelRiskMCPError):
    """Raised when another MCP server instance already holds the writer mutex."""


class CatalogueError(ModelRiskMCPError):
    pass


class ModelRiskComputationError(ModelRiskMCPError):
    """A ModelRisk worksheet computation returned an error or a
    non-numeric result — e.g. a distribution property evaluated to
    `#VALUE!`, or a fit returned 'parameter must be a valid Fit
    Object'. Carries the offending expression so the caller can see
    what ModelRisk rejected."""


class ReadOnlyModeError(ModelRiskMCPError):
    """The server is running in read-only mode; write/simulate/save
    operations are disabled. Launch without --read-only (or unset
    MODELRISK_MCP_READ_ONLY) to enable them."""
