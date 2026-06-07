class ModelRiskMCPError(Exception):
    """Base class for all ModelRisk MCP server errors."""


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
