class ModelRiskMCPError(Exception):
    """Base class for all ModelRisk MCP server errors."""


class ExcelNotRunningError(ModelRiskMCPError):
    pass


class ModelRiskNotLoadedError(ModelRiskMCPError):
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
