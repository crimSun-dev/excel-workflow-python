"""Multi-workflow processing package.

Each report workflow (Akumulasi, Rincian Vol TF, Rincian Portal BG) is a
`WorkflowStrategy` subclass sharing a common ingest -> optional enrich ->
aggregate -> export skeleton. Concrete strategies and the dispatch registry
live in submodules; this package root is intentionally kept import-light so
`schemas` can reference `workflows.base.WorkflowId` without a circular import.

Import specifics from the submodules directly, e.g.::

    from src.workflows.base import WorkflowId
    from src.workflows.registry import get_strategy, WORKFLOW_REGISTRY
"""
