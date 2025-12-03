"""Customer CLI toolkit.

Exposes the high-level ``run_customer_sync`` API for programmatic use.
"""

from .runner import run_customer_sync  # Public API for synchronisation

__all__ = ["run_customer_sync"]  # Re-exported symbol