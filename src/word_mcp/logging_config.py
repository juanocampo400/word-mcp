"""Structured logging configuration for word-mcp.

Provides JSON-formatted structured logging using structlog for production observability.
All logs are output to stderr with ISO timestamps, log levels, and contextual information.
"""

import logging
import sys
import structlog


def configure_logging() -> None:
    """Configure structlog with JSON rendering for production use.

    Sets up structlog processors for:
    - ISO timestamp formatting
    - Log level addition
    - JSON rendering

    Also configures stdlib logging to route through structlog.
    """
    structlog.configure(
        processors=[
            structlog.stdlib.filter_by_level,
            structlog.stdlib.add_logger_name,
            structlog.stdlib.add_log_level,
            structlog.processors.TimeStamper(fmt="iso"),
            structlog.processors.StackInfoRenderer(),
            structlog.processors.format_exc_info,
            structlog.processors.JSONRenderer()
        ],
        wrapper_class=structlog.stdlib.BoundLogger,
        context_class=dict,
        logger_factory=structlog.stdlib.LoggerFactory(),
        cache_logger_on_first_use=True,
    )

    # Configure stdlib logging to also use structlog
    logging.basicConfig(
        format="%(message)s",
        stream=sys.stderr,
        level=logging.INFO,
    )


def get_logger(name: str) -> structlog.stdlib.BoundLogger:
    """Get a logger instance bound with the module name.

    Args:
        name: Module name (typically __name__)

    Returns:
        A structlog BoundLogger instance with module context
    """
    return structlog.get_logger(name)


# Configure on module import
configure_logging()
