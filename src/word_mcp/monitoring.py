"""
Health monitoring for word-mcp server.

Provides HealthMonitor class that tracks Python process memory, system memory
percentage, COM pool utilization, and open document count. Exposes metrics
through MCP tool interface for production observability.

Key features:
- psutil-based resource tracking (process memory, system memory %)
- COM pool metrics integration (active instances, lifecycle counts)
- Status assessment (healthy/degraded/unhealthy) based on thresholds
- Alert generation for actionable issues
"""

import psutil
from typing import Dict, List

from .logging_config import get_logger
from .com_pool import com_pool
from .document_manager import document_manager

logger = get_logger(__name__)


class HealthMonitor:
    """
    Health monitor with configurable thresholds.

    Provides check_health() method that returns comprehensive server health
    including process memory, system memory, COM pool status, open document
    count, and actionable alerts.

    Status logic:
    - UNHEALTHY: System memory >80% OR active COM instances >= limit
    - DEGRADED: System memory >70% OR COM failures > 0
    - HEALTHY: Otherwise
    """

    def __init__(
        self,
        memory_threshold_percent: float = 80.0,
        com_instance_limit: int = 5
    ):
        """
        Initialize health monitor with thresholds.

        Args:
            memory_threshold_percent: System memory % threshold for unhealthy (default: 80%)
            com_instance_limit: Max active COM instances before unhealthy (default: 5)
        """
        self.memory_threshold_percent = memory_threshold_percent
        self.com_instance_limit = com_instance_limit

        logger.debug(
            "health_monitor_initialized",
            memory_threshold=memory_threshold_percent,
            com_limit=com_instance_limit
        )

    def check_health(self) -> Dict:
        """
        Check server health and return comprehensive metrics.

        Returns:
            Dictionary with keys:
            - status: "healthy" | "degraded" | "unhealthy"
            - process_memory_mb: Current process memory usage in MB
            - system_memory_percent: System-wide memory usage percentage
            - com_pool: Dict with active_instances, total_created, total_failed, pool_size
            - open_documents: Number of documents currently open in memory
            - alerts: List of actionable alert messages (empty if healthy)
        """
        # Gather metrics
        process = psutil.Process()
        process_memory_mb = process.memory_info().rss / (1024 * 1024)
        system_memory_percent = psutil.virtual_memory().percent

        com_metrics = com_pool.get_metrics()
        open_documents = len(document_manager.list_documents())

        # Determine status and alerts
        alerts: List[str] = []
        status = "healthy"

        # Unhealthy conditions
        if system_memory_percent > self.memory_threshold_percent:
            status = "unhealthy"
            alerts.append(
                f"System memory at {system_memory_percent:.1f}% "
                f"(threshold: {self.memory_threshold_percent:.1f}%)"
            )

        if com_metrics["active_count"] >= self.com_instance_limit:
            status = "unhealthy"
            alerts.append(
                f"Active COM instances at {com_metrics['active_count']} "
                f"(limit: {self.com_instance_limit})"
            )

        # Degraded conditions (if not already unhealthy)
        if status == "healthy":
            degraded_threshold = self.memory_threshold_percent - 10
            if system_memory_percent > degraded_threshold:
                status = "degraded"
                alerts.append(
                    f"System memory at {system_memory_percent:.1f}% "
                    f"(warning threshold: {degraded_threshold:.1f}%)"
                )

            if com_metrics["total_failed"] > 0:
                status = "degraded"
                alerts.append(
                    f"COM instance failures detected: {com_metrics['total_failed']} "
                    f"(check logs for details)"
                )

        # Log warnings for non-healthy states
        if status in ("degraded", "unhealthy"):
            logger.warning(
                "health_check_warning",
                status=status,
                process_memory_mb=process_memory_mb,
                system_memory_percent=system_memory_percent,
                com_active=com_metrics["active_count"],
                com_failed=com_metrics["total_failed"],
                alerts=alerts
            )

        return {
            "status": status,
            "process_memory_mb": process_memory_mb,
            "system_memory_percent": system_memory_percent,
            "com_pool": {
                "active_instances": com_metrics["active_count"],
                "total_created": com_metrics["total_created"],
                "total_failed": com_metrics["total_failed"],
                "pool_size": com_metrics["pool_size"],
            },
            "open_documents": open_documents,
            "alerts": alerts,
        }


# Module-level singleton
health_monitor = HealthMonitor()
