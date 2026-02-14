"""
Monitoring MCP tool for word-mcp server.

Provides get_server_health() function that returns formatted health report
for consumption by MCP clients (Claude, etc.).
"""


def get_server_health() -> str:
    """
    Get server health metrics and status.

    Returns formatted health report showing:
    - Status (HEALTHY/DEGRADED/UNHEALTHY)
    - Process memory usage
    - System memory percentage
    - Open document count
    - COM pool metrics (active, created, failed, pool size)
    - Active alerts (if any)

    Returns:
        Formatted multi-line string with health metrics
    """
    from ..monitoring import health_monitor

    metrics = health_monitor.check_health()

    # Format as readable string
    lines = [
        f"Server Health: {metrics['status'].upper()}",
        "",
        f"Process Memory: {metrics['process_memory_mb']:.1f} MB",
        f"System Memory: {metrics['system_memory_percent']:.1f}%",
        f"Open Documents: {metrics['open_documents']}",
        "",
        "COM Pool:",
        f"  Active instances: {metrics['com_pool']['active_instances']}",
        f"  Total created: {metrics['com_pool']['total_created']}",
        f"  Total failed: {metrics['com_pool']['total_failed']}",
        f"  Pool size limit: {metrics['com_pool']['pool_size']}",
    ]

    if metrics['alerts']:
        lines.append("")
        lines.append("Alerts:")
        for alert in metrics['alerts']:
            lines.append(f"  - {alert}")

    return "\n".join(lines)
