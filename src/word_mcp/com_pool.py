"""
COM process pool for word-mcp.

This module provides a semaphore-limited COM instance pool to prevent resource
exhaustion under concurrent load. Instead of creating unlimited Word COM instances
(one per operation), the pool limits concurrent instances to a configurable maximum
(default: 3).

Key design:
- threading.Semaphore (NOT asyncio) for concurrency limiting since MCP tools are sync
- Context manager interface compatible with existing WordApplication pattern
- Lifecycle tracking for observability (active instances, total created, failures)
- Emergency cleanup for graceful shutdown
"""

import threading
import gc
from contextlib import contextmanager
from typing import Dict
import win32com.client

from .logging_config import get_logger

logger = get_logger(__name__)


class COMPool:
    """
    COM instance pool with semaphore-based concurrency limiting.

    Limits the number of concurrent Word COM instances to prevent resource
    exhaustion. Provides a context manager interface compatible with the
    existing WordApplication pattern.

    Usage:
        with com_pool.get_word_app() as word:
            doc = word.Documents.Open(abs_path)
            doc.TrackRevisions = True
            doc.Save()
            doc.Close()
    """

    def __init__(self, pool_size: int = 3):
        """
        Initialize COM pool with concurrency limit.

        Args:
            pool_size: Maximum number of concurrent Word instances (default: 3)
        """
        self.pool_size = pool_size
        self._semaphore = threading.Semaphore(pool_size)
        self._active_instances = []
        self._lock = threading.Lock()

        # Metrics
        self.total_created = 0
        self.total_failed = 0

        logger.debug("com_pool_initialized", pool_size=pool_size)

    @contextmanager
    def get_word_app(self, visible: bool = False):
        """
        Get a Word COM application instance with semaphore limiting.

        Returns a context manager that acquires a semaphore slot, creates a
        Word COM instance, and ensures cleanup on exit. Interface is compatible
        with existing WordApplication context manager.

        Args:
            visible: Whether Word window should be visible (default: False)

        Yields:
            Word.Application COM object

        Example:
            with com_pool.get_word_app() as word:
                doc = word.Documents.Open("C:/test.docx")
                doc.Save()
                doc.Close()
        """
        # Acquire semaphore (blocks if pool_size instances already running)
        self._semaphore.acquire()

        app = None
        try:
            # Create Word COM instance (DispatchEx for isolated instance)
            app = win32com.client.DispatchEx("Word.Application")
            app.Visible = visible
            app.DisplayAlerts = 0  # wdAlertsNone - prevent automation hangs

            # Track active instance
            with self._lock:
                self._active_instances.append(app)
                self.total_created += 1
                active_count = len(self._active_instances)

            logger.debug(
                "com_instance_created",
                active_count=active_count,
                total_created=self.total_created,
                pool_size=self.pool_size
            )

            yield app

        except Exception as e:
            with self._lock:
                self.total_failed += 1
            logger.error(
                "com_instance_creation_failed",
                error=str(e),
                error_type=type(e).__name__,
                total_failed=self.total_failed
            )
            raise

        finally:
            # Cleanup: close documents, quit Word, release COM references
            if app:
                try:
                    # Close any open documents without saving
                    while app.Documents.Count > 0:
                        app.Documents(1).Close(SaveChanges=0)  # wdDoNotSaveChanges
                except Exception as e:
                    logger.warning("com_document_cleanup_failed", error=str(e))

                try:
                    app.Quit()
                except Exception as e:
                    logger.warning("com_quit_failed", error=str(e))

                finally:
                    # Remove from active instances
                    with self._lock:
                        if app in self._active_instances:
                            self._active_instances.remove(app)
                        active_count = len(self._active_instances)

                    # Force COM reference cleanup
                    del app
                    gc.collect()

                    logger.debug(
                        "com_instance_cleaned_up",
                        active_count=active_count,
                        pool_size=self.pool_size
                    )

            # Release semaphore slot
            self._semaphore.release()

    def close_all(self):
        """
        Emergency cleanup: close all active COM instances.

        Called during server shutdown to ensure no zombie WINWORD.EXE processes
        remain. Iterates through all tracked instances and attempts cleanup.
        """
        with self._lock:
            instances_to_close = list(self._active_instances)
            count = len(instances_to_close)

        if count == 0:
            logger.info("com_pool_shutdown_no_active_instances")
            return

        logger.info("com_pool_shutdown_closing_instances", count=count)

        for app in instances_to_close:
            try:
                # Close documents without saving
                while app.Documents.Count > 0:
                    app.Documents(1).Close(SaveChanges=0)
            except Exception as e:
                logger.warning("com_shutdown_document_cleanup_failed", error=str(e))

            try:
                app.Quit()
            except Exception as e:
                logger.warning("com_shutdown_quit_failed", error=str(e))

            finally:
                try:
                    del app
                except Exception:
                    pass

        # Clear tracking
        with self._lock:
            self._active_instances.clear()

        gc.collect()
        logger.info("com_pool_shutdown_complete", instances_closed=count)

    def get_metrics(self) -> Dict:
        """
        Get pool metrics for observability.

        Returns:
            Dictionary with current pool state:
            - active_count: Number of currently active instances
            - total_created: Total instances created since server start
            - total_failed: Total failed instance creation attempts
            - pool_size: Maximum concurrent instances allowed
            - available_slots: Number of available semaphore slots
        """
        with self._lock:
            active_count = len(self._active_instances)

        # Calculate available slots (pool_size - active_count)
        # Note: This is approximate since we can't query Semaphore._value directly
        available_slots = self.pool_size - active_count

        return {
            "active_count": active_count,
            "total_created": self.total_created,
            "total_failed": self.total_failed,
            "pool_size": self.pool_size,
            "available_slots": available_slots
        }


# Module-level singleton
com_pool = COMPool()
