"""
COM automation infrastructure for word-mcp.

This module provides the WordApplication context manager, which bridges between
python-docx (Phase 1) and win32com (Phase 2+). All COM operations MUST use this
context manager to prevent zombie WINWORD.EXE processes.

Key design:
- DispatchEx (NOT Dispatch) creates isolated Word instances that don't conflict
  with user's open Word documents
- DisplayAlerts=0 prevents automation hangs on dialog popups
- Explicit cleanup (Quit + del + gc.collect) prevents zombie processes
- Context manager pattern ensures cleanup happens even on errors
"""

import win32com.client
import gc


class WordApplication:
    """Context manager for Word.Application COM lifecycle.

    Uses DispatchEx (NOT Dispatch) for isolated instances that don't
    conflict with user's open Word documents.

    Usage:
        with WordApplication() as word:
            doc = word.Documents.Open(abs_path)
            doc.TrackRevisions = True
            doc.Save()
            doc.Close()
    """
    def __init__(self, visible=False):
        self.visible = visible
        self.app = None

    def __enter__(self):
        self.app = win32com.client.DispatchEx("Word.Application")
        self.app.Visible = self.visible
        self.app.DisplayAlerts = 0  # wdAlertsNone - prevent automation hangs
        return self.app

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.app:
            try:
                # Close any open documents without saving (caller handles save)
                while self.app.Documents.Count > 0:
                    self.app.Documents(1).Close(SaveChanges=0)  # wdDoNotSaveChanges
            except Exception:
                pass
            try:
                self.app.Quit()
            except Exception:
                pass
            finally:
                del self.app
                self.app = None
                gc.collect()  # Force COM reference cleanup
        return False  # Propagate exceptions
