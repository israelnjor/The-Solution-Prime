"""Microsoft Word automation (Windows COM) for baselines: page counts, PDF renders, re-saves.

Word is the compatibility target, so its rendering of a file is the reference
for "what the user will see" when they open our output in Word.
"""
import os

import pythoncom
import win32com.client

WD_STATISTIC_PAGES = 2
WD_EXPORT_FORMAT_PDF = 17
WD_FORMAT_DOCUMENT_DEFAULT = 16
WD_DO_NOT_SAVE_CHANGES = 0


class Word:
    def __enter__(self):
        pythoncom.CoInitialize()
        self.app = win32com.client.DispatchEx("Word.Application")
        self.app.Visible = False
        self.app.DisplayAlerts = 0
        return self

    def __exit__(self, *exc):
        try:
            self.app.Quit(WD_DO_NOT_SAVE_CHANGES)
        except Exception:
            # Word shares one hidden instance between sessions: another session may have closed it already.
            pass
        finally:
            pythoncom.CoUninitialize()

    def _open(self, path, read_only=True):
        return self.app.Documents.Open(FileName=os.path.abspath(path), ConfirmConversions=False,
                                       ReadOnly=read_only, AddToRecentFiles=False, Visible=False)

    def inspect(self, path, pdf_path=None):
        """Return {'pages': n}; optionally export a PDF render alongside."""
        doc = self._open(path)
        try:
            doc.Repaginate()
            pages = doc.ComputeStatistics(WD_STATISTIC_PAGES)
            if pdf_path:
                doc.ExportAsFixedFormat(OutputFileName=os.path.abspath(pdf_path), ExportFormat=WD_EXPORT_FORMAT_PDF)
            return {"pages": pages}
        finally:
            doc.Close(WD_DO_NOT_SAVE_CHANGES)

    def resave(self, src, dst):
        doc = self._open(src)
        try:
            doc.SaveAs2(FileName=os.path.abspath(dst), FileFormat=WD_FORMAT_DOCUMENT_DEFAULT)
        finally:
            doc.Close(WD_DO_NOT_SAVE_CHANGES)
