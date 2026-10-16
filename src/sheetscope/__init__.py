"""Read Excel workbooks the way an analyst does: structure first, then search, then exact cells with citations."""

from .book import Book, SheetError, load

__version__ = "0.2.0"
__all__ = ["Book", "SheetError", "load", "__version__"]
