"""
pptx-slide-merger: Copy slides between PowerPoint files preserving original formatting.

Pure-Python PPTX slide merger that copies slides with their original layouts,
masters, and themes intact. No .NET runtime required.
"""

from .merger import PptxMerger, list_slides

__version__ = "0.1.0"
__all__ = ["PptxMerger", "list_slides"]
