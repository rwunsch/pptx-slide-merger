"""
pptx-slide-merger: Copy slides between PowerPoint files preserving original formatting.

Pure-Python PPTX slide merger that copies slides with their original layouts,
masters, and themes intact. No .NET runtime required.
"""

from .merger import PptxMerger, list_slides, reorder_slides, move_slide
from .review import build_review_viewer, extract_slide_shapes, render_slide_pngs

__version__ = "0.1.0"
__all__ = [
    "PptxMerger", "list_slides", "reorder_slides", "move_slide",
    "build_review_viewer", "extract_slide_shapes", "render_slide_pngs",
]
