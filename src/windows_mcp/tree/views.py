"""Data models for the accessibility tree state.

After the OSWorld compatibility rewrite, TreeState holds:
- xml_tree: raw XML accessibility tree string
- tsv_tree: filtered + linearized TSV (what the model sees)
- filtered_nodes: lxml Element nodes that passed the filter (for SoM drawing)
- marks: bounding boxes [x, y, w, h] for each filtered node (for SoM drawing)

BoundingBox and Center are kept for desktop/service.py compatibility.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Optional


@dataclass
class TreeState:
    """Accessibility tree state, OSWorld-compatible."""

    # ── Primary data (OSWorld-compatible) ─────────────────────
    xml_tree: str = ""  # Raw XML a11y tree
    tsv_tree: str = ""  # Filtered + linearized TSV (model observation)
    filtered_nodes: list = field(default_factory=list)  # lxml Elements that passed filter
    marks: list = field(default_factory=list)  # [[x, y, w, h], ...] for SoM

    # ── DOM scraping compat (for Scrape tool use_dom mode) ────
    dom_node: Optional[Any] = None
    dom_informative_nodes: list = field(default_factory=list)

    # ── Backward-compat methods (used by fastapi_server.py + __main__.py) ──

    def interactive_elements_to_string(self) -> str:
        """Returns the TSV tree (replaces old pipe-delimited interactive elements)."""
        return self.tsv_tree or "No interactive elements"

    def scrollable_elements_to_string(self) -> str:
        """Scrollable elements are now included in the TSV tree."""
        return ""

    @property
    def interactive_nodes(self) -> list:
        """Backward compat for SoM — returns empty list.
        Use filtered_nodes + marks for OSWorld-style SoM instead."""
        return []


@dataclass
class BoundingBox:
    left: int
    top: int
    right: int
    bottom: int
    width: int
    height: int

    @classmethod
    def from_bounding_rectangle(cls, bounding_rectangle: Any) -> "BoundingBox":
        return cls(
            left=bounding_rectangle.left,
            top=bounding_rectangle.top,
            right=bounding_rectangle.right,
            bottom=bounding_rectangle.bottom,
            width=bounding_rectangle.width(),
            height=bounding_rectangle.height(),
        )

    def get_center(self) -> "Center":
        return Center(x=self.left + self.width // 2, y=self.top + self.height // 2)

    def xywh_to_string(self):
        return f"({self.left},{self.top},{self.width},{self.height})"

    def xyxy_to_string(self):
        x1, y1, x2, y2 = self.convert_xywh_to_xyxy()
        return f"({x1},{y1},{x2},{y2})"

    def convert_xywh_to_xyxy(self) -> tuple[int, int, int, int]:
        x1, y1 = self.left, self.top
        x2, y2 = self.left + self.width, self.top + self.height
        return x1, y1, x2, y2


@dataclass
class Center:
    x: int
    y: int

    def to_string(self) -> str:
        return f"({self.x},{self.y})"


# ── Legacy types (kept for desktop/service.py import compatibility) ──


@dataclass
class TreeElementNode:
    """Legacy node type. Kept for backward compat with desktop/service.py imports."""

    bounding_box: BoundingBox = field(default_factory=lambda: BoundingBox(0, 0, 0, 0, 0, 0))
    center: Center = field(default_factory=lambda: Center(0, 0))
    name: str = ""
    control_type: str = ""
    window_name: str = ""
    value: str = ""
    shortcut: str = ""
    xpath: str = ""
    is_focused: bool = False


@dataclass
class ScrollElementNode:
    """Legacy node type. Kept for backward compat."""

    name: str = ""
    control_type: str = ""
    xpath: str = ""
    window_name: str = ""
    bounding_box: BoundingBox = field(default_factory=lambda: BoundingBox(0, 0, 0, 0, 0, 0))
    center: Center = field(default_factory=lambda: Center(0, 0))
    horizontal_scrollable: bool = False
    horizontal_scroll_percent: float = 0
    vertical_scrollable: bool = False
    vertical_scroll_percent: float = 0
    is_focused: bool = False


@dataclass
class TextElementNode:
    """Legacy node type. Kept for scrape tool compat."""

    text: str = ""


ElementNode = TreeElementNode | ScrollElementNode | TextElementNode
