"""Accessibility tree service using pywinauto, compatible with OSWorld.

Generates XML accessibility tree → filters nodes → linearizes to TSV.
Matches OSWorld's observation format exactly (xlang-ai/OSWorld commit afa0876).

Architecture:
  1. pywinauto.Desktop(backend="uia") enumerates top-level windows
  2. _create_pywinauto_node() recursively builds XML (port of OSWorld)
  3. _judge_node() filters relevant nodes (port of OSWorld heuristic_retrieve)
  4. _linearize_tree() produces TSV output (port of OSWorld agent.py)
"""

from __future__ import annotations

import concurrent.futures
import logging
import weakref
from time import time
from typing import TYPE_CHECKING, Any, Dict, List, Optional, Set, Tuple

import lxml.etree
from lxml.etree import _Element

from windows_mcp.tree.views import TreeState

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

if TYPE_CHECKING:
    from windows_mcp.desktop.service import Desktop

# ── OSWorld namespace map (Windows) ──────────────────────────────────────────
# Exact copy from OSWorld desktop_env/server/main.py
NS_MAP = {
    "st": "https://accessibility.windows.example.org/ns/state",
    "attr": "https://accessibility.windows.example.org/ns/attributes",
    "cp": "https://accessibility.windows.example.org/ns/component",
    "doc": "https://accessibility.windows.example.org/ns/document",
    "docattr": "https://accessibility.windows.example.org/ns/document/attributes",
    "txt": "https://accessibility.windows.example.org/ns/text",
    "val": "https://accessibility.windows.example.org/ns/value",
    "act": "https://accessibility.windows.example.org/ns/action",
    "class": "https://accessibility.windows.example.org/ns/class",
}

MAX_DEPTH = 50
MAX_WIDTH = 1024

# ── Node filter tags (from OSWorld heuristic_retrieve.py) ────────────────────
# These control which nodes survive the filter and appear in the TSV output.

FILTER_TAG_SUFFIXES = {
    "item",
    "button",
    "heading",
    "label",
    "scrollbar",
    "searchbox",
    "textbox",
    "link",
    "tabelement",
    "textfield",
    "textarea",
    "menu",
}

FILTER_TAG_EXACT = {
    "alert",
    "canvas",
    "check-box",
    "combo-box",
    "entry",
    "icon",
    "image",
    "paragraph",
    "scroll-bar",
    "section",
    "slider",
    "static",
    "table-cell",
    "terminal",
    "text",
    "netuiribbontab",
    "start",
    "trayclockwclass",
    "traydummysearchcontrol",
    "uiimage",
    "uiproperty",
    "uiribboncommandbar",
}


# ── XML tree builder (port of OSWorld _create_pywinauto_node) ────────────────


def _create_pywinauto_node(
    node, nodes: Optional[Set] = None, depth: int = 0, flag: Optional[str] = None
) -> Optional[_Element]:
    """Build an lxml Element from a pywinauto wrapper node.

    Port of OSWorld desktop_env/server/main.py L580-L745.
    """
    if nodes is None:
        nodes = set()
    if id(node) in nodes:
        return None
    nodes.add(id(node))

    attribute_dict: Dict[str, str] = {"name": node.element_info.name or ""}

    # Get base properties with OSWorld's fallback mechanism
    base_properties: Dict[str, Any] = {}
    try:
        base_properties.update(node.get_properties())
    except Exception:
        try:
            import pywinauto.base_wrapper

            _element_class = node.__class__

            class TempElement(node.__class__):
                writable_props = pywinauto.base_wrapper.BaseWrapper.writable_props

            node.__class__ = TempElement
            properties = node.get_properties()
            node.__class__ = _element_class
            base_properties.update(properties)
        except Exception:
            pass

    # States (19 attributes — matching OSWorld exactly)
    for attr_name, attr_func in [
        ("enabled", lambda: node.is_enabled()),
        ("visible", lambda: node.is_visible()),
        ("minimized", lambda: node.is_minimized()),
        ("maximized", lambda: node.is_maximized()),
        ("normal", lambda: node.is_normal()),
        ("unicode", lambda: node.is_unicode()),
        ("collapsed", lambda: node.is_collapsed()),
        ("checkable", lambda: node.is_checkable()),
        ("checked", lambda: node.is_checked()),
        ("focused", lambda: node.is_focused()),
        ("keyboard_focused", lambda: node.is_keyboard_focused()),
        ("selected", lambda: node.is_selected()),
        ("selection_required", lambda: node.is_selection_required()),
        ("pressable", lambda: node.is_pressable()),
        ("pressed", lambda: node.is_pressed()),
        ("expanded", lambda: node.is_expanded()),
        ("editable", lambda: node.is_editable()),
        ("has_keyboard_focus", lambda: node.has_keyboard_focus()),
        ("is_keyboard_focusable", lambda: node.is_keyboard_focusable()),
    ]:
        try:
            attribute_dict[f"{{{NS_MAP['st']}}}{attr_name}"] = str(attr_func()).lower()
        except Exception:
            pass

    # Component (screen coordinates + size)
    try:
        rectangle = node.rectangle()
        attribute_dict[f"{{{NS_MAP['cp']}}}screencoord"] = (
            f"({rectangle.left:d}, {rectangle.top:d})"
        )
        attribute_dict[f"{{{NS_MAP['cp']}}}size"] = (
            f"({rectangle.width():d}, {rectangle.height():d})"
        )
    except Exception as e:
        logger.debug(f"Error accessing rectangle: {e}")

    # Text
    text: str = node.window_text() or ""
    if text == attribute_dict.get("name", ""):
        text = ""

    # Selection
    if hasattr(node, "select"):
        attribute_dict["selection"] = "true"

    # Value
    for attr_name, attr_funcs in [
        ("step", [lambda: node.get_step()]),
        ("value", [lambda: node.value(), lambda: node.get_value(), lambda: node.get_position()]),
        ("min", [lambda: node.min_value(), lambda: node.get_range_min()]),
        ("max", [lambda: node.max_value(), lambda: node.get_range_max()]),
    ]:
        for attr_func in attr_funcs:
            try:
                attribute_dict[f"{{{NS_MAP['val']}}}{attr_name}"] = str(attr_func())
                break
            except Exception:
                pass

    # Class info
    attribute_dict[f"{{{NS_MAP['class']}}}class"] = str(type(node))
    for attr_name in ["class_name", "friendly_class_name"]:
        try:
            val = base_properties.get(attr_name)
            if val is not None:
                attribute_dict[f"{{{NS_MAP['class']}}}{attr_name}"] = str(val).lower()
        except Exception:
            pass

    # Tag name from class_name (OSWorld logic), with fallback to
    # friendly_class_name for elements that have an empty class_name
    # (e.g. Desktop icon ListItems inside Progman/SysListView32).
    node_role_name: str = node.class_name().lower().replace(" ", "-")
    if not node_role_name.strip():
        try:
            node_role_name = node.friendly_class_name().lower().replace(" ", "-")
        except Exception:
            pass
    node_role_name = "".join(
        ch if ch.isidentifier() or ch in {"-"} or ch.isalnum() else "-" for ch in node_role_name
    )
    if not node_role_name.strip():
        node_role_name = "unknown"
    if not node_role_name[0].isalpha():
        node_role_name = "tag" + node_role_name

    # Create XML element
    xml_node = lxml.etree.Element(
        node_role_name,
        attrib=attribute_dict,
        nsmap=NS_MAP,
    )

    if text and text != attribute_dict.get("name", ""):
        xml_node.text = text

    if depth >= MAX_DEPTH:
        logger.warning("Max depth reached")
        return xml_node

    # Recurse into children (parallel, like OSWorld)
    children = node.children()
    if children:
        with concurrent.futures.ThreadPoolExecutor() as executor:
            futures = [
                executor.submit(_create_pywinauto_node, ch, nodes, depth + 1, flag)
                for ch in children[:MAX_WIDTH]
            ]
            try:
                for future in concurrent.futures.as_completed(futures):
                    try:
                        child_xml = future.result()
                        if child_xml is not None:
                            xml_node.append(child_xml)
                    except Exception as e:
                        logger.debug(f"Error processing child node: {e}")
            except Exception as e:
                logger.error(f"Exception in child processing: {e}")

    return xml_node


# ── Node filter (port of OSWorld heuristic_retrieve.py judge_node) ───────────


def _judge_node(node: _Element) -> bool:
    """Determine if an XML node should be included in the TSV output.

    Port of OSWorld mm_agents/accessibility_tree_wrap/heuristic_retrieve.py L38-L93.
    """
    tag = node.tag

    # Tag check
    keeps = (
        tag.startswith("document")
        or any(tag.endswith(suffix) for suffix in FILTER_TAG_SUFFIXES)
        or tag in FILTER_TAG_EXACT
    )
    if not keeps:
        return False

    st_ns = NS_MAP["st"]
    cp_ns = NS_MAP["cp"]

    # Visibility check (Windows: only visible required, not showing)
    if node.get(f"{{{st_ns}}}visible", "false") != "true":
        return False

    # Capability check (at least one of enabled/editable/expandable/checkable)
    enabled = node.get(f"{{{st_ns}}}enabled", "false") == "true"
    editable = node.get(f"{{{st_ns}}}editable", "false") == "true"
    expandable = node.get(f"{{{st_ns}}}expandable", "false") == "true"
    checkable = node.get(f"{{{st_ns}}}checkable", "false") == "true"
    if not (enabled or editable or expandable or checkable):
        return False

    # Name/text check
    name = node.get("name", "")
    text = node.text
    if not name and (text is None or len(text) == 0):
        return False

    # Coordinate validity check
    coords_str = node.get(f"{{{cp_ns}}}screencoord", "(-1, -1)")
    size_str = node.get(f"{{{cp_ns}}}size", "(-1, -1)")
    try:
        coords = tuple(map(int, coords_str.strip("()").split(", ")))
        sizes = tuple(map(int, size_str.strip("()").split(", ")))
    except (ValueError, TypeError):
        return False

    return coords[0] >= 0 and coords[1] >= 0 and sizes[0] > 0 and sizes[1] > 0


def _filter_nodes(root: _Element) -> List[_Element]:
    """Filter XML tree nodes using OSWorld's heuristic.

    Port of OSWorld heuristic_retrieve.py filter_nodes().
    """
    return [node for node in root.iter() if _judge_node(node)]


# ── TSV linearization (port of OSWorld agent.py linearize_accessibility_tree) ─


def _linearize_tree(filtered_nodes: List[_Element]) -> str:
    """Linearize filtered nodes into TSV format.

    Port of OSWorld mm_agents/agent.py L71-L118.
    Header: tag  name  text  class  description  position (top-left x&y)  size (w&h)
    """
    cp_ns = NS_MAP["cp"]
    val_ns = NS_MAP["val"]
    class_ns = NS_MAP["class"]

    lines = ["tag\tname\ttext\tclass\tdescription\tposition (top-left x&y)\tsize (w&h)"]

    for node in filtered_nodes:
        # Text extraction (OSWorld logic: EditWrapper uses val:value)
        if node.text:
            text = (
                node.text if '"' not in node.text else '"{}"'.format(node.text.replace('"', '""'))
            )
        elif node.get(f"{{{class_ns}}}class", "").endswith("EditWrapper") and node.get(
            f"{{{val_ns}}}value"
        ):
            node_text = node.get(f"{{{val_ns}}}value", "")
            text = (
                node_text if '"' not in node_text else '"{}"'.format(node_text.replace('"', '""'))
            )
        else:
            text = '""'

        lines.append(
            "{}\t{}\t{}\t{}\t{}\t{}\t{}".format(
                node.tag,
                node.get("name", ""),
                text,
                node.get(f"{{{class_ns}}}class", ""),
                "",  # description not available in Windows
                node.get(f"{{{cp_ns}}}screencoord", ""),
                node.get(f"{{{cp_ns}}}size", ""),
            )
        )

    return "\n".join(lines)


# ── Tree service class ───────────────────────────────────────────────────────


class Tree:
    """Accessibility tree service using pywinauto (OSWorld-compatible)."""

    def __init__(self, desktop: "Desktop"):
        self.desktop = weakref.proxy(desktop)
        self.screen_size = desktop.get_screen_size()
        self.tree_state: Optional[TreeState] = None

    def get_state(
        self,
        active_window_handle: Optional[int] = None,
        other_windows_handles: Optional[List[int]] = None,
        use_dom: bool = False,
    ) -> TreeState:
        """Build accessibility tree and return filtered TSV state.

        Args:
            active_window_handle: Ignored (pywinauto enumerates all windows).
            other_windows_handles: Ignored (pywinauto enumerates all windows).
            use_dom: Ignored (DOM detection not implemented in OSWorld compat mode).

        Returns:
            TreeState with xml_tree, tsv_tree, filtered_nodes, and marks.
        """
        start_time = time()

        xml_tree_str, filtered_nodes, tsv_str, marks = self._build_tree()

        self.tree_state = TreeState(
            xml_tree=xml_tree_str,
            tsv_tree=tsv_str,
            filtered_nodes=filtered_nodes,
            marks=marks,
        )

        elapsed = time() - start_time
        logger.info(f"Tree State capture took {elapsed:.2f} seconds")
        return self.tree_state

    def _build_tree(self) -> Tuple[str, List[_Element], str, List[List[int]]]:
        """Build full a11y tree, filter, and linearize to TSV.

        Returns:
            (xml_string, filtered_nodes, tsv_string, marks)
        """
        from pywinauto import Desktop as PywinautoDesktop

        desktop = PywinautoDesktop(backend="uia")
        root_xml = lxml.etree.Element("desktop", nsmap=NS_MAP)

        # Build XML for each window (parallel, like OSWorld)
        windows = desktop.windows()

        # Explicitly include the Desktop (Progman) window which holds desktop
        # icons.  desktop.windows() skips it because Windows marks it as
        # hidden / 1×1, but its children (the icon ListItems) are real,
        # visible UI elements that agents need to interact with.
        try:
            progman = desktop.window(class_name="Progman")
            if progman.exists() and progman not in windows:
                windows.append(progman)
        except Exception:
            pass

        if windows:
            with concurrent.futures.ThreadPoolExecutor() as executor:
                futures = [
                    executor.submit(_create_pywinauto_node, wnd, set(), 1) for wnd in windows
                ]
                for future in concurrent.futures.as_completed(futures):
                    try:
                        xml_tree = future.result()
                        if xml_tree is not None:
                            root_xml.append(xml_tree)
                    except Exception as e:
                        logger.debug(f"Error building window tree: {e}")

        xml_str = lxml.etree.tostring(root_xml, encoding="unicode")

        # Filter nodes
        filtered = _filter_nodes(root_xml)

        # Linearize to TSV
        tsv = _linearize_tree(filtered)

        # Extract marks (bounding boxes for SoM)
        cp_ns = NS_MAP["cp"]
        marks: List[List[int]] = []
        for node in filtered:
            coords_str = node.get(f"{{{cp_ns}}}screencoord", "")
            size_str = node.get(f"{{{cp_ns}}}size", "")
            if coords_str and size_str:
                try:
                    coords = tuple(map(int, coords_str.strip("()").split(", ")))
                    size = tuple(map(int, size_str.strip("()").split(", ")))
                    marks.append([coords[0], coords[1], size[0], size[1]])
                except (ValueError, TypeError):
                    pass

        return xml_str, filtered, tsv, marks

    def _on_focus_change(self, sender: Any):
        """Handle focus change events (kept for watchdog compatibility).

        In OSWorld compat mode, this is a no-op since we don't use the custom
        uia module's Control.CreateControlFromElement.
        """
        pass
