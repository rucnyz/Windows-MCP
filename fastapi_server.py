"""
FastAPI HTTP wrapper for the new Windows-MCP Desktop service.

Maps the legacy HTTP API endpoints (used by DecodingTrust-Agent evaluation framework)
to the new Desktop service class methods.

Usage:
    uv run fastapi_server.py [--port PORT]
"""

from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from typing import Literal, Optional, List, Union
from textwrap import dedent
import base64
import uvicorn
import sys
import os

# Add src to path so we can import windows_mcp
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

from windows_mcp.desktop.service import Desktop

import pyautogui as pg
import pyperclip as pc

pg.FAILSAFE = False
pg.PAUSE = 1.0

desktop = Desktop()
windows_version = desktop.get_windows_version()
default_language = desktop.get_default_language()

app = FastAPI(
    title="Windows MCP API",
    description=f"FastAPI server providing tools to interact with {windows_version} desktop",
    version="2.0.0",
)


class LaunchToolRequest(BaseModel):
    name: str


class PowershellToolRequest(BaseModel):
    command: str
    timeout: int = 30


class StateToolRequest(BaseModel):
    use_vision: bool = False


class ClipboardToolRequest(BaseModel):
    mode: Literal["copy", "paste", "get", "set"]
    text: Optional[str] = None


class ClickToolRequest(BaseModel):
    loc: List[int]
    button: Literal["left", "right", "middle"] = "left"
    clicks: int = 1


class TypeToolRequest(BaseModel):
    loc: List[int]
    text: str
    clear: bool = False
    press_enter: bool = False


class ResizeToolRequest(BaseModel):
    size: Optional[List[int]] = None
    loc: Optional[List[int]] = None


class SwitchToolRequest(BaseModel):
    name: str


class ScrollToolRequest(BaseModel):
    loc: Optional[List[int]] = None
    type: Literal["horizontal", "vertical"] = "vertical"
    direction: Literal["up", "down", "left", "right"] = "down"
    wheel_times: int = 1


class DragToolRequest(BaseModel):
    from_loc: List[int]
    to_loc: List[int]


class MoveToolRequest(BaseModel):
    to_loc: List[int]


class ShortcutToolRequest(BaseModel):
    shortcut: Union[List[str], str]


class KeyToolRequest(BaseModel):
    key: str


class WaitToolRequest(BaseModel):
    duration: int


class ScrapeToolRequest(BaseModel):
    url: str


class ReadFileRequest(BaseModel):
    path: str
    encoding: str = "utf-8"


class EditFileRequest(BaseModel):
    path: str
    content: str
    encoding: str = "utf-8"
    mode: Literal["overwrite", "append"] = "overwrite"


class ToolResponse(BaseModel):
    result: Union[str, dict, List]
    status: str = "success"


@app.get("/")
async def root():
    return {
        "message": "Windows MCP FastAPI Server",
        "version": windows_version,
        "default_language": default_language,
    }


@app.post("/tools/launch", response_model=ToolResponse)
async def launch_tool(request: LaunchToolRequest):
    try:
        response, status, pid = desktop.launch_app(request.name.lower())
        if status != 0:
            return ToolResponse(result=response)
        consecutive_waits = 2
        for _ in range(consecutive_waits):
            if not desktop.is_app_running(request.name):
                pg.sleep(1.0)
            else:
                return ToolResponse(result=response)
        return ToolResponse(result=f"Launching {request.name.title()} wait for it to come load.")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/tools/powershell", response_model=ToolResponse)
async def powershell_tool(request: PowershellToolRequest):
    try:
        response, status_code = desktop.execute_command(request.command, timeout=request.timeout)
        return ToolResponse(result=f"Response: {response}\nStatus Code: {status_code}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/tools/state", response_model=ToolResponse)
async def state_tool(request: StateToolRequest):
    try:
        desktop_state = desktop.get_state(use_vision=request.use_vision)
        interactive_elements = desktop_state.tree_state.interactive_elements_to_string()
        scrollable_elements = desktop_state.tree_state.scrollable_elements_to_string()
        windows = desktop_state.windows_to_string()
        active_window = desktop_state.active_window_to_string()

        state_text = dedent(f"""
        Default Language of User:
        {default_language} with encoding: {desktop.encoding}

        Focused Window:
        {active_window}

        Opened Windows:
        {windows}

        List of Interactive Elements:
        {interactive_elements or "No interactive elements found."}

        List of Scrollable Elements:
        {scrollable_elements or "No scrollable elements found."}
        """)

        result = {"state": state_text}
        if request.use_vision and desktop_state.screenshot:
            result["screenshot"] = base64.b64encode(desktop_state.screenshot).decode("utf-8")

        return ToolResponse(result=result)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/tools/clipboard", response_model=ToolResponse)
async def clipboard_tool(request: ClipboardToolRequest):
    try:
        if request.mode in ("copy", "set"):
            if request.text:
                pc.copy(request.text)
                return ToolResponse(result=f'Copied "{request.text}" to clipboard')
            else:
                raise HTTPException(status_code=400, detail="No text provided to copy")
        elif request.mode in ("paste", "get"):
            clipboard_content = pc.paste()
            return ToolResponse(result=f'Clipboard Content: "{clipboard_content}"')
        else:
            raise HTTPException(
                status_code=400,
                detail='Invalid mode. Use "copy"/"set" or "paste"/"get".',
            )
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/tools/click", response_model=ToolResponse)
async def click_tool(request: ClickToolRequest):
    try:
        if len(request.loc) != 2:
            raise HTTPException(
                status_code=400,
                detail="Location must be a list of exactly 2 integers [x, y]",
            )
        x, y = request.loc[0], request.loc[1]
        desktop.click(loc=(x, y), button=request.button, clicks=request.clicks)
        num_clicks = {0: "Hover", 1: "Single", 2: "Double", 3: "Triple"}
        return ToolResponse(
            result=f"{num_clicks.get(request.clicks, 'Multi')} {request.button} Clicked at ({x},{y})."
        )
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/tools/type", response_model=ToolResponse)
async def type_tool(request: TypeToolRequest):
    try:
        if len(request.loc) != 2:
            raise HTTPException(
                status_code=400,
                detail="Location must be a list of exactly 2 integers [x, y]",
            )
        x, y = request.loc[0], request.loc[1]
        desktop.type(
            loc=(x, y),
            text=request.text,
            clear=request.clear,
            press_enter=request.press_enter,
        )
        return ToolResponse(result=f"Typed {request.text} at ({x},{y}).")
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/tools/resize", response_model=ToolResponse)
async def resize_tool(request: ResizeToolRequest):
    try:
        if request.size is not None and len(request.size) != 2:
            raise HTTPException(
                status_code=400,
                detail="Size must be a list of exactly 2 integers [width, height]",
            )
        if request.loc is not None and len(request.loc) != 2:
            raise HTTPException(
                status_code=400,
                detail="Location must be a list of exactly 2 integers [x, y]",
            )
        size_tuple = tuple(request.size) if request.size is not None else None
        loc_tuple = tuple(request.loc) if request.loc is not None else None
        response, _ = desktop.resize_app(size_tuple, loc_tuple)
        return ToolResponse(result=response)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/tools/switch", response_model=ToolResponse)
async def switch_tool(request: SwitchToolRequest):
    try:
        response, status = desktop.switch_app(request.name)
        return ToolResponse(result=response)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/tools/scroll", response_model=ToolResponse)
async def scroll_tool(request: ScrollToolRequest):
    try:
        loc_tuple = None
        if request.loc:
            if len(request.loc) != 2:
                raise HTTPException(
                    status_code=400,
                    detail="Location must be a list of exactly 2 integers [x, y]",
                )
            loc_tuple = (request.loc[0], request.loc[1])

        response = desktop.scroll(
            loc=loc_tuple,
            type=request.type,
            direction=request.direction,
            wheel_times=request.wheel_times,
        )
        if response:
            return ToolResponse(result=response)
        return ToolResponse(
            result=f"Scrolled {request.type} {request.direction} by {request.wheel_times} wheel times."
        )
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/tools/drag", response_model=ToolResponse)
async def drag_tool(request: DragToolRequest):
    try:
        if len(request.from_loc) != 2:
            raise HTTPException(
                status_code=400,
                detail="from_loc must be a list of exactly 2 integers [x, y]",
            )
        if len(request.to_loc) != 2:
            raise HTTPException(
                status_code=400,
                detail="to_loc must be a list of exactly 2 integers [x, y]",
            )
        x1, y1 = request.from_loc[0], request.from_loc[1]
        x2, y2 = request.to_loc[0], request.to_loc[1]
        pg.moveTo(x1, y1)
        desktop.drag(loc=(x2, y2))
        return ToolResponse(result=f"Dragged from ({x1},{y1}) to ({x2},{y2}).")
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/tools/move", response_model=ToolResponse)
async def move_tool(request: MoveToolRequest):
    try:
        if len(request.to_loc) != 2:
            raise HTTPException(
                status_code=400,
                detail="to_loc must be a list of exactly 2 integers [x, y]",
            )
        x, y = request.to_loc[0], request.to_loc[1]
        desktop.move(loc=(x, y))
        return ToolResponse(result=f"Moved the mouse pointer to ({x},{y}).")
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/tools/shortcut", response_model=ToolResponse)
async def shortcut_tool(request: ShortcutToolRequest):
    try:
        # Support both old format (list of keys) and new format (single string with +)
        if isinstance(request.shortcut, list):
            shortcut_str = "+".join(request.shortcut)
        else:
            shortcut_str = request.shortcut
        desktop.shortcut(shortcut_str)
        return ToolResponse(result=f"Pressed {shortcut_str}.")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/tools/key", response_model=ToolResponse)
async def key_tool(request: KeyToolRequest):
    try:
        pg.press(request.key)
        return ToolResponse(result=f"Pressed the key {request.key}.")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/tools/wait", response_model=ToolResponse)
async def wait_tool(request: WaitToolRequest):
    try:
        pg.sleep(request.duration)
        return ToolResponse(result=f"Waited for {request.duration} seconds.")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/tools/scrape", response_model=ToolResponse)
async def scrape_tool(request: ScrapeToolRequest):
    try:
        content = desktop.scrape(request.url)
        return ToolResponse(result=f"Scraped the contents of the entire webpage:\n{content}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/tools/read", response_model=ToolResponse)
async def read_file_tool(request: ReadFileRequest):
    try:
        path = request.path.replace("'", "''")
        response, status_code = desktop.execute_command(
            f"Get-Content -Path '{path}' -Raw -Encoding {request.encoding}",
            timeout=10,
        )
        if status_code != 0:
            return ToolResponse(result=f"Error reading {request.path}: {response}", status="error")
        return ToolResponse(result=response)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/tools/edit", response_model=ToolResponse)
async def edit_file_tool(request: EditFileRequest):
    try:
        path = request.path.replace("'", "''")
        content_b64 = base64.b64encode(request.content.encode("utf-8")).decode("ascii")
        if request.mode == "append":
            ps_cmd = (
                f"$bytes = [Convert]::FromBase64String('{content_b64}'); "
                f"$text = [System.Text.Encoding]::UTF8.GetString($bytes); "
                f"Add-Content -Path '{path}' -Value $text -Encoding {request.encoding}"
            )
        else:
            ps_cmd = (
                f"$bytes = [Convert]::FromBase64String('{content_b64}'); "
                f"$text = [System.Text.Encoding]::UTF8.GetString($bytes); "
                f"Set-Content -Path '{path}' -Value $text -Encoding {request.encoding}"
            )
        response, status_code = desktop.execute_command(ps_cmd, timeout=10)
        if status_code != 0:
            return ToolResponse(result=f"Error writing {request.path}: {response}", status="error")
        return ToolResponse(result=f"Written to {request.path} ({request.mode})")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


if __name__ == "__main__":
    port = 8005
    if "--port" in sys.argv:
        idx = sys.argv.index("--port")
        if idx + 1 < len(sys.argv):
            port = int(sys.argv[idx + 1])
    uvicorn.run(app, host="0.0.0.0", port=port)
