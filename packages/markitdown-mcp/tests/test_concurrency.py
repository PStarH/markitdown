import asyncio
import time
from unittest.mock import Mock

from markitdown_mcp import __main__ as server


def test_a_conversion_does_not_block_the_event_loop(monkeypatch):
    """The HTTP transport serves concurrent requests from one event loop.

    A conversion that runs on the loop stalls every other request for its whole
    duration. Sync tools are dispatched to a worker thread by the MCP SDK, so
    keep this tool sync and pin the behavior with a heartbeat.
    """

    class SlowConverter:
        def convert_uri(self, uri: str) -> Mock:
            time.sleep(0.3)
            result = Mock()
            result.markdown = "ok"
            return result

    monkeypatch.setattr(server, "MarkItDown", Mock(return_value=SlowConverter()))

    async def scenario() -> int:
        ticks = 0

        async def heartbeat() -> None:
            nonlocal ticks
            while True:
                await asyncio.sleep(0.01)
                ticks += 1

        task = asyncio.create_task(heartbeat())
        await asyncio.sleep(0)
        await server.mcp.call_tool(
            "convert_to_markdown", {"uri": "data:text/plain,hello"}
        )
        task.cancel()
        return ticks

    # A blocked loop yields about one tick; an offloaded conversion yields ~30.
    assert asyncio.run(scenario()) >= 10
