import asyncio

from fastapi import APIRouter
from fastapi.responses import StreamingResponse

import time

from app.config import SSE_KEEPALIVE_SECONDS, SSE_POLL_SECONDS
import app.state as state

router = APIRouter()


async def _event_generator():
    """Generate SSE events for a connected client."""
    client = {"needs_update": False, "last_version": state.data_version}
    state.connected_clients.append(client)
    last_send_time = time.monotonic()

    try:
        while True:
            if client["needs_update"] or client["last_version"] != state.data_version:
                client["needs_update"] = False
                client["last_version"] = state.data_version
                last_send_time = time.monotonic()
                yield f"data: {state.data_version}\n\n"
            elif time.monotonic() - last_send_time >= max(SSE_KEEPALIVE_SECONDS, 1):
                # Keep idle SSE connections alive through proxies/load balancers.
                last_send_time = time.monotonic()
                yield ": keepalive\n\n"

            await asyncio.sleep(max(SSE_POLL_SECONDS, 0.1))
    finally:
        if client in state.connected_clients:
            state.connected_clients.remove(client)


@router.get("/events")
async def sse_events():
    """SSE endpoint for real-time data update notifications."""
    return StreamingResponse(
        _event_generator(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no",
        },
    )
