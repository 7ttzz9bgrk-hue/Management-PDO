from datetime import datetime, timezone

from fastapi import APIRouter

import app.state as state

router = APIRouter()


@router.get('/health')
async def health():
    return {
        'status': 'ok',
        'data_version': state.data_version,
        'last_updated': state.cached_data.get('last_updated'),
        'timestamp': datetime.now(timezone.utc).isoformat(),
    }
