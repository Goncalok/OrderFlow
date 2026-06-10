from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from api._cloud_state_store import save_json_state, supabase_enabled  # noqa: E402


MIGRATIONS = [
    ("team_state:v1", ROOT / "data" / "team_state.json"),
    ("shortage_sessions:v1", ROOT / "data" / "shortage_sessions.json"),
    ("shortage_day_sessions:v1", ROOT / "data" / "shortage_day_sessions.json"),
]


def main() -> None:
    if not supabase_enabled():
        raise SystemExit("Set SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY before running this migration.")

    for key, path in MIGRATIONS:
        if not path.is_file():
            print(f"skip {key}: {path} not found")
            continue
        data = json.loads(path.read_text(encoding="utf-8"))
        save_json_state(key, data)
        print(f"migrated {key} from {path}")


if __name__ == "__main__":
    main()
