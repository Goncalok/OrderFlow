from __future__ import annotations

from shortage_app import ShortageHandler


class handler(ShortageHandler):
    def _normalize_path(self) -> None:
        if self.path.startswith("/shortage/api/"):
            self.path = "/api/" + self.path.removeprefix("/shortage/api/")
        elif self.path == "/shortage/api":
            self.path = "/api"

    def do_GET(self) -> None:
        self._normalize_path()
        super().do_GET()

    def do_POST(self) -> None:
        self._normalize_path()
        super().do_POST()

    def do_PUT(self) -> None:
        self._normalize_path()
        super().do_PUT()

    def do_DELETE(self) -> None:
        self._normalize_path()
        super().do_DELETE()
