"""Local duplicate prevention. This does not promise server-side exactly-once delivery."""

import hashlib
import json
import os
import tempfile
from datetime import UTC, datetime
from pathlib import Path

from .errors import OutlookError
from .session import state_dir


def atomic_json(path, data):
    path = Path(path)
    path.parent.mkdir(mode=0o700, parents=True, exist_ok=True)
    fd, temporary = tempfile.mkstemp(dir=path.parent, prefix=".pending-")
    try:
        with os.fdopen(fd, "w") as stream:
            json.dump(data, stream, ensure_ascii=False, indent=2)
            stream.write("\n")
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)


def execute(request_id, account, payload, action, *, compatible_payloads=()):
    if not request_id:
        return action()
    if not isinstance(request_id, str) or not request_id.strip() or len(request_id) > 200:
        raise OutlookError("invalid_request_id", "Request IDs must contain 1–200 characters.")

    def fingerprint_for(value):
        return hashlib.sha256(
            json.dumps(
                {"account": account.casefold(), "payload": value},
                sort_keys=True,
                ensure_ascii=False,
            ).encode()
        ).hexdigest()

    fingerprint = fingerprint_for(payload)
    key = hashlib.sha256(request_id.encode()).hexdigest()
    path = state_dir() / "operations" / f"{key}.json"
    if path.exists():
        try:
            saved = json.loads(path.read_text())
            if (
                not isinstance(saved, dict)
                or not isinstance(saved.get("fingerprint"), str)
                or (saved.get("state") == "completed" and not isinstance(saved.get("result"), dict))
            ):
                raise ValueError("invalid operation record")
        except (OSError, ValueError):
            raise OutlookError(
                "operation_state_invalid",
                "Cannot read the prior operation record; inspect it before retrying.",
            ) from None
        # Compatibility only recognizes existing records; it never executes a
        # legacy payload, retries a pending write, or crosses an account boundary.
        accepted = {fingerprint, *(fingerprint_for(p) for p in compatible_payloads)}
        if saved.get("fingerprint") not in accepted:
            raise OutlookError(
                "request_id_conflict",
                "This request ID was already used with different input or a different account.",
            )
        if saved.get("state") == "completed":
            return {**saved["result"], "replayed": True}
        raise OutlookError(
            "write_outcome_unknown",
            "This request ID has an unfinished attempt. Inspect the affected service before using a new request ID.",
            request_id=request_id,
        )
    record = {
        "request_id": request_id,
        "fingerprint": fingerprint,
        "state": "pending",
        "started_at": datetime.now(UTC).isoformat(),
    }
    atomic_json(path, record)
    # Keep pending on every failure: even a local interruption might follow a successful write.
    result = action()
    atomic_json(path, {**record, "state": "completed", "result": result})
    return result


def lookup(request_id):
    """Read one private operation record without contacting Microsoft."""
    if not request_id or len(request_id) > 200:
        raise OutlookError("invalid_request_id", "Request IDs must contain 1–200 characters.")
    key = hashlib.sha256(request_id.encode()).hexdigest()
    path = state_dir() / "operations" / f"{key}.json"
    try:
        data = json.loads(path.read_text())
        if not isinstance(data, dict):
            raise ValueError("invalid operation record")
    except FileNotFoundError:
        raise OutlookError(
            "operation_not_found", "No local record exists for this request ID."
        ) from None
    except (ValueError, OSError):
        raise OutlookError("operation_state_invalid", "Cannot read the operation record.") from None
    return {"request_id": request_id, **{k: v for k, v in data.items() if k != "fingerprint"}}
