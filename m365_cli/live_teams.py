"""Opt-in CLI live validation. Creates only labelled self/channel test traffic."""

import argparse
import json
import os
import secrets
import subprocess
import sys
from pathlib import Path

from m365_cli.operations import atomic_json
from m365_cli.session import state_dir


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--account", required=True)
    parser.add_argument("--channel", help="Explicitly authorized standard channel ID; optional")
    parser.add_argument("--allow-test-writes", action="store_true", required=True)
    args = parser.parse_args()
    marker = "M365 CLI validation " + secrets.token_hex(6)
    journal_path = state_dir() / ("teams-live-" + marker.rsplit(" ", 1)[-1] + ".json")
    journal = {"account": args.account, "marker": marker, "operations": [], "cleanup": []}
    root = Path(__file__).resolve().parents[1]
    env = {**os.environ, "PYTHONPATH": str(root)}
    for previous in state_dir().glob("teams-live-*.json"):
        old = json.loads(previous.read_text())
        if any(c.get("state") != "cleaned" for c in old.get("cleanup", [])) or any(
            o.get("state") == "pending" or o.get("error", {}).get("code") == "write_outcome_unknown"
            for o in old.get("operations", [])
        ):
            raise RuntimeError("Prior Teams validation needs write/cleanup reconciliation.")
    tracked = []
    group = None

    def save():
        atomic_json(journal_path, journal)

    def invoke(argv, *, mutation=False, expected_error=None, replay=None):
        request_id = replay or marker.replace(" ", "-") + "-" + str(len(journal["operations"]))
        if mutation:
            argv = [*argv, "--request-id", request_id]
        op = {"arguments": argv, "state": "pending"}
        journal["operations"].append(op)
        save()
        result = subprocess.run(
            [sys.executable, "-m", "m365_cli.cli", "--account", args.account, "teams", *argv],
            cwd=root,
            env=env,
            capture_output=True,
            text=True,
            timeout=100,
        )
        try:
            payload = json.loads(result.stdout)
        except ValueError:
            raise RuntimeError("CLI returned no valid JSON; inspect the private journal") from None
        if expected_error:
            if payload.get("ok") or payload.get("error", {}).get("code") != expected_error:
                raise RuntimeError("Expected error was not returned")
            op.update(state="expected_error", code=expected_error)
            save()
            return None
        if not payload.get("ok"):
            op.update(state="failed", error=payload.get("error"))
            save()
            raise RuntimeError("CLI failed: " + payload.get("error", {}).get("code", "unknown"))
        op.update(state="completed", result=payload["data"])
        save()
        return payload["data"]

    def send(target, body, reply=None):
        argv = ["message", "reply" if reply else "send", *target, "--body", body]
        if reply:
            argv += ["--reply-to", reply]
        result = invoke(argv, mutation=True)
        item = {"target": target, "id": result["message"]["id"], "state": "created"}
        tracked.append(item)
        journal["cleanup"].append(item)
        save()
        return result

    try:
        invoke(["status"])
        self_chat = invoke(["chat", "self"])["id"]
        invoke(["person", "get", args.account])
        empty = invoke(["person", "search", marker.replace(" ", "")])
        assert not empty["items"]
        invoke(["chat", "list", "--limit", "1"])
        invoke(["chat", "list", "--limit", "1", "--offset", "1"])
        invoke(["chat", "members", self_chat])
        reused = invoke(["chat", "create", "--to", args.account], mutation=True)
        assert not reused["created"]
        target = ["--chat", self_chat]
        history = invoke(["message", "list", *target, "--limit", "1"])
        if history["next_cursor"]:
            invoke(["message", "list", *target, "--limit", "1", "--cursor", history["next_cursor"]])
        sent = send(target, marker)
        mid, version = sent["message"]["id"], sent["message"]["version"]
        send_id = journal["operations"][-1]["arguments"][-1]
        replayed = invoke(
            ["message", "send", *target, "--body", marker], mutation=True, replay=send_id
        )
        assert replayed.get("replayed") and replayed["message"]["id"] == mid
        invoke(["message", "get", mid, *target])
        invoke(
            [
                "message",
                "update",
                mid,
                *target,
                "--body",
                marker + " updated",
                "--if-version",
                str(version),
            ],
            mutation=True,
        )
        invoke(
            [
                "message",
                "update",
                mid,
                *target,
                "--body",
                "must not be written",
                "--if-version",
                str(version),
            ],
            mutation=True,
            expected_error="conflict",
        )
        invoke(
            ["message", "delete", mid, *target, "--if-version", str(version)],
            mutation=True,
            expected_error="unsupported_precondition",
        )
        print(
            "Self-chat send/read/edit, stale edit rejection, and duplicate replay passed.",
            flush=True,
        )
        teams = invoke(["team", "list", "--limit", "500"])
        if args.channel:
            channel = invoke(["channel", "get", args.channel])
            team = channel.get("parentTeamId")
            if team:
                invoke(["channel", "list", "--team", team, "--limit", "500"])
            assert teams["items"]
            ct = ["--channel", args.channel]
            root_post = send(ct, marker + " channel")
            reply = send(ct, marker + " reply", root_post["message"]["id"])
            invoke(
                ["message", "list", *ct, "--reply-to", root_post["message"]["id"], "--limit", "10"]
            )
            invoke(
                [
                    "message",
                    "update",
                    reply["message"]["id"],
                    *ct,
                    "--body",
                    marker + " reply updated",
                    "--if-version",
                    str(reply["message"]["version"]),
                ],
                mutation=True,
            )
            print("Channel post/reply/read/edit passed.", flush=True)
        created = invoke(["chat", "create", "--group", "--topic", marker], mutation=True)
        group = {"id": created["chat"]["id"], "state": "created", "kind": "group"}
        journal["cleanup"].append(group)
        save()
        invoke(["chat", "rename", group["id"], "--topic", marker + " updated"], mutation=True)
        print("Self-only group creation and rename passed.", flush=True)
    finally:
        failed = False
        for item in reversed(tracked):
            try:
                invoke(["message", "delete", item["id"], *item["target"]], mutation=True)
                result = invoke(["message", "get", item["id"], *item["target"]])
                assert result["deleted"] and not result.get("content")
                item["state"] = "cleaned"
                save()
            except Exception:
                failed = True
                item["state"] = "cleanup_pending"
                save()
        if group:
            try:
                invoke(["chat", "leave", group["id"]], mutation=True)
                group["state"] = "cleaned"
                save()
            except Exception:
                failed = True
                group["state"] = "cleanup_pending"
                save()
        if failed:
            raise RuntimeError("Cleanup needs attention; inspect the private Teams live journal")
        print(
            "All known test messages were soft-deleted and independently verified; test group left.",
            flush=True,
        )
    from m365_cli.cli import parser as command_parser
    from m365_cli.maintainer import command_path

    journal["command_coverage"] = sorted(
        {
            command_path(command_parser().parse_args(["teams", *o["arguments"]]))
            for o in journal["operations"]
            if o["state"] == "completed"
        }
    )
    journal["cleanup_complete"] = all(c["state"] == "cleaned" for c in journal["cleanup"])
    journal["result"] = "passed"
    save()
    print("Teams CLI live validation passed.", flush=True)


if __name__ == "__main__":
    main()
