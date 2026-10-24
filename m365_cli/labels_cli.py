"""Friendly category/catalog and fixed-slot Planner label commands."""

from .api import API, reference
from .contracts import changed
from .errors import OutlookError
from .labels import Categories, PlannerLabels, applied_categories, color, label_name, palette, slot
from .operations import execute
from .references import References
from .session import Session, read_config
from .task_validation import identifier
from .tasks_api import TasksAPI, normalized, q, review, version
from .tasks_session import TasksSession
from .validation import email


def register_category(root, write_options, *, action_dest="command"):
    sub = root.add_subparsers(dest=action_dest, required=True)
    for action in ("colors", "list", "get", "create", "update", "delete", "apply", "remove"):
        p = sub.add_parser(action)
        p.set_defaults(label_domain="category", label_action=action)
        if action not in ("colors", "list", "get"):
            write_options(p)
        if action in ("get", "update", "delete", "apply"):
            p.add_argument("id", help="Category provider ID from category list/get")
        if action == "remove":
            p.add_argument("id", nargs="?")
            p.add_argument("--name", help="Remove a name even when its definition was deleted")
        if action in ("create", "update"):
            p.add_argument("--name", required=action == "create")
            p.add_argument("--color", help="Name or preset0–preset24; see category colors")
        if action == "update":
            p.add_argument(
                "--definition-only",
                action="store_true",
                help="Leave existing item category names unchanged when renaming",
            )
        if action in ("update", "delete", "apply", "remove"):
            p.add_argument(
                "--if-version",
                help="Reviewed item version for apply/remove; category review hash for definition changes (not atomic)",
            )
        if action in ("apply", "remove"):
            target = p.add_mutually_exclusive_group(required=True)
            target.add_argument(
                "--task", dest="task_id", help="To Do task ID; also requires --list"
            )
            target.add_argument("--mail", dest="mail_id")
            target.add_argument(
                "--event",
                dest="event_id",
                help="Calendar occurrence ID; personal metadata, no notifications",
            )
            p.add_argument("--list", dest="list_id")
        if action == "list":
            p.add_argument("--query", help="Case-insensitive name substring")


def register_planner(root, write_options):
    sub = root.add_subparsers(dest="task_action", required=True)
    for action in (
        "colors",
        "list",
        "get",
        "create",
        "rename",
        "recolor",
        "delete",
        "apply",
        "remove",
    ):
        p = sub.add_parser(action)
        p.set_defaults(label_domain="planner", label_action=action, task_source="planner")
        if action != "colors":
            p.add_argument("--plan", dest="plan_id", required=True)
        if action not in ("colors", "list", "create"):
            p.add_argument("id", help="Fixed color slot: category1–category25")
        if action in ("create", "rename"):
            p.add_argument("--name", required=True)
        if action in ("create", "recolor"):
            p.add_argument("--color", required=True, help="Color name from planner label colors")
        if action not in ("colors", "list", "get"):
            write_options(p)
        if action not in ("colors", "list", "get", "create"):
            p.add_argument(
                "--if-version", help="Reviewed plan-details ETag; task ETag for apply/remove"
            )
        if action in ("create", "recolor", "delete"):
            p.add_argument(
                "--max-pages",
                type=int,
                default=20,
                help="Complete task scan budget (1–100); incomplete scans refuse writes",
            )
        if action == "delete":
            p.add_argument(
                "--remove-from-tasks",
                action="store_true",
                help="Remove this slot from all tasks in the plan before clearing its name",
            )
        if action in ("apply", "remove"):
            p.add_argument("--task", dest="task_id", required=True)


def prepare(args):
    if args.account:
        email(args.account)
    for key in ("id", "plan_id", "task_id", "mail_id", "event_id", "list_id"):
        value = getattr(args, key, None)
        if value is not None:
            identifier(value, key)
    if getattr(args, "name", None) is not None:
        label_name(args.name, category=args.label_domain == "category")
    if getattr(args, "color", None) is not None:
        args.color_value = color(args.color, args.label_domain)
    if getattr(args, "if_version", None) is not None:
        identifier(args.if_version, "version")
    if args.label_domain == "planner" and getattr(args, "id", None):
        slot(args.id)
    if getattr(args, "max_pages", 20) not in range(1, 101):
        raise OutlookError("invalid_argument", "--max-pages must be between 1 and 100.")
    if args.label_domain == "category":
        if args.label_action == "update" and args.name is None and args.color is None:
            raise OutlookError("invalid_argument", "Choose --name and/or --color.")
        if args.label_action == "remove" and bool(args.id) == bool(args.name):
            raise OutlookError("invalid_argument", "Choose a category ID or --name, exclusively.")
        if args.label_action in ("apply", "remove") and bool(args.task_id) != bool(args.list_id):
            raise OutlookError(
                "invalid_argument",
                "A To Do target requires both --task and --list; --list cannot be used for mail/events.",
            )
    return args


def category_result(refs, result):
    for item in result.get("items", [result]):
        refs.outlook_item(item, "category", previous=item.get("previous_id"))
    return refs.registry.finish(result)


def category_dispatch(session, args, refs):
    cats, action = Categories(session), args.label_action
    if action == "list":
        result = cats.listing()
        if args.query:
            result["items"] = [
                i for i in result["items"] if args.query.casefold() in i["name"].casefold()
            ]
    elif action == "get":
        result = cats.get(args.id)
    else:
        result = cats.mutate(
            action,
            identifier=getattr(args, "id", None),
            name=getattr(args, "name", None),
            color_id=getattr(args, "color_value", None),
            expected=getattr(args, "if_version", None),
        )
    result = category_result(refs, result)
    if action == "update" and result.get("previous_name") and not args.definition_only:
        try:
            ack = cats.call(
                "ApplyCategoryOperation",
                {
                    "categoryOperation": {
                        "Action": 0,
                        "CategoryName": result["previous_name"],
                        "RenameCategoryName": result["name"],
                    }
                },
            )
            if type(ack.get("IsSearchFolderReady")) is not bool:
                changed("outlook.category.rename_items", "IsSearchFolderReady", write=True)
        except OutlookError as exc:
            raise OutlookError(
                "partial_completion",
                "The category definition was renamed, but item rename has an uncertain outcome. Inspect before retrying.",
                category=result,
                cause=exc.as_dict(),
            ) from None
        result.update(
            item_assignments="rename_submitted",
            propagation_complete=False,
            propagation_note="Microsoft acknowledged the mailbox-wide operation; completion across all items is not independently verified.",
        )
    return result


def category_assignment(api, args, name):
    task = isinstance(api, TasksAPI)
    field = "categories" if task else "Categories"
    target = args.task_id if task else args.mail_id or args.event_id
    path = api.todo_path(args.list_id, target) if task else None
    current = api.entity("graph", path) if task else api.get(target, calendar=bool(args.event_id))
    expected = getattr(args, "if_version", None)
    if task:
        review(current, expected)
    else:
        api._kind(current, calendar=bool(args.event_id))
        reference(current, expected)
    values = current.get(field, [])
    if not isinstance(values, list) or any(not isinstance(v, str) for v in values):
        changed("labels.assignment", field)
    new = [v for v in values if v.casefold() != name.casefold()]
    if args.label_action == "apply":
        new.append(name)
    if new == values:
        return {**(normalized(current, "todo", "task") if task else current), "unchanged": True}
    if task:
        result = api.patch("graph", path, {field: new}, "todo", "task", version(current))
    else:
        result = api.update(
            target,
            {field: new},
            calendar=bool(args.event_id),
            metadata=True,
            expected=current["change_key"],
        )
    if result.get(field, []) != new:
        changed("labels.assignment.readback", field, write=True)
    return result


def planner_dispatch(api, args):
    labels, action = PlannerLabels(api, args.plan_id), args.label_action
    if action == "list":
        return labels.listing()
    if action == "get":
        return labels.item(labels.details(), args.id)
    if action in ("apply", "remove"):
        label = labels.item(labels.details(), args.id)
        if action == "apply" and not label["defined"]:
            raise OutlookError("label_not_defined", "Name this plan label before applying it.")
        path = "beta/planner/tasks/" + q(args.task_id)
        current = api.entity("graph", path)
        if current.get("planId") != args.plan_id:
            raise OutlookError("invalid_argument", "Task does not belong to the supplied plan.")
        review(current, args.if_version)
        enabled = action == "apply"
        values = applied_categories(current)
        if values.get(args.id, False) is enabled:
            return {**normalized(current, "planner", "task"), "unchanged": True}
        result = api.patch(
            "graph",
            path,
            {"appliedCategories": {args.id: enabled}},
            "planner",
            "task",
            version(current),
        )
        if applied_categories(result, write=True).get(args.id, False) is not enabled:
            changed("planner.label.assignment", "appliedCategories", write=True)
        return result
    return labels.mutate(
        action,
        getattr(args, "id", None),
        name=getattr(args, "name", None),
        target=getattr(args, "color_value", None),
        expected=getattr(args, "if_version", None),
        max_pages=getattr(args, "max_pages", 20),
        remove_from_tasks=getattr(args, "remove_from_tasks", False),
    )


def run(args, plan):
    if args.label_action == "colors":
        return palette(args.label_domain)
    account = args.account or read_config().get("account")
    if not account:
        raise OutlookError("account_required", "Pin the account with m365 auth login first.")
    if args.label_domain == "planner":
        with TasksSession(account) as session:
            api, refs = TasksAPI(session), References(session.account, session.tenant)

            def action():
                result = planner_dispatch(api, args)
                if result.get("source_label"):
                    refs.task(result["source_label"], "planner", "label", {"plan_id": args.plan_id})
                return refs.tasks(result, args)

            return (
                execute(
                    getattr(args, "request_id", None),
                    session.account,
                    {"service": "planner", "tenant": session.tenant, "command": plan(args)},
                    action,
                )
                if args.mutation
                else action()
            )

    def action():
        assignment = args.label_action in ("apply", "remove")
        # Playwright sync sessions cannot be nested. Both sessions verify this account.
        with Session() as session:
            session.authenticate(expected=account)
            refs = References(session.account)
            if not assignment:
                return category_dispatch(session, args, refs)
            name = (
                args.name
                if args.label_action == "remove" and args.name
                else Categories(session).get(args.id)["name"]
            )
            if not args.task_id:
                result = category_assignment(API(session), args, name)
                refs.outlook_item(result, "event" if args.event_id else "mail")
                return refs.registry.finish(result)
        with TasksSession(account) as session:
            refs = References(session.account, session.tenant)
            result = category_assignment(TasksAPI(session), args, name)
            refs.task(result, "todo", "task", {"list_id": args.list_id})
            return refs.registry.finish(result)

    return (
        execute(getattr(args, "request_id", None), account, plan(args), action)
        if args.mutation
        else action()
    )
