"""Small, additive-field-tolerant contracts shared by service adapters."""

from .errors import OutlookError


def changed(contract, field, *, write=False):
    raise OutlookError(
        "write_outcome_unknown" if write else "api_changed",
        "The service response no longer satisfies the expected contract."
        + (
            " Inspect the affected item before retrying the write."
            if write
            else " Run maintainer check."
        ),
        contract=contract,
        field=field,
    )


def object_value(value, contract, *, fields=(), write=False):
    if not isinstance(value, dict):
        changed(contract, "$", write=write)
    for field in fields:
        if not isinstance(value.get(field), str) or not value[field]:
            changed(contract, field, write=write)
    return value


def array_value(value, contract, *, fields=(), write=False):
    if not isinstance(value, list):
        changed(contract, "$[]", write=write)
    for item in value:
        object_value(item, contract, fields=fields, write=write)
    return value


def typed_field(value, key, kind, contract, *, optional=False):
    if optional and key not in value:
        return
    if type(value.get(key)) is not kind:
        changed(contract, key)
