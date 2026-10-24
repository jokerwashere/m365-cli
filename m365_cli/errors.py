class OutlookError(Exception):
    def __init__(self, code, message, **details):
        super().__init__(message)
        self.code, self.details = code, details

    def as_dict(self):
        result = {"code": self.code, "message": str(self), **self.details}
        if self.code in (
            "api_changed",
            "unexpected_response",
            "missing_version",
            "missing_change_key",
        ):
            result["recovery"] = (
                "Stop this capability and run maintainer check; an adapter update may be required."
            )
        elif self.code in ("write_outcome_unknown", "partial_completion", "readback_failed"):
            result["recovery"] = (
                "Inspect the service and operation journal before another write. Never automatically retry with a new request ID."
            )
        elif self.code in ("login_required", "session_rejected"):
            result["recovery"] = (
                "Check the pinned account and complete sign-in in the saved browser. No automatic write retry was made."
            )
        return result
