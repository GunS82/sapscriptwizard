from __future__ import annotations

import pytest

from sapscriptwizard.core.com_gateway import ComGateway, RetryPolicy
from sapscriptwizard.core.errors import ErrorCode, SapError


def test_invoke_retries_and_raises():
    gateway = ComGateway(RetryPolicy(max_attempts=2, base_delay=0, backoff_factor=1))
    calls = {"count": 0}

    def failing() -> None:
        calls["count"] += 1
        raise RuntimeError("boom")

    with pytest.raises(SapError) as excinfo:
        gateway.invoke(failing)
    assert excinfo.value.code is ErrorCode.COM_FAIL
    assert calls["count"] == 2


def test_call_method():
    class Dummy:
        def method(self, arg: str) -> str:
            return arg.upper()

    gateway = ComGateway(RetryPolicy(max_attempts=1))
    dummy = Dummy()
    assert gateway.call_method(dummy, "method", "ok") == "OK"
