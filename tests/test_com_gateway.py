from __future__ import annotations

import pytest

from sapscriptwizard.core.com_gateway import ComGateway, RetryPolicy
from sapscriptwizard.core.errors import SapError, SapErrorCode


def test_com_gateway_retries_then_succeeds():
    calls = []

    def flaky(arg):
        calls.append(arg)
        if len(calls) < 2:
            raise RuntimeError("fail")
        return arg * 2

    gateway = ComGateway(
        retry_policy=RetryPolicy(max_attempts=3, base_delay=0),
        retry_exceptions=(RuntimeError,),
    )
    result = gateway.call(flaky, 21)
    assert result == 42
    assert len(calls) == 2


def test_com_gateway_raises_sap_error_after_retries():
    def always_fail():
        raise ValueError("boom")

    gateway = ComGateway(retry_policy=RetryPolicy(max_attempts=1, base_delay=0))
    with pytest.raises(SapError) as exc:
        gateway.call(always_fail)
    assert exc.value.code is SapErrorCode.COM_FAIL
