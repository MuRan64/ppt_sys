"""Keep temporary book renderer hooks local to one generation call."""

from __future__ import annotations

from functools import wraps
import sys
import threading
from typing import Callable, ParamSpec, TypeVar


_P = ParamSpec("_P")
_R = TypeVar("_R")
_GENERATION_LOCK = threading.RLock()
_RENDERER_MODULES = (
    "test_md_to_ppt",
    "test_sanbu_md_to_ppt",
    "test_wenyan_yuedu_jiafen_md_to_ppt",
    "test_zuhe_jiafen_md_to_ppt",
    "test_zhuji_100_md_to_ppt",
    "test_yilun_xinfangan_md_to_ppt",
    "test_bibei_gushiwen_md_to_ppt",
)


def isolated_generation(function: Callable[_P, _R]) -> Callable[_P, _R]:
    """Restore renderer bindings and hook flags, including after a failure.

    Legacy book renderers install functions and constants on shared modules.
    Snapshot their bindings rather than reloading modules: existing dataclasses
    and function references must keep their identities. The lock also prevents
    overlapping calls from observing another book's temporary hooks.
    """

    @wraps(function)
    def wrapped(*args: _P.args, **kwargs: _P.kwargs) -> _R:
        with _GENERATION_LOCK:
            snapshots = []
            for name in _RENDERER_MODULES:
                module = sys.modules.get(name)
                if module is not None:
                    snapshots.append((module, {k: v for k, v in vars(module).items() if not k.startswith("__")}))
            try:
                return function(*args, **kwargs)
            finally:
                for module, previous in snapshots:
                    for key in tuple(vars(module)):
                        if not key.startswith("__") and key not in previous:
                            delattr(module, key)
                    vars(module).update(previous)

    return wrapped
