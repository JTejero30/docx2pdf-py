"""Process-tree timeout behavior."""

import subprocess

import pytest

from docx2pdf_py import processes
from docx2pdf_py.processes import run_process


def test_timeout_terminates_process_tree(monkeypatch):
    calls = []

    class FakeProcess:
        pid = 4321
        returncode = None

        def __init__(self):
            self.attempt = 0
            self.killed = False

        def communicate(self, timeout=None):
            self.attempt += 1
            if self.attempt == 1:
                raise subprocess.TimeoutExpired(["command"], timeout)
            return b"out", b"err"

        def kill(self):
            self.killed = True

    fake = FakeProcess()
    monkeypatch.setattr(processes.subprocess, "Popen", lambda *args, **kwargs: fake)
    monkeypatch.setattr(
        processes.subprocess,
        "run",
        lambda command, **kwargs: calls.append(command),
    )
    if processes.os.name != "nt":
        monkeypatch.setattr(processes.os, "killpg", lambda *args: calls.append(args))

    with pytest.raises(subprocess.TimeoutExpired):
        processes.run_process(["command"], timeout=1)

    assert fake.killed
    if processes.os.name == "nt":
        assert calls and calls[0][:2] == ["taskkill", "/PID"]
    else:
        assert calls


def test_run_process_success_path():
    result = run_process(["echo", "hello"], timeout=10)
    assert result.returncode == 0
    assert b"hello" in result.stdout


def test_run_process_nonzero_exit():
    result = run_process(["false"], timeout=10)
    assert result.returncode != 0


def test_terminate_tree_handles_process_lookup_error(monkeypatch):
    """ProcessLookupError during killpg is swallowed (process already gone)."""
    if processes.os.name == "nt":
        pytest.skip("POSIX-only path")

    def raise_plookup(*args):
        raise ProcessLookupError

    killed = []

    class FakeProcess:
        pid = 9999
        returncode = None

        def communicate(self, timeout=None):
            raise subprocess.TimeoutExpired(["cmd"], timeout)

        def kill(self):
            killed.append(True)

    fake = FakeProcess()
    monkeypatch.setattr(processes.subprocess, "Popen", lambda *a, **kw: fake)
    monkeypatch.setattr(processes.os, "killpg", raise_plookup)

    with pytest.raises(subprocess.TimeoutExpired):
        run_process(["cmd"], timeout=1)
