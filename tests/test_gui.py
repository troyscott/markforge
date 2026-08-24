from __future__ import annotations

from markforge.gui import MarkForgeApp


def test_gui_does_not_override_tkinter_internal_options_method() -> None:
    assert "_options" not in MarkForgeApp.__dict__
    assert "_conversion_options" in MarkForgeApp.__dict__
