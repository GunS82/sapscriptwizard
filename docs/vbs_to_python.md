# Translating VBS calls to Python

The low-level API mirrors the original VBS syntax. Replace calls like::

    session.findById("wnd[0]/usr/txtSOME").text = "value"

with::

    window.write("wnd[0]/usr/txtSOME", "value")

Use :class:`~sapscriptwizard.core.window.Window` for imperative interactions and
higher level recipes in :mod:`sapscriptwizard.features` for complex flows.
