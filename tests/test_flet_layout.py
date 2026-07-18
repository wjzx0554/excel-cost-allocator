import flet as ft

from flet_app import FletAllocatorApp


class FakeWindow:
    pass


class FakePage:
    def __init__(self):
        self.overlay = []
        self.window = FakeWindow()
        self.controls = []

    def add(self, *controls):
        self.controls.extend(controls)

    def update(self):
        pass


def _walk(control):
    yield control
    for child in control._get_children():
        yield from _walk(child)


def test_wrapping_rows_do_not_contain_expanded_children():
    page = FakePage()
    FletAllocatorApp(page)

    invalid_rows = [
        control
        for root in page.controls
        for control in _walk(root)
        if isinstance(control, ft.Row)
        and control.wrap
        and any(child.expand for child in control.controls)
    ]

    assert invalid_rows == []


def test_main_views_use_flex_layout_instead_of_stack():
    page = FakePage()
    app = FletAllocatorApp(page)

    assert isinstance(app.views, ft.Column)
    assert app.views.expand
    assert [view.visible for view in app.views.controls] == [True, False, False, False]
