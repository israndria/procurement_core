import word_merge


class _NoActiveWindow:
    @property
    def ActiveWindow(self):
        raise RuntimeError("This command is not available because no document is open.")


def test_word_process_id_ignores_missing_active_window():
    assert word_merge._word_process_id(_NoActiveWindow(), _NoActiveWindow()) is None
