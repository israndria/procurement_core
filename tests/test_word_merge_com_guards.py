import word_merge


class _NoActiveWindow:
    @property
    def ActiveWindow(self):
        raise RuntimeError("This command is not available because no document is open.")


def test_word_process_id_ignores_missing_active_window():
    assert word_merge._word_process_id(_NoActiveWindow(), _NoActiveWindow()) is None


class _Sections:
    Count = 2


class _NoPageNavigationDocument:
    Sections = _Sections()
    go_to_called = False

    def ComputeStatistics(self, _statistic):
        return 3

    def GoTo(self, *_args):
        self.go_to_called = True
        raise AssertionError("BA export must not navigate pages through Word COM")


def test_ba_header_cleanup_does_not_navigate_pages():
    document = _NoPageNavigationDocument()
    assert word_merge._strip_pl_ba_signature_header(document) is False
    assert document.go_to_called is False
