from macro import Macro
import pytest


@pytest.mark.parametrize(
    "string, expected_regex",
    [
        ("48 Stk. à 19,5 m, 6", ("48", "19,5", "6")),
        ("Lfr 12,5m, Üli 1m, Üre 0m", ("12,5", "1", "0")),
        ("1 1,5 0,5", ("1", "1,5", "0,5")),
    ],
)
def test_find_all_lenghts_regex(string, expected_regex):
    macro = Macro()
    assert macro.find_all_lengths(string) == expected_regex
