import pytest
import multi_supermodel


def test_run_multi_supermodel_raises_on_empty(tmp_path):
    with pytest.raises(RuntimeError, match="Nessun supermodel"):
        multi_supermodel.run_multi_supermodel(
            str(tmp_path), str(tmp_path / 'out'), 'run_config.json'
        )
