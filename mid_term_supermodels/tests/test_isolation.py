def test_package_is_self_contained():
    from mid_term_supermodels import pipeline, matching, multi_supermodel
    # pipeline must use the COPY's load_lead_times (with the Task-1 Description fix),
    # not the original mid_term package.
    assert pipeline.load_lead_times.__module__ == 'mid_term_supermodels.matching'
    assert hasattr(matching, 'pool_safety_stock')
    assert hasattr(matching, 'match_skus_preserve_run_vectors')
    # no symbol should resolve to the original mid_term package
    assert matching.__name__ == 'mid_term_supermodels.matching'
