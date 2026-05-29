"""Unit tests for the pure chart-data helpers added in 0.3.1 to support
the chart style guide: `_nice_bins`, `_percentile`, `_axis_scale_format`.

These are the deterministic core of the polished-histogram work — the
COM styling itself needs a live Excel, but the binning / scaling logic
that makes the X axis read in round numbers is pure Python and fully
testable here."""

from __future__ import annotations

import math

from modelrisk_mcp.bridge.reports import (
    _axis_scale_format,
    _nice_bins,
    _percentile,
)


class TestNiceBins:
    def test_round_boundaries_for_millions_range(self) -> None:
        """A 2M-8M spread should bin onto round million-ish edges so the
        thinned labels land on clean values."""
        # deterministic lognormal-ish spread roughly 2M..8M
        samples = [2_000_000 + (i % 600) * 10_000 for i in range(6000)]
        bins = _nice_bins(samples, target=30)
        assert bins.n >= 10
        # low edge is floored to a multiple of the bin width
        assert math.isclose(bins.lo % bins.bin_width, 0.0, abs_tol=1e-6)
        # bin width is a 1/2/2.5/5 x 10^n "nice" number
        mant = bins.bin_width / 10 ** math.floor(math.log10(bins.bin_width))
        assert any(math.isclose(mant, m, rel_tol=1e-6)
                   for m in (1, 2, 2.5, 5))

    def test_label_every_spaces_labels_about_one_major_unit(self) -> None:
        samples = [2_000_000 + (i % 600) * 10_000 for i in range(6000)]
        bins = _nice_bins(samples, target=30)
        # labels every `label_every` bins → between 4 and 9 labels total
        n_labels = bins.n / bins.label_every
        assert 3 <= n_labels <= 12

    def test_counts_sum_to_sample_count(self) -> None:
        samples = [float(i % 100) for i in range(1000)]
        bins = _nice_bins(samples)
        assert sum(bins.counts) == len(samples)

    def test_cumulative_is_monotonic_and_ends_at_one(self) -> None:
        samples = [float(i % 100) for i in range(1000)]
        bins = _nice_bins(samples)
        assert bins.cumulative == sorted(bins.cumulative)
        assert math.isclose(bins.cumulative[-1], 1.0, abs_tol=1e-9)

    def test_centres_count_matches_n(self) -> None:
        samples = [float(i % 100) for i in range(1000)]
        bins = _nice_bins(samples)
        assert len(bins.centres) == bins.n
        assert len(bins.counts) == bins.n
        assert len(bins.cumulative) == bins.n

    def test_empty_samples_safe(self) -> None:
        bins = _nice_bins([])
        assert bins.n == 0
        assert bins.centres == []

    def test_degenerate_all_equal_safe(self) -> None:
        """All-identical samples must not divide by zero."""
        bins = _nice_bins([5.0] * 100)
        assert bins.n >= 1
        assert sum(bins.counts) == 100


class TestPercentile:
    def test_median_of_known_set(self) -> None:
        assert _percentile([1, 2, 3, 4, 5], 0.50) == 3

    def test_p10_p90(self) -> None:
        xs = list(range(1, 101))  # 1..100
        assert math.isclose(_percentile(xs, 0.10), 10.9, abs_tol=1e-9)
        assert math.isclose(_percentile(xs, 0.90), 90.1, abs_tol=1e-9)

    def test_min_max_edges(self) -> None:
        xs = [10, 20, 30]
        assert _percentile(xs, 0.0) == 10
        assert _percentile(xs, 1.0) == 30

    def test_empty_safe(self) -> None:
        assert _percentile([], 0.5) == 0.0

    def test_unsorted_input(self) -> None:
        assert _percentile([5, 1, 3, 2, 4], 0.50) == 3


class TestAxisScaleFormat:
    def test_millions(self) -> None:
        assert _axis_scale_format([4_200_000, 1_000_000]) == '#,##0,,"M"'

    def test_thousands(self) -> None:
        assert _axis_scale_format([8500, 42000]) == '#,##0,"K"'

    def test_units(self) -> None:
        assert _axis_scale_format([12, 420, 99]) == '#,##0'

    def test_uses_max_magnitude(self) -> None:
        # one big value pushes the whole axis to the millions format
        assert _axis_scale_format([5, 9_000_000]) == '#,##0,,"M"'

    def test_negative_magnitude(self) -> None:
        assert _axis_scale_format([-4_200_000, 10]) == '#,##0,,"M"'

    def test_empty_safe(self) -> None:
        assert _axis_scale_format([]) == '#,##0'
