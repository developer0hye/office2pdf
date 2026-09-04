"""Minimum-cost pairing for repeated visual elements.

Text alone is not an identity: charts, legends, tables, and axes routinely
repeat the same label. Pair repeated instances by their position or geometry
feature vectors so one displaced instance cannot disappear from a report.
"""

from __future__ import annotations

import math
from collections.abc import Sequence


Vector = Sequence[float]


def minimum_cost_pairs(
    references: Sequence[Vector],
    candidates: Sequence[Vector],
    *,
    allowed_pairs: Sequence[Sequence[bool]] | None = None,
) -> list[tuple[int, int]]:
    """Return a minimum-total-distance one-to-one assignment.

    The Hungarian algorithm handles rectangular groups and remains cheap for
    pages containing many identical table values. Returned indexes always
    address ``references`` first and ``candidates`` second. When
    ``allowed_pairs`` is supplied, the result first maximises the number of
    allowed matches, then minimises their total distance; either side may keep
    unmatched items.
    """
    if not references or not candidates:
        return []

    if allowed_pairs is not None:
        if len(allowed_pairs) != len(references) or any(
            len(row) != len(candidates) for row in allowed_pairs
        ):
            raise ValueError("allowed_pairs must match the reference/candidate matrix")
        real_costs = [
            [math.dist(reference, candidate) for candidate in candidates]
            for reference in references
        ]
        feasible_costs = [
            real_costs[reference_index][candidate_index]
            for reference_index, row in enumerate(allowed_pairs)
            for candidate_index, allowed in enumerate(row)
            if allowed
        ]
        if not feasible_costs:
            return []

        # One private dummy column per reference lets the assignment leave any
        # reference unmatched. A dummy costs more than every possible change
        # across a maximum-cardinality real assignment, so cardinality wins
        # first; the Euclidean costs break ties between equally large matches.
        maximum_pair_count = min(len(references), len(candidates))
        unmatched_cost = (max(feasible_costs) + 1.0) * (maximum_pair_count + 1)
        forbidden_cost = unmatched_cost * (len(references) + 1)
        costs = [
            [
                real_costs[reference_index][candidate_index]
                if allowed
                else forbidden_cost
                for candidate_index, allowed in enumerate(row)
            ]
            + [unmatched_cost] * len(references)
            for reference_index, row in enumerate(allowed_pairs)
        ]
        pairs = _minimum_cost_pairs_from_costs(costs)
        return [
            (reference_index, candidate_index)
            for reference_index, candidate_index in pairs
            if candidate_index < len(candidates)
            and allowed_pairs[reference_index][candidate_index]
        ]

    swapped = len(references) > len(candidates)
    rows = candidates if swapped else references
    columns = references if swapped else candidates
    costs = [[math.dist(row, column) for column in columns] for row in rows]
    pairs = _minimum_cost_pairs_from_costs(costs)
    if swapped:
        pairs = [(column_index, row_index) for row_index, column_index in pairs]
    return sorted(pairs)


def _minimum_cost_pairs_from_costs(
    costs: Sequence[Sequence[float]],
) -> list[tuple[int, int]]:
    """Run rectangular Hungarian assignment for a rows-by-columns cost matrix."""

    # Rectangular Hungarian algorithm, with rows <= columns.
    row_count = len(costs)
    column_count = len(costs[0])
    if row_count > column_count:
        raise ValueError("Hungarian cost matrix must have rows <= columns")
    u = [0.0] * (row_count + 1)
    v = [0.0] * (column_count + 1)
    column_to_row = [0] * (column_count + 1)
    previous_column = [0] * (column_count + 1)

    for row_index in range(1, row_count + 1):
        column_to_row[0] = row_index
        current_column = 0
        minimum = [math.inf] * (column_count + 1)
        used = [False] * (column_count + 1)
        while True:
            used[current_column] = True
            current_row = column_to_row[current_column]
            delta = math.inf
            next_column = 0
            for column_index in range(1, column_count + 1):
                if used[column_index]:
                    continue
                reduced = (
                    costs[current_row - 1][column_index - 1]
                    - u[current_row]
                    - v[column_index]
                )
                if reduced < minimum[column_index]:
                    minimum[column_index] = reduced
                    previous_column[column_index] = current_column
                if minimum[column_index] < delta:
                    delta = minimum[column_index]
                    next_column = column_index
            for column_index in range(column_count + 1):
                if used[column_index]:
                    u[column_to_row[column_index]] += delta
                    v[column_index] -= delta
                else:
                    minimum[column_index] -= delta
            current_column = next_column
            if column_to_row[current_column] == 0:
                break
        while True:
            next_column = previous_column[current_column]
            column_to_row[current_column] = column_to_row[next_column]
            current_column = next_column
            if current_column == 0:
                break

    pairs = [
        (row_index - 1, column_index - 1)
        for column_index, row_index in enumerate(column_to_row[1:], start=1)
        if row_index
    ]
    return sorted(pairs)
