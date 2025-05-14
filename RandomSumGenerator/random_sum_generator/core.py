import random
import logging

logger = logging.getLogger(__name__)

class RandomSumGenerator:
    def __init__(self, seed=None, debug=False):
        self.debug = debug
        if seed is not None:
            random.seed(seed)
            if self.debug:
                logger.debug(f"Seed set to: {seed}")

    def _normalize_bounds(self, val, parts, default):
        if isinstance(val, list):
            if len(val) != parts:
                raise ValueError(f"List length of bounds must match parts. Got {len(val)} for {parts} parts.")
            return val
        else:
            return [val] * parts

    def generate(self, total, parts, min_val=0, max_val=None, mode='float', precision=2, max_attempts=1000):
        if max_val is None:
            max_val = total

        min_vals = self._normalize_bounds(min_val, parts, 0)
        max_vals = self._normalize_bounds(max_val, parts, total)

        # 🧠 Feasibility Check
        if sum(min_vals) > total:
            raise ValueError("Sum of min_vals is too high for the given total.")
        if sum(max_vals) < total:
            raise ValueError("Sum of max_vals is too low to reach the total.")

        avg = total / parts
        if all(isinstance(x, (int, float)) for x in max_vals):
            tight_upper = all(round(m, 5) <= round(avg, 5) for m in max_vals)
            if tight_upper:
                raise ValueError(
                    f"max_val too tight to generate varied values. Try setting max_val > total/parts = {avg:.2f}"
                )

        for attempt in range(max_attempts):
            raw = [random.gammavariate(1, 1) for _ in range(parts)]
            total_raw = sum(raw)
            proportions = [r / total_raw for r in raw]
            scaled = [min_vals[i] + proportions[i] * (max_vals[i] - min_vals[i]) for i in range(parts)]

            if mode == 'int':
                scaled_total = sum(scaled)
                adjusted = [int(round(s * total / scaled_total)) for s in scaled]
                diff = total - sum(adjusted)
                adjusted[-1] += diff
                if all(min_vals[i] <= adjusted[i] <= max_vals[i] for i in range(parts)):
                    if self.debug:
                        logger.debug(f"[Attempt {attempt}] Result: {adjusted}")
                    return adjusted

            elif mode == 'float':
                scaled_total = sum(scaled)
                adjusted = [round(s * total / scaled_total, precision) for s in scaled]
                diff = round(total - sum(adjusted), precision)
                adjusted[-1] = round(adjusted[-1] + diff, precision)
                if all(min_vals[i] <= adjusted[i] <= max_vals[i] for i in range(parts)):
                    if self.debug:
                        logger.debug(f"[Attempt {attempt}] Result: {adjusted}")
                    return adjusted

            else:
                raise ValueError("Mode must be 'int' or 'float'")

        raise RuntimeError(f"Failed to generate valid output in {max_attempts} attempts.")