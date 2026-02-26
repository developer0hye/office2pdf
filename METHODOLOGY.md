## Universal software project methodology (concise, project-agnostic)

### 1) Clarify the intent before coding

* State the **goal**, **constraints**, **assumptions**, and **unknowns**.
* If ambiguous, document a **safe default** and/or list options with trade-offs.

### 2) Deliver the smallest useful increment

* Build only what is required to meet the goal **now**.
* Avoid speculative features and unnecessary abstraction.
* Minimize dependencies; **justify** each new dependency or custom build by overall risk and maintenance cost.

### 3) Keep changes focused and consistent

* Change only what the goal requires, plus necessary integration points.
* Defer unrelated refactors unless they block delivery.
* Follow existing conventions; if none exist, define lightweight ones and apply consistently.

### 4) Make “done” measurable and verified

* Define **success criteria** (acceptance criteria, metrics, or observable outcomes).
* Validate with automated checks when feasible; otherwise provide clear manual verification steps.
* Record what was verified and what remains unverified.

### 5) Engineer robustness proportional to risk

* Handle realistic edge cases at **Input/Output (I/O)** boundaries and external interactions.
* Treat “impossible” cases as assumptions and document them.
* Address quality attributes as needed (security, privacy, reliability, performance, compliance), scaling rigor to impact.
