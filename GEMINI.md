# UnoGenerator Development Notes

## Architectural Decisions & Bug Fixes

### 1. Row Height and Column Width Interactions (Calc)
**Problem:** In LibreOffice Calc, setting `OptimalHeight = True` on an empty cell or range results in a default height (e.g., 452 for 10pt font). If data is inserted afterwards, or if column widths are changed, Calc may fail to automatically shrink the row height back to its minimum single-line value, or it may jump to a larger height (e.g., 841) if it considers the previous calculation "locked".

**Decisions:**
- **Data First, Height Second:** In `addListOfRowsWithStyle`, we now insert all data (`setDataArray`/`setFormulaArray`) *before* calling `_set_rows_optimal_height`. This ensures Calc has the actual content to perform a correct single-pass height calculation.
- **Forced Refresh on Width Change:** In `setColumnsWidth`, we explicitly toggle `OptimalHeight` (False then True) for all rows that are supposed to be wrapped. This forces Calc to recalculate heights based on the *new* column widths, preventing rows from staying at an excessively large height when they could now fit in a single line.
- **Block Processing Performance:** To maintain high performance with large datasets, the refresh logic in `setColumnsWidth` only iterates over rows known to have wrapping enabled (`_wrapped_rows`) and processes them in contiguous blocks to minimize UNO API calls.

### 2. Style Application Optimization
**Problem:** An optimization in `addListOfRowsWithStyle` incorrectly skipped applying styles if they matched `self.default_cell_style`. This was problematic for `ODS_Standard` which uses `Normal`, because the underlying template often defaults to `Default`. If we didn't explicitly set `Normal`, the cells would remain as `Default`, leading to styling inconsistencies.

**Decision:**
- **Safety-First Optimization:** The optimization now only skips style application if both the intended style is `Default` AND the document's default is also `Default`. If a custom `default_cell_style` (like `Normal`) is defined, it will always be explicitly applied to ensure document consistency.

### 5. Localization Workflow (`poe translate`)
**Observation:** The project uses `poe translate` (linked to `unogenerator.poethepoet:translate`) to automate the synchronization of `.po` files with the `.pot` template.
**Experience:** 
- **English Redundancy:** `en.po` has been removed as it was redundant. Since the source strings in the code are in English, the system now uses `fallback=True` in `gettext.translation` calls within `demo.py`. This ensures that when the English locale is requested, it falls back to the original English strings without needing a separate `.po` file.
- **Header Fragility:** Running `poe translate` may overwrite manual header improvements or reset them to defaults.
- **Multilingual Support:** While efficient for syncing Spanish (the developer's primary language), it might leave other languages (French, Romanian) with empty `msgstr` entries if not carefully monitored.
- **Verification:** After running `poe translate`, always verify that all message strings in all supported languages are correctly populated and that headers maintain the correct project metadata and dates. Manual restoration of translations for non-primary languages may be required if the tool only targets one language.

### 6. Regression Testing
New tests have been added to `tests/test_unogenerator.py` to protect these fixes:
- `test_ods_row_height_consistency`: Ensures heights stay at 452 even after `setColumnsWidth` with many columns.
- `test_ods_normal_style_applied`: Verifies that `Normal` style is correctly applied by `ODS_Standard`.

### 4. Dependency: `envwrap`
**Issue:** When using `uno` (LibreOffice Python API), it modifies the Python import hook. In certain environments, this causes `tqdm` (a project dependency) to fail if `envwrap` is not explicitly installed, resulting in an `ImportError`.

**Decision:**
- **Explicit Dependency:** `envwrap` has been added to `pyproject.toml`. While not a direct dependency of the library's core logic, it is essential for the environment's stability when `tqdm` and `uno` coexist.
- **Safety:** It is a safe, lightweight utility for environment variable wrapping. Adding it explicitly prevents the intermittent `ImportError` and ensures that tests and demo scripts run reliably across different setups.

### 7. PyUNO Initialization in Concurrent Processes
**Problem:** When using `multiprocessing` with the `spawn` start method (as used in the demo), child processes would fail with `SystemError: pyuno runtime is not initialized`. This is because `uno.getComponentContext()` (from the standard `uno.py`) returns a cached context object that is initialized during the module's import phase. In `spawn` mode, this initialization happens during the child process's bootstrap and is not valid for the subsequent execution phase.

**Decision:**
- **Direct PyUNO Access:** We now use `pyuno.getComponentContext()` directly instead of the cached `uno.getComponentContext()`.
- **Deferred Initialization:** In `unogenerator.py`, we defined a local `getComponentContext()` wrapper that calls the underlying `pyuno` function. This ensures that every time the context is requested (e.g., during `ODF.__init__` or `createUnoService`), `pyuno` performs a fresh initialization check in the current process phase.
- **Verification:** A new test `tests/test_concurrency.py` has been added to specifically verify that `spawn`-ed processes can successfully initialize and use UNO objects.

### 8. Optimized Thread-Safety via Instance-Based Serialization
**Problem:** Running the demo with concurrent threads and a shared LibreOffice server (`COMMONSERVER_CONCURRENT_THREADS`) was unstable due to non-thread-safe sharing of a single connection. However, using a single global lock serialized execution even for workers using independent LibreOffice instances, causing unnecessary performance loss.

**Decision:**
- **Instance-Based Reentrant Lock:** Each `LibreofficeServer` object now maintains its own `threading.RLock()`.
- **Smart Synchronization:** The `@uno_safe` class decorator dynamically retrieves the lock from the document's associated server instance (`self.server._lock`).
- **Parallel Performance:** This allows workers using independent servers (different ports) to run at full parallel speed without waiting for each other.
- **Shared-Server Stability:** Workers sharing the same `LibreofficeServer` instance will correctly share the same lock, ensuring serialized access to the single connection and maintaining stability.
- **Internal Protection:** Global UNO bridge calls (like `getComponentContext`) remain protected by a global `_uno_bridge_lock` to ensure the process-wide PyUNO state is not corrupted during initialization.

### 9. Test Execution Scope (Performance Optimization)
**Decision:**
- **Targeted Testing for AI Assistant:** The AI assistant must only run targeted test files or specific test functions relevant to the modified code (e.g., `poetry run pytest tests/test_helpers.py`), and NOT run the complete global `pytest` suite. The full global test suite is executed exclusively by the user to optimize task execution times.
