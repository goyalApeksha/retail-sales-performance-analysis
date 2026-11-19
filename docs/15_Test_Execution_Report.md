# Step 15 Test Execution Report

**Project:** Retail Sales Performance Analysis  
**Date:** 19 November 2026  
**Build tasks:** BUILD-01–BUILD-14

## Automated test run

- Command: `python -m pytest -q`
- Result: **29 passed in 21.92 seconds**
- AppTest exercised the app entry point and all seven multipage screens.
- Runtime used for this completed run: system Python 3.13.5 with the installed compatibility packages (Streamlit 1.51.0, pandas 2.3.1, NumPy 2.3.1, SciPy 1.16.0); pytest 9.2.0 was loaded from the project virtual environment.
- The run is a functional verification of the current source and tests, but it is not evidence that the pinned Python 3.12 environment is fully installed.

## Build and runtime notes

- Pinned requirements are recorded in `requirements.txt`. The Python 3.12 virtual environment was created with Python 3.12.14. Installing all pinned packages did not complete during this run because large package downloads were still progressing very slowly.
- A direct `streamlit run app.py` launch was attempted; the local HTTP probe was blocked by the execution environment's local socket restrictions. AppTest startup and multipage rendering passed.
- No manual browser execution of the 24 functional scenarios was recorded. Those cases remain for learner walkthrough and evidence capture.
- The Step 1 target of at least 1,000 fact rows conflicts with the Step 8 plan and Step 10 synthetic baseline of 240 rows. This build retains the later approved seed-42 baseline of 240 rows; update the generator, checks, validation report, and documents together if the owner changes the decision.
- Deployment and GitHub remote setup remain for Step 16. No remote is configured in this workspace.

## Defects fixed

- Added an HTML rendering compatibility fallback for Streamlit versions without `st.iframe`.
- Corrected the test fixture to provide the required payment method field. The production validation rule was retained.

## Release status

Automated tests pass in the noted compatibility environment. The exact pinned environment, manual functional walkthrough, and deployment smoke test remain outstanding.
