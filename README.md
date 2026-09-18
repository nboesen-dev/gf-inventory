# GF Inventory

Internal inventory, job-planning, order-guide, and cost-estimator site for Garage Force Southeast Valley.

## Source layout

- `index.html` — GitHub Pages frontend.
- `apps-script/Code.gs` — Google Apps Script backend source. Keep this file synchronized with the deployed Apps Script project.

## Editor access

Reading inventory and using the cost estimator require no PIN. Changes to inventory or planned jobs require editor authorization.

Before deploying the backend:

1. In Apps Script, open **Project Settings**.
2. Under **Script properties**, add a property named `EDITOR_PIN`.
3. Set its value to the numeric PIN chosen by the editor.
4. Replace the Apps Script project's `Code.gs` with `apps-script/Code.gs`.
5. Create a new deployment version while retaining the existing web-app URL.
6. Deploy the matching `index.html` frontend only after the backend deployment succeeds.

The browser receives a signed editor token valid for 30 days. The PIN and signing secret are never stored in this repository.
