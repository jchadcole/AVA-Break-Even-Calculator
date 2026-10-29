# AVA Break-Even Calculator

Interactive calculator for the share of inbound calls that AVA (the AI virtual agent) must contain before agent savings cover what AVA costs, plus the savings as containment grows.

**Live site:** https://jchadcole.github.io/AVA-Break-Even-Calculator/

## Call flow

Every inbound call takes one of four paths:

1. Direct to agent
2. Contained in IVR
3. IVR → contained in AVA
4. IVR → AVA → agent

## How savings are calculated

- Handle hours per FTE = paid hours per month × % of agent time on these calls × occupancy
- Agent hours automated = AVA-contained calls × average call time ÷ 3600
- FTEs saved = agent hours automated ÷ handle hours per FTE
- Labor saved = FTEs saved × fully loaded agent cost
- Net savings = labor saved − AVA per-call cost − AVA fixed costs

The comparison point is the same flow with no AVA, where every call that leaves the IVR goes to an agent. Break-even is the lowest AVA containment where monthly cost is at or below that figure.

## Files

- `index.html` is the calculator. It is a single self-contained page; inputs are saved in the viewer's browser.
- `ava-break-even-check.xlsx` recomputes everything with visible spreadsheet formulas and runs sanity checks, so results can be validated independently. Regenerate it with `pip install openpyxl && python tools/build_workbook.py` after changing the model.
- `.cursor/rules/ava-break-even.mdc` gives Cursor the model, formulas and conventions for this project.

Values on the page are examples until replaced with real figures.
