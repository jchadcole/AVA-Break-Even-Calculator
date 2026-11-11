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

- Handle hours per FTE = paid hours per month × (1 − shrinkage) × occupancy. Shrinkage is breaks, training, meetings and time off. Time agents spend on other tasks (email, back office) is deliberately left out: automating calls frees only the call hours, and the other work remains
- % of working time on these calls is used only in headcount mode: calls = agent FTEs × that % × calls per FTE ÷ (direct % + % reaching AVA). Every FTE figure in the results is FTEs of call work
- Agent hours automated = AVA-contained calls × average call time ÷ 3600
- FTEs saved = agent hours automated ÷ handle hours per FTE
- Labor saved = FTEs saved × fully loaded agent cost
- IVR fees avoided = calls reaching AVA × IVR cost per call
- Net savings = labor saved + IVR fees avoided − AVA per-call cost − AVA fixed costs
- Agent-hour equivalents: dollar amounts ÷ agent cost per handle hour (cost ÷ handle hours per FTE). The page shows hours first with dollars beside, except the year-by-year chart, which is in dollars
- Year by year: containment starts at the target and rises by a configurable number of points each year (capped at the share of calls reaching AVA); each year's savings = months in production at that containment. Year 1 leaves out a configurable implementation time (default 3 months) with no savings or AVA per-call fees, and the implementation cost is charged in year 1

Each call is billed once for automation: IVR-contained calls pay the IVR fee; calls that reach AVA pay the AVA fee instead (not IVR + AVA), plus agent time if they escalate.

The comparison point is the same flow with no AVA, where every call that leaves the IVR goes to an agent. Break-even is the lowest AVA containment where monthly cost is at or below that figure.

## Files

- `index.html` is the calculator. It is a single self-contained page; inputs are saved in the viewer's browser.
- The **Executive report** button opens a two-page report of the current scenario. Page 1 is a short value statement (multi-year savings headline, annual savings, FTEs freed, payback, and a bar showing break-even against the target). Page 2 holds the supporting details (break-even, year-by-year savings, savings build-up, call paths, assumptions and method). Use **Print / save as PDF** to export both pages on US Letter.
- `ava-break-even-check.xlsx` recomputes everything with visible spreadsheet formulas and runs sanity checks, so results can be validated independently. Like the page, it supports either input: set the volume input mode to 1 to enter calls per month, or 2 to enter agent FTEs today. Regenerate it with `pip install openpyxl && python tools/build_workbook.py` after changing the model.
- `customer-data-gathering.docx` is a customer-facing Word form that collects every calculator input (call volume or agents, call paths, agent cost and time, fees, fixed costs, rollout and goals), with a description, where to find it and an example for each. Regenerate it with `npm install docx && node tools/build_intake_form.js`.
- `.cursor/rules/ava-break-even.mdc` gives Cursor the model, formulas and conventions for this project.

Values on the page are examples until replaced with real figures.
