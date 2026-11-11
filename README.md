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
- `example.html` is an illustrated two-agent, one-task example, worked through by hand with round numbers. It ends in a button that opens the calculator with the same inputs, so the results can be compared line by line: 1,200 calls, $6,000 per agent, 100 paid hours, 20% shrinkage, 75% occupancy, 6-minute calls, AVA at $2 per call, 50% target. That gives 1 of 2 agents freed, $3,600 saved per month and break-even at 20%. It is linked from the calculator's header.
- **Inputs tab**: every input from the customer data-gathering form in one place, grouped the same way (customer, call volume, call flow, agent cost and time, per-call fees, fixed costs, rollout and goals, notes), with break-even and savings shown at the top. It shares its values with the **Calculator** tab, so a change in either tab shows in both. Customer details and notes are saved with scenarios, exports and links, and the company name appears on the executive report.
- **Scenarios** (bar under the title): **New** starts an untitled scenario with every input back at its default, including customer details and notes, and warns first if the current inputs aren't saved. **Save** stores the current inputs under a name, such as one per customer or version. Pick a name from the list to recall it. **Save as**, **Rename** and **Delete** manage the list, and "Unsaved changes" shows when the inputs differ from the saved copy. Scenarios are kept in your browser only. **Export** downloads them all as a JSON file and **Import** loads such a file, to move them to another computer or keep a backup. **Copy link** copies a URL with the current inputs in it (after the `#`, so they are not sent to any server). Anyone who opens it sees those inputs and can save them. The executive report shows the scenario name.
- The **Executive report** button opens a two-page report of the current scenario. Page 1 is a short value statement (multi-year savings headline, annual savings, FTEs freed, payback, and a bar showing break-even against the target). Page 2 holds the supporting details (break-even, year-by-year savings, savings build-up, call paths, assumptions and method). Use **Print / save as PDF** to export both pages on US Letter.
- `ava-break-even-check.xlsx` recomputes everything with visible spreadsheet formulas and runs sanity checks, so results can be validated independently. Like the page, it supports either input: set the volume input mode to 1 to enter calls per month, or 2 to enter agent FTEs today. Regenerate it with `pip install openpyxl && python tools/build_workbook.py` after changing the model.
- `customer-data-gathering.docx` is a customer-facing Word form that collects every calculator input (call volume or agents, call paths, agent cost and time, fees, fixed costs, rollout and goals), with a description, where to find it and an example for each. Regenerate it with `npm install docx && node tools/build_intake_form.js`.
- `.cursor/rules/ava-break-even.mdc` gives Cursor the model, formulas and conventions for this project.

Values on the page are examples until replaced with real figures.

## Benchmark sources for default inputs

Defaults are examples. Replace them with the customer's figures when available.

- **Shrinkage, default 30%.** [Call Centre Helper](https://www.callcentrehelper.com/industry-standards-metrics-125584.htm): "The Call Centre Industry Standard Metric for Shrinkage is 30%". Its [shrinkage guide](https://www.callcentrehelper.com/how-to-calculate-contact-centre-shrinkage-90353.htm) says 30–35%. [NiCE](https://www.nice.com/glossary/shrinkage): typically 25–40%, with averages around 30–35%. [Five9](https://www.five9.com/blog/understanding-and-managing-shrinkage-contact-centers): industry average of about 30–35%.
- **Occupancy, default 80%.** [SQM Group](https://www.sqmgroup.com/resources/library/blog/industry-standards-top-call-center-kpis): the industry standard is 75–85%, and above 85% is not sustainable for long periods. [Call Centre Helper](https://www.callcentrehelper.com/industry-standards-metrics-125584.htm): 83.3% is its standard for maximum occupancy. [ICMI](https://www.icmi.com/resources/2023/agent-efficiency-between-calls): agents burn out when occupancy is above the low 90s for long periods. [TechTarget](https://www.techtarget.com/enterprise-software/answer/How-to-calculate-call-center-utilization-rates): burnout is likely when occupancy is consistently above 88–92%.

