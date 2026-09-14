# Excel Audit Assistant functional specification

## Objective

Build a read-only VBA prototype with AI. The prototype inspects the open
`TargetCo_historical_model.xlsx` workbook and writes candidate findings to a new report
workbook. You are specifying, reviewing and testing a controlled finance tool. You are not
being tested on VBA syntax.

## Required behaviour

The prototype must:

- inspect the open TargetCo workbook;
- distinguish declared input areas from calculation areas;
- compare formula presence across comparable year cells;
- use `FormulaR1C1` where practical to compare structural formula patterns;
- flag constants inside otherwise formula-driven rows;
- detect formula and reference errors;
- where practical, flag inconsistencies between formula role and colour metadata;
- report worksheet, cell, detected role, rule, formula or value and reason for flag;
- leave `Analyst disposition` blank;
- create findings only in a new report workbook or report sheet;
- leave every existing TargetCo business cell unchanged;
- leave reported versus adjusted EBITDA for the analyst;
- apply no correction.

The finance checks should cover EBIT, operating NWC, net debt, the named EBITDA denominator
in leverage and the historical free-cash-flow control. Growth and same-period margin checks
remain useful calibration rules.

## Safety checklist

Reject generated code if it can:

- save or save-as the TargetCo workbook;
- overwrite a TargetCo value or formula;
- delete a file or TargetCo sheet;
- close TargetCo without the user;
- call `Shell` or another system command;
- make a network or API call;
- modify an external file.

Writing to the newly created report workbook is allowed. Reading TargetCo cells is allowed.

## Desktop Excel setup

1. Keep `TargetCo_historical_model.xlsx` open.
2. Open a new blank workbook.
3. Save the new workbook as a macro-enabled `.xlsm` controller.
4. Open the VBA editor. On Windows use **Developer > Visual Basic**. On Mac use
   **Tools > Macro > Visual Basic Editor**, or **Developer > Visual Basic** when the Developer
   tab is enabled.
5. In the VBA editor choose **Insert > Module**.
6. Paste the VBA produced by ChatGPT or Claude.
7. Read the code against the safety checklist before running it.
8. Run the entry macro with TargetCo still open as the designated target workbook.
9. Confirm that the macro creates a separate report and does not edit or save TargetCo.

If the Developer tab is hidden, enable it in Excel ribbon settings. On managed computers,
follow the institution's macro-security policy. Do not weaken security settings simply to run
the exercise.

## Development workflow

1. Translate the requirements above into your own first AI coding request.
2. Ask the AI for VBA and an explanation of every procedure that writes output.
3. Inspect the code against the safety checklist.
4. Paste the reviewed code into the blank `.xlsm` controller.
5. Run it against the open TargetCo workbook.
6. Inspect the generated report against the workbook.
7. Look for false positives and false negatives under the convention used in class.
8. Improve the prompt or code, then run the revised prototype again.

The full reference build prompt appears on the prompt card. Use it after drafting
your first request, not as a substitute for specifying the task yourself.

## Test record

Record:

- one correctly flagged issue;
- one legitimate construction the tool flags, if any;
- one in-scope issue it misses, if any;
- one change you made to the prompt or code;
- the evidence that TargetCo remained unchanged.

Candidate flags require analyst attention. They do not become confirmed diagnoses until an
independent check supports the classification.
