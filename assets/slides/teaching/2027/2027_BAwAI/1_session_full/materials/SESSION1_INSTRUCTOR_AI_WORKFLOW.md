# Session 1 Instructor AI Workflow — TargetCo

## Purpose and boundary

This source-constrained workflow is demonstrated after redo. It uses only the Session-1-visible TargetCo profile. It introduces no external facts, valuation, hidden TargetCo data or proprietary source content. The baseline and redo occur in different fresh ephemeral chats; no prompt asks a model to forget a prior conversation.

## Operational policy

- **ChatGPT:** start a new non-personalized Temporary Chat where available.
- **Claude:** start a new Incognito chat.
- **Redo budget:** at most three user messages: evidence/analysis; challenge+draft; verify+revise.
- **Pre-class check:** test both product modes before class because product UIs and availability can change.
- **Contingency:** use the Markdown/text profile if uploads create a limit problem. ChatGPT/Claude limits are product-plan constraints that can change; the four-message course budget is not a vendor entitlement claim.

Official locators: OpenAI Temporary Chat FAQ, Anthropic Incognito/privacy documentation, Anthropic usage-limit guidance and OpenAI file-upload FAQ are recorded in `SESSION1_SOURCE_NOTES.md`.

## Authoritative input and success criteria

- **Authoritative material:** Student TargetCo profile PDF or semantically equivalent Markdown version.
- **Task:** Student VP request.
- **Four criteria:** business-model synthesis; historical economics; evidence and uncertainty discipline; prioritization / VP usefulness.
- **Independent check:** source table plus manual recomputation, including `481.0 / 454.0 - 1 = 5.9%`.

## Strong near-final prompt sequence

### Message 1 — evidence / analysis

~~~text
The supplied TargetCo profile is the authoritative source for this exercise. Use no external facts, tools or valuation.

Extract: product specification/qualification, end markets, channels and footprint; revenue and g_t = R_t / R_(t-1) - 1; disclosed volume and price/mix contribution points, confirming g_t = v_t + p_t at displayed precision; utilization and reported EBITDA margin; and M-01/M-02 as management statements.

Return source location, classification, evidence/calculation, support and limit. Flag the utilization/margin tension: 2025A utilization is 84% versus 82% in 2022A, while margin is 14.3% versus 14.8%. Do not draft.
~~~

**Human review:** confirm 2023A has -4.0 ppt volume plus +6.1 ppt price/mix for 2.1% displayed revenue growth; confirm the utilization/margin comparison remains a tension, not causality.

### Message 2 — challenge + draft

~~~text
Challenge against four criteria: coherent business-model synthesis; correct contribution arithmetic and qualified causality; evidence/uncertainty discipline; concise VP usefulness with exactly three ranked questions naming a later judgement.

Correct only what the profile supports. Draft one page with exactly: Business model / positioning; Historical trajectory; Key tensions; Three ranked questions + why. Preserve M-01/M-02 as management statements. No valuation.
~~~

**Human review:** reject invented pricing power, retention, contract durability, concentration classification, causal margin bridge or automatic normalization.

### Message 3 — verify + revise

~~~text
Reconcile each material number and statement to a profile location or calculation. Recheck 2023A (-4.0 ppt + 6.1 ppt = 2.1%) and 2025A (4.0 ppt + 1.9 ppt = 5.9%). Confirm no management statement became fact, causal language remains qualified and exactly three questions remain ranked.

Return a short reconciliation list, then revise only where the check requires it. Flag unresolved uncertainty.
~~~

**Human review:** independently recompute material arithmetic and inspect the source. An AI ``all checks passed'' response is not independent verification.

## Analyst retains judgement

The analyst determines source authority, whether causal language is justified, which tension matters for valuation, whether a cost qualifies for normalization and whether the final answer is acceptable.

## Work Record fields

Capture actual baseline prompt; fresh-chat mode; redo prompts/labels; delegation versus retained judgement; challenged statement; numerical/source check; remaining uncertainty; and a short baseline-to-redo reflection.
