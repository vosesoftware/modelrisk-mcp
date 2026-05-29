# Risk-analysis knowledge base

A curated distillation of risk-modelling guidance from the **ModelRisk
Help** (Vose Software). It is written for the LLM to consult while
building or critiquing a model: principle-dense, decision-oriented,
and grounded in the authoritative source rather than general intuition.

This is an *attributed distillation in our own words* — each section
cites the ModelRisk Help article it draws from. For the full treatment,
consult that article in the ModelRisk Help. Where this guidance and a
specific client's situation conflict, the source article and the
analyst's judgement win.

---

## 1. Why quantify uncertainty at all

Risk analysis is the systematic identification and assessment of the
uncertainties that bear on a goal, followed by finding a strategy to
control them efficiently.

The reason it changes decisions: two options with the *same* expected
outcome are not equally good if one is more uncertain than the other.
Consider two vaccines with identical reported efficacy (60%), but one
twice as uncertain about that figure as the other — all else equal you
buy the less uncertain one. Swap "vaccine/efficacy" for
"investment/profit" and it's the same decision. A single point estimate
throws away exactly the information — the spread — that distinguishes
the choices. Determining that spread is the central job of risk
analysis.

*Source: ModelRisk Help — Introduction to risk analysis.*

---

## 2. The two forms of uncertainty

A risk model deals with uncertainty in two distinct shapes, and they
are modelled differently:

1. **Variability / general uncertainty** — a quantity that has a range
   of possible values (demand, cost, duration). Modelled as a
   **distribution**.
2. **Risk events** — discrete events that may or may not occur, each
   with an impact if it does (a breach, a recall, a failure). Modelled
   with `VoseRiskEvent`, *not* by multiplying probability by impact —
   the multiplication erases the event's all-or-nothing character and
   misstates the tail.

Keeping these two straight is foundational: don't smear a discrete
event into a continuous distribution, and don't pretend a continuous
quantity only takes one value.

*Source: ModelRisk Help — Introduction to risk analysis; Vose Risk
Event.*

---

## 3. Selecting the appropriate distribution

> "Inappropriate use of probability distributions has proven to be a
> very common failure of risk analysis models." — ModelRisk Help

Distribution choice is where precision is won or lost. Select on five
properties, in roughly this order:

1. **Discrete or continuous.** The most basic and most-overlooked
   distinction. A discrete quantity takes identifiable separate values
   (number of bridges, personnel, customers) — you cannot have half a
   bridge or 2.7 people. A continuous quantity is infinitely divisible
   (time, mass, distance, and in practice cost and exchange rate,
   whose steps are insignificant). Match the distribution's nature to
   the variable's.
2. **Bounded or unbounded.** A bounded distribution lies between two
   fixed limits; an unbounded one runs to ±∞; a partially-bounded one
   is constrained at one end. The classic failure: using a Normal
   (unbounded) for a quantity that can't go negative (sales volume,
   cost), which leaks probability below zero. If a meaningful share of
   draws would be nonsensical, constrain the distribution — e.g.
   `VoseNormal(10, 3, VoseXBounds(5, 17))` truncates to [5, 17], or
   `VoseXBounds(, 15)` cuts only the right tail. (No ModelRisk
   distribution is bounded only on the right; to get one, invert a
   left-bounded distribution — e.g. `=10 - VoseGamma(2, 1.5)`.)
3. **Parametric or non-parametric.** A *parametric* (model-based)
   distribution gets its shape from an underlying probability model
   (Poisson, Lognormal, Weibull…); a *non-parametric* (empirical) one
   gets its shape from parameters that are direct features of the graph
   (Triangle's min/mode/max, a Histogram, a Cumulative). Use a
   parametric distribution **only when** the theory behind it actually
   applies to the problem, or it has an established track record for
   that quantity. Otherwise prefer the intuitive, flexible
   non-parametric forms — they make fewer hidden assumptions.
4. **Univariate or multivariate.** Model one quantity, or a set of
   quantities whose joint behaviour matters together (see §5).
5. **First- or second-order.** First-order captures variability;
   second-order separates variability from *uncertainty about the
   parameters themselves* — the same idea as fitting with
   `uncertainty=TRUE` (carry parameter uncertainty through, don't
   pretend the fitted parameters are exact).

*Source: ModelRisk Help — Selecting the appropriate distributions for
your model.*

---

## 4. Eliciting expert opinion

Almost every real model needs subjective estimates: the data were never
collected, are too costly, are no longer relevant, or are too sparse.
Expert opinion is then a legitimate and often the only source — but it
must be elicited carefully to limit bias. Broad principles:

- **Choose the expert for knowledge and lack of bias.** If possible,
  involve them in the original model design.
- **Give them the information first.** Collate what's known and present
  it well so the expert is oriented before estimating.
- **Explain why the estimate is needed.** It improves cooperation and
  surfaces factors the expert knows about (correlations, caveats).
- **Brainstorm together, estimate apart.** With several experts,
  restrict the group to *discussing information*; take each estimate
  privately. Consistent private estimates signal the information was
  well understood.
- **Let the expert describe the uncertainty in their own terms, then
  match the model to it.** Don't force every opinion into a bare
  min/most-likely/max. **Disaggregation** — breaking a hard-to-estimate
  quantity into parts that are each easier to judge — is particularly
  effective. Use the full range of subjective distributions, not just
  PERT/Triangle, when the expert's described shape calls for it.

*Source: ModelRisk Help — Modeling expert opinion.*

---

## 5. Modeling correlation

Independent variables take values without regard to each other;
**dependent** variables don't — the probable value of one is tied to
the value of another (design time and coding time on a software
project; price and demand). Ignoring real dependence produces
artificially tight outputs because the model assumes the bad cases
won't coincide.

ModelRisk supports several ways to model dependence, in rough order of
sophistication:

- **Rank-order correlation** — quick, available in most MC add-ins,
  forces a correlation through the random-number generation. Easy but
  non-intuitive, and it doesn't model the *direction* of influence.
- **Envelope method** — the dependent variable's distribution
  parameters are functions of the independent variable. Intuitive,
  easy to check, good for expert-opinion correlations and one-to-many
  relationships; awkward for many-to-many.
- **Lookup tables** — switch or modify the dependent distribution based
  on the driving variable's value. Also good for expert-opinion
  one-to-many relationships.
- **Conditional logic** — `IF`/`AND`/`OR` to switch values by other
  cells' values; full control, more manual.
- **Copulas** — the most general and robust: a copula binds the
  marginals into a joint distribution while preserving each marginal's
  shape, and lets you choose the *dependence structure* (Normal/t for
  symmetric association; Clayton/Gumbel for tail dependence). Preferred
  for serious correlation modelling. Verify the realised correlation
  after the run.

Pick the lightest method that captures the dependence that actually
matters; reach for copulas when tail co-movement or many-to-many
structure is in play.

*Source: ModelRisk Help — Modeling correlation; Copulas.*

---

## 6. Modeling a quantity over time

When you care about a variable's path across periods — not just one
end-point (share price, exchange rate, import volume, outbreak counts,
consumption) — use a **time-series** model, not independent per-period
draws. Independent draws throw away the period-to-period relationship
that defines such a series.

A sound time-series model must reflect:

- **The relationship between the value at each period** (memory /
  autocorrelation).
- **Realistic ranges of the variable over time.**
- **Trend (drift), seasonality, and cyclicity** where present.
- **How uncertainty itself behaves over time** — typically widening
  the further out you project.

The first fork in choosing a family is **does the quantity revert to a
level, or wander freely?** Reverting series (many commodities, interest
rates) call for mean-reverting processes; freely-drifting positive
series (prices, FX) call for geometric Brownian motion. Getting this
fork wrong produces paths that either fan out far too wide or never
explore the realistic range.

*Source: ModelRisk Help — Time series introduction.*

---

## How to use this knowledge base

- When **building**: work down §3's five properties for every input;
  apply §4 when the input is an expert judgement; apply §5 when inputs
  co-move; apply §6 when a quantity has a time path. Keep §2's two-forms
  distinction front of mind.
- When **auditing**: these principles are the *why* behind the audit
  rules — see `modelrisk://methodology` for the principle-to-rule map.
- For the **mechanics** (which exact function, how to compose
  structures): see `modelrisk://distributions` and the project's
  modeling-patterns and distribution-selection guides.

*All sections distilled from the ModelRisk Help (Vose Software). Consult
the corresponding Help article for the authoritative, complete
treatment.*
