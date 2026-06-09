---
name: "deliver-prd-to-plan-tasks"
description: "Transforms an approved PRD into a practical delivery plan and implementation-ready task breakdown. Use after requirements are stable enough for planning, when teams need sequencing, dependencies, estimates, owners, validation steps, and backlog-ready tasks grounded in the original product requirements."
version: 1.0.0
updated: 2026-06-08
license: Apache-2.0
phase: deliver
category: coordination
---

<!-- PM-Skills | https://github.com/product-on-purpose/pm-skills | Apache 2.0 -->
# PRD to Plan and Tasks

This skill converts an existing Product Requirements Document into a delivery
plan and a set of implementation-ready tasks. It preserves traceability back to
the PRD while adding the sequencing, dependencies, risks, and verification
details needed for engineering execution.

## When to Use

- After a PRD has been drafted or approved and needs delivery planning
- When a feature is too large to hand directly to engineering as one ticket
- During sprint or milestone planning when requirements must become tasks
- When teams need explicit dependencies, sequencing, and validation work
- When converting product scope into backlog items without losing intent

## When NOT to Use

- To create the PRD itself; use the PRD skill first
- To write only user stories; use the user stories skill when persona-centered
  story format is the primary artifact
- To define only Given/When/Then criteria; use the acceptance criteria skill
- To create launch readiness tasks after implementation is complete; use the
  launch checklist skill
- When the PRD is too vague to identify scope, users, requirements, or outcomes

## Instructions

When asked to generate a plan and tasks from a PRD, follow these steps:

1. **Read the PRD for Intent and Boundaries**
   Identify the product goal, target users, success metrics, in-scope
   requirements, out-of-scope items, assumptions, and constraints. Preserve
   traceability to requirement IDs, section names, or quoted labels when
   available.

2. **Extract Deliverable Workstreams**
   Group requirements into coherent workstreams such as frontend, backend,
   data, integrations, permissions, migration, observability, QA, docs, and
   rollout. Avoid mirroring the PRD section structure if a different execution
   grouping would be more practical.

3. **Sequence the Delivery Plan**
   Build a realistic plan with milestones or phases. Identify ordering,
   dependencies, parallelizable work, critical path items, and integration
   points. Call out decisions or missing inputs that could block execution.

4. **Create Implementation Tasks**
   Break each workstream into tasks that an implementer can understand and
   complete. Each task should include objective, source requirement, owner role,
   dependencies, implementation notes, deliverable, validation, and estimated
   size or complexity.

5. **Add Validation and Release Readiness**
   Include testing tasks, analytics or instrumentation tasks, operational
   readiness, documentation, migration, feature flag, rollout, and rollback
   considerations when relevant.

6. **Surface Risks and Open Questions**
   Separate true delivery risks from normal task details. For each risk, include
   impact, likelihood, mitigation, and the owner or decision needed.

7. **Check for Plan Quality**
   Verify that every major PRD requirement maps to a plan item or task, tasks
   are small enough to estimate, dependencies are explicit, and validation steps
   prove the PRD outcomes.

## Output Format

Use the template in `references/TEMPLATE.md` to structure the output.

## Output Contract

The final artifact must include:

- PRD source summary and planning assumptions
- Requirement coverage map
- Delivery plan with milestones, workstreams, dependencies, and critical path
- Implementation-ready tasks grouped by milestone or workstream
- Validation, instrumentation, release, and operational readiness tasks
- Risks, open questions, and decision log
- Definition of done for the plan

## Quality Checklist

Before finalizing, verify:

- [ ] Every major PRD requirement is covered by at least one task or explicitly
      marked deferred
- [ ] Tasks include source requirement references when available
- [ ] Tasks are implementation-ready, not vague themes
- [ ] Dependencies and sequencing are explicit
- [ ] Validation tasks prove the intended user and business outcomes
- [ ] Risks include concrete mitigations or decisions needed
- [ ] Open questions are separated from assumptions
- [ ] Plan is realistic for delivery, not just a reformatted PRD
- [ ] Out-of-scope PRD items are not accidentally pulled into the task list

## Examples

See `references/EXAMPLE.md` for a completed example.
