# Project evidence and report writing

## Evidence map

Build `project_evidence.md` before writing. For each claim, record:

| Claim area | Evidence to inspect | Minimum useful evidence |
|---|---|---|
| Architecture | entrypoints, package tree, configuration, startup scripts | concrete components and call direction |
| Features | routes, controllers, services, UI handlers | file and symbol implementing each feature |
| Data model | schema, migrations, ORM/models, seed data | entities, keys, relationships, constraints |
| Roles and permissions | auth middleware, dependencies, guards | actor, allowed action, rejection behavior |
| Business flow | service logic, state machine, validations | success path, failure exits, persistence effects |
| Algorithms | implementation and configuration | inputs, steps, outputs, complexity or scoring rule |
| Deployment | scripts, ports, databases, external services | actual host, process, and network relationship |
| Testing | test files and fresh execution | command, result, and coverage limits |
| UI | source plus current screenshots | page purpose and implemented interaction |

Prefer current code and fresh execution over README claims. If the repository and prose disagree, use the implementation and note the mismatch internally.

## Content map

Create `content_map.md` with one row for every output block:

| Template location | Output kind | Project evidence | Planned content | Status |
|---|---|---|---|---|
| Heading or placeholder | heading/body/table/caption/figure | file and symbol | concise draft or data source | verified/draft/blank |

The map prevents missing sections, duplicated claims, orphan captions, and fabricated formal metadata.

## Writing rules

- Match the template's academic or professional register.
- Explain how components relate; do not turn file names into prose.
- Tie every diagram and table to the surrounding analysis.
- Distinguish implemented, tested, planned, assumed, and unverified behavior.
- Use the project's real names, endpoints, states, fields, and constraints.
- Do not inflate responsibilities, test coverage, performance, security, or deployment maturity.
- Unknown name, student ID, date, signature, stamp, score, and approval fields remain blank.

## Report completeness

Typical software-project coverage includes project overview, requirements, actors/use cases, business process, architecture, module design, data model, key algorithms, interface or API design, deployment, testing, UI evidence, limitations, conclusion, and references. Follow the supplied template rather than forcing this list when its structure differs.
