# Drawio generation and figure QA

For UML or requested editable Drawio diagrams, use the Drawio workflow below. For ordinary timelines, explanatory schematics or data plots, a deterministic plotting script with PNG plus editable SVG is also suitable; do not require Drawio installation for a non-UML course report. Inspect both the standalone export and final Word size. Never add diagram types merely to satisfy this reference's project checklist.

Plan geometry before generation, reserve whitespace for labels, and reject a figure before Word insertion if any connector or label is ambiguous.

## Source-first workflow

1. Define the figure's purpose, abstraction level, reader, and source-code evidence.
2. Write a small node-and-edge specification before laying out shapes.
3. Choose canvas size and orientation from expected Word width and content density.
4. Create an editable `.drawio` source with explicit coordinates and a consistent monochrome style unless the template requires color.
5. Run `scripts/lint_drawio.py source.drawio --fail-on-error`; resolve errors and assess warnings.
6. Export with the official Drawio desktop CLI. Do not use an unofficial SVG preview as proof that the `.drawio` renders correctly.
7. Inspect the standalone PNG at full resolution.
8. Insert the PNG inline at a deliberate physical width, keep its caption immediately after it, render Word, and inspect it again at final size.

## Phase coverage and required figures

Use the project and assignment requirements to choose figures. Do not add diagrams only to inflate the report.

| Phase | Diagram | Default status | Purpose |
|---|---|---|---|
| Requirements analysis | Use-case diagram | required for multi-actor systems | shows actors, goals, system boundary, and include/extend relationships |
| Requirements analysis | Business activity diagram | recommended for a core cross-role process | shows time order, responsibility handoff, decisions, success path, and failure exits |
| Requirements analysis | Conceptual class diagram | optional | explains domain concepts and their semantic relationships before implementation details; it omits technical service classes and usually omits methods |
| Overview design | Package diagram | recommended when the codebase has meaningful packages or layers | shows namespace/module grouping and dependencies |
| Overview design | Component diagram | required when the assignment asks for architecture/component design | shows deployable or replaceable components, provided/required interfaces, and component dependencies; a generic architecture block diagram is not automatically a UML component diagram |
| Overview design | Deployment diagram | recommended for nontrivial runtime topology | shows devices, execution environments, artifacts, processes, ports, databases, and external services |
| Detailed design | Complete design class diagram | required for object-oriented detailed design | shows important entity, boundary/controller/service, and infrastructure classes with key attributes, operations, associations, multiplicities, and dependencies |
| Detailed design | Sequence diagram | required for at least one core interaction | shows participants, calls, returns, activations, alternative outcomes, and implemented interaction order |
| Detailed design | State diagram | required when a core entity has meaningful lifecycle constraints | shows valid states, transitions, events, guards, and terminal states |

Do not claim coverage by title alone. A class-like ER diagram with only fields is not a complete design class diagram, and a layer diagram without UML component semantics is not a component diagram.

### Conceptual class diagram

This diagram is optional unless the user or assignment requires it. Its job is to stabilize domain vocabulary during analysis: core concepts, important attributes, and semantic associations. It should not be confused with:

- an ER diagram, which emphasizes database keys and storage relationships;
- a design class diagram, which includes implementation classes, operations, interfaces, dependencies, and detailed multiplicities.

If omitted, the report should still explain the domain model through use cases, ER modeling, or prose. Do not silently label an ER diagram as a conceptual class diagram.

## Layout planning

- Prefer one primary reading direction: top-to-bottom for processes, left-to-right for layered architecture, and chronological top-to-bottom for sequence diagrams.
- Allocate label corridors before routing edges. A narrow gap cannot safely hold both several connectors and multi-line labels.
- Keep nodes on a grid with consistent width, height, padding, font size, and lane alignment.
- Use whitespace to separate clusters. Do not solve density by shrinking text below readable Word size.
- Estimate final Word size before choosing font size. A label readable in a 1800-pixel PNG may become unreadable at 14.8 cm.
- For complex diagrams, split by concern instead of forcing every class, endpoint, or table into one page.

## Connector invariants

- Every connector visibly touches the intended node boundary.
- Start and end coordinates must be calculated from actual shape bounds, not guessed nearby points.
- Avoid edge-label backgrounds or masks that erase a segment of the line.
- Put labels above or beside connectors in reserved whitespace. Never place text on a line, arrowhead, activation bar, lifeline, node border, or table box.
- Keep at least one label-height of clearance between a label and unrelated connectors.
- Use orthogonal routes when they make endpoints clearer; keep bend counts low.
- Do not route through nodes, actors, swimlane titles, class compartments, or other labels.
- Avoid unexplained crossings. When a crossing is necessary, change routing or use line jumps consistently.
- Do not merge unrelated associations into a common trunk when it makes endpoints ambiguous.
- Do not add arrows to undirected associations. Use arrowheads that match UML semantics.
- A decision branch must terminate, rejoin explicitly, or loop back through a clear route. Never leave a branch visually open.
- Check that failure paths do not accidentally flow into success actions.

## Diagram-specific rules

### Use-case diagrams

- Actors remain outside the system boundary.
- Each association ends on the actor and use-case boundaries.
- Avoid actor fan-out trunks that look like an extra node or bus.
- Route `include` and `extend` dependencies through unused whitespace; place stereotypes beside, not on, dashed lines.
- The arrow points toward the included use case for `include`.

### Activity and swimlane diagrams

- Each action sits in the lane of the responsible actor or system.
- Use one dominant temporal direction.
- Put decision questions inside a sufficiently large diamond or beside it without touching connectors.
- Put `yes` and `no` labels in whitespace beside their branch, not on the branch.
- Each success and failure path has an explicit result or terminal node.
- Avoid long backtracking lines that cross multiple actions; use a separate failure terminal when a retry loop would make the diagram ambiguous.

### Component diagrams

- Use UML component notation or clearly stereotyped components.
- Show important provided/required interfaces when they matter.
- Dependencies must follow real build/runtime calls.
- Distinguish a component from a package, deployment node, or database entity.

### Class diagrams

- A conceptual class diagram contains domain concepts and semantic relations, with no pressure to show methods.
- A complete design class diagram includes key attributes and operations, implementation-level class categories, associations, multiplicities, inheritance/realization, and significant dependencies.
- Keep role and multiplicity labels off relationship lines and class boxes.
- Do not use an ER schema with English table fields as a substitute for a design class diagram.
- Split a complete design model into focused diagrams if one page would become unreadable, and provide an overview relationship figure.

### Sequence diagrams

- Lifelines align vertically and messages follow chronological order.
- Activations begin and end at the actual calls they represent.
- Keep each message label within the horizontal interval between adjacent lifelines; wrap it if necessary.
- A message that crosses an intermediate lifeline must place its label entirely in one free interval, not across that lifeline.
- Put return labels above dashed return arrows and away from activations.
- Use `alt`/`opt` frames when alternatives are essential; a note in the caption does not replace branching semantics.

### State diagrams

- States are nouns or stable conditions; transitions are events/actions/guards.
- Initial and final nodes are distinct and connected.
- Show only legal transitions and mark terminal states.
- Avoid duplicating the activity diagram; focus on one entity's lifecycle.

### Deployment diagrams

- Distinguish device, execution environment, artifact, database, and external service.
- Label protocols and ports beside connectors.
- Do not place protocol labels over node borders or network lines.

## Export and inspection

Use the official CLI, for example:

```powershell
draw.io.exe --export --format png --scale 2 --border 20 --output output.png source.drawio
```

After export, confirm:

- nonzero exit code is treated as failure;
- PNG opens and has plausible dimensions and byte size;
- all Unicode glyphs render;
- text is readable at full size and at expected Word size;
- no cropped labels, hidden arrows, broken lines, duplicated arrows, or label/line collisions;
- orthogonal routes do not cross nodes or unexplained swimlane boundaries;
- all branches and returns have correct direction;
- `.drawio` and PNG basenames match and the PNG was generated after the source.

## Word insertion QA

- Prefer inline images.
- Preserve aspect ratio and scale down from a high-resolution PNG.
- Keep the image paragraph and caption together when possible.
- Verify the caption number and cross-reference after fields update.
- Inspect the rendered Word page, not only the standalone image.
- If Word scaling makes text small, redesign the diagram or split it; do not rely on the reader zooming in.

## Rejection conditions

Regenerate the figure before insertion if any connector is visually detached, any label covers a line, an arrow is semantically wrong, a class diagram lacks the information required by its claimed type, a mandatory phase diagram is absent, or the final Word-size preview is hard to read.
