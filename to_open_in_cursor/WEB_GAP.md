# Gap analysis: VBA against the web app

Compared `to_open_in_cursor/` (see `HOW_THE_VBA_WORKS.md`) with `web/js/org.js`, `web/js/layout.js`, `web/js/app.js`, and `web/js/export.js`.

The web app covers the ribbon path: import a 13-column sheet, reject a broken reporting line, draw a fitted tree, and offer the manual branch, shade, link, and arrange commands. The branch geometry (one up-tick, left spine, two-per-row, four-per-row, full tree) follows the VBA. The gaps below are the places a chart, a number, or a command still comes out differently.

## What already matches

- Columns A–M, title row on by default, blank rows skipped, empty surname or name rejected.
- `FullNameKey` is the lowercased name with spaces removed. The second person with the same name gets suffix `"2"`, the third `"3"`.
- One person with an empty solid-line manager is the CEO. Zero or several is a fatal error. A missing manager, a dotted manager equal to the solid one, and a manager name that is duplicated all block the chart.
- Solid links are the ones that expand. A dotted report is drawn as a card with a bubble, not as its own tree. `DefaultManagerType` in the VBA is the literal `"Solid"`; the web uses the same rule.
- Headcount walks solid links, counts the person, skips names starting with `TBH_`, and hides the badge when the count is 1 or less. The statistics sentences use the palette wording (`Team Size – N people`, ordinary, non-ordinary, outsourcers).
- Direct reports are ordered by team size, largest first, then by key.
- Worker columns are drawn before manager branches, then branches with the same department are pulled together.
- `branch_auto`: 1 report is an upward tick, 2–6 is the left spine (`createBranch1D`), 7–14 is two per row, 15 or more is four per row. The shift lists are `0`, `−0.5 0.5`, `−0.5 0 0.5`, and `−1.5 −0.5 0.5 1.5`. A short last row keeps the leftmost shifts.
- The spine is `originX − cardWidth/2 − gapX`. The tick meets the left edge at mid-height. The trunk stops at the last tick of that style. Solid is `RGB(91, 155, 213)`. Dotted is `RGB(0, 32, 96)`.
- Auto-fit starts at Small and tick 5, hides the largest solid-line team until the tree is within 920×350, then tries Medium and Large. A company of 100 or more starts collapsed and reveals the smallest teams that still fit. More than 7 side-by-side branches with every manager hidden falls back to four per row.
- Full org structure gives each hidden solid-line manager a slide, with “See slide N” on the folded card and “Up to …” on the child slide.
- Direct reports (`branch_clever`) treats managers as workers. Subordinates (`full_tree`) expands every solid-line manager and does not run the fit. 1-in-row, 2-in-row, and 4-in-row call the same layouts.
- The ribbon has the same commands, in the same order, including sticky sizes, group, ungroup, align, and distribute.

## Visible differences

These change the slide.

| What | VBA | Web |
| --- | --- | --- |
| Head card | `makeHeadOfTree` forces Large and a 15-point tick pointing up, then hides the tree’s own top tick and shifts the tree up by that tick’s height. | The head uses the body card size. The drop under it is the tree’s own tick (`gapY`), not 15 points. |
| Tick growth | X grows in steps of 1/3, capped at half the card width. Y grows in steps of 1, capped at a third of the card height. | One number is used for both. It grows by 1 and stops at the smaller of those two caps. |
| Gap between branches | `inputWhiteSpaceExt` adds `Abs(treeWidth − 920) / (recursiveBranchCount − 1)`, whether the tree is narrower or wider than 920, capped at the card width. The tree is not measured again. | Extra gap is added only when the root tree is narrower than 920, using the root’s branch count, and only if the result still fits. |
| Card size floor | After hiding teams, the size loop starts at Small and never goes back to ExtraSmall. | If Small still overflows, ExtraSmall is used when it is smaller. |
| Termination | A comment box is placed under the card. Its height is added to the stack, to a row, and to the parent’s drop when the parent has a termination date. | “Works until …” replaces the location line inside the card. The stack does not grow, and the location is not shown on that card. |
| Two-manager alignment | `checkSolidLineDirectReportsNeedShiftDown` pushes a lone card down by one bubble height and stretches its tick so the row lines up. | That shift is not done. A row that mixes bubbles does lengthen every tick in the row, which is the ND part of the same idea. |
| Bubble on an expanded manager | The bubble sits `gapY/2` above the card, and a second tick reaches its center. | `leafLayout` and the 1D stack do that. `layoutExpandedManager` puts the bubble flush on the card and does not draw the second tick. |
| Department shades | Each shade starts as the slide guide (140, 15, 945, 505) and is pulled in to the nearest card on each side. A shade that contains another is reordered. Clicking an existing shade lightens the template color. | Auto slides use a padded bounding box of the cards and stretch every shade down to the lowest card. The Departments button uses a smaller pad and does not stretch them to a common bottom. Containment is not reordered. |
| Card face | A copy of the palette group. Outsource overwrites Important when both flags are set. Fonts are 7/4, 9/5, or 11/6. Headcount turns white when there is no location, except on an outsource-only card. | Drawn in CSS. Important is chosen before outsource. Type sizes are chosen from the card height, not from those pairs. The footer is always a dark bar when location or termination is present. |
| Bracket tags | `SurnameModerated` removes the text inside `[]` and leaves the brackets. | `moderated` removes the whole `[tag]`, including the brackets, and the card shows that. |
| “Up to” on the Links button | The label is the person’s manager (`FullNameManagerKey`, which is the dotted manager when they have one). | The label is the person who appears on both slides. The infinity slide’s “Up to” uses the parent slide’s person, which is the other way around. |

The fit differences matter on a large sheet: which teams fold, and whether the tick is 10 or 13, can change even when every connector rule matches.

## Data and checks

| What | VBA | Web |
| --- | --- | --- |
| How many errors | The first `MsgBox` stops the import. | Every error is listed, and the chart stays blocked until none remain. |
| Dotted chain | `verifyReportingSequence` runs for `"solid"` and `"dotted"`. After the first hop the walk follows `FullNameManagerKey`, so a later two-manager person is followed through the dotted key. | `reachesCeo` follows `solidKey` only. A solid chain that reaches the CEO passes even when the dotted chain does not. |
| Both flags set | Outsource and important are counted separately. Ordinary is neither. The card template is Outsource. | The person is counted as outsource and not as important. The card is styled important. |
| CEO name spelling | `checkPersonExists` tries every order of the words against the report dictionary. | The CEO is the row with an empty solid-line manager. A differently ordered name is not matched. |
| Not enough columns | Windows always reads A:M. Mac keeps a row only when `UBound(cells) >= 13`, which rejects a 0-based 13-cell row (`UBound` 12). | The sheet is rejected when no row has 13 cells. A shorter row on a wide sheet is read with empty cells. |
| Categories | A–Z are always in the dropdown. Choosing “All, A–Z” always shows the 1000-person message. The list itself is capped at 1000. | Only letters that occur are listed. The list is capped at 1000 and no message is shown. |
| Cost center | Integer when every character is a digit, otherwise left at 0. `showCostCenters` stamps a green box on each card. | Stored as text. The Cost centers button stamps a green box that grows with its text. |
| Multiple CEOs | The message prints the array index as the line number. | The message prints the Excel row. |
| Forgotten people | `showWhomIForgot` scans shape names for `InfoBox`, buckets missing solid-line cards by slide, and the duplicate report is commented out. It is not on the ribbon. | The inspector says how many people in the solid-line tree are absent from the slides. It does not list them by slide, and it does not report duplicate cards. |

## Commands and output that are absent

- Palette slide. The VBA copies real PowerPoint groups. The web draws cards and writes a new pptx with `pptxgenjs`. Export has the same positions and line colors. It does not copy the template groups, and “See slide” / “Up to” are text, not hyperlinks.
- Progress form and Cancel (`stopExecutionFlag`).
- The size-check `MsgBox` when the drawn group and `sizeOfTree` differ by more than 25 points.
- Cost-center stamps and “erase text boxes”.
- Clicking a card does not push the dropdown, the box size, and the tick the way `CEventClass` does. Sticky sizes only sets the size dropdown from the card height, and only after a click, within 5 pixels. It does not read the tick.
- Manual draw does not delete descendant shapes the way `eraseSubordinates` does. Direct reports, 1-in-row, 2-in-row, 4-in-row, and Subordinates clear the slide down to the title, legend, statistics, and the selected card, then draw the branch under that card.
- The Manager button places a Large card and the title block. It does not draw the 15-point upward tick or delete the legend the way `element_head` does before rebuilding it.
- Empty department becomes the key `<no_department_specified>` in the VBA and still gets a shade. The web skips a card with a blank department when it builds shades.

## VBA behavior the web does not copy

These are in the source and called out in `HOW_THE_VBA_WORKS.md`. The web takes the reading that matches the comment or the later branch.

- `create_Links` assigns `slideNumberToBeLinked` twice when the first slide index is greater, and never sets `slideNumberToLink`. The web Links button sorts the two slide numbers and always puts “See slide” on the earlier one.
- `sizeOfFullTree` compares `ManagerType` with the literal `"Solid"`. `makeFullTree` compares it with `DefaultManagerType`. Both are `"Solid"` after 30/03/2018, and the web uses `"Solid"`.
- The fourth fit step adds gap even when the tree is already wider than 920. The web only uses space that is left.
- On a company of 100 or more, a manager who does not fit is set back to hidden and remains eligible, so the search can pick them again. The web stops.

## Only in the web app

Undo, drag, marquee select, zoom, PNG, and a JSON project file. Import errors stay on the page instead of a `MsgBox`. A chart that is wider than the slide is shifted so the cards stay on the page (`containOnSlide`); the VBA centers the group and can leave a 1D tick correction on top of that.
