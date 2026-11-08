# How the MyBoss VBA works

This is a reading of the source in `to_open_in_cursor/`. It was written by going through every module, class, form, the ribbon XML, and the Mac script. Line numbers below refer to those files. Where a statement is surprising, the condition is quoted so it can be checked against the file.

The add-in is a PowerPoint ribbon (`customUI/customUI14.xml`, tab label `Org chart`). It loads one Excel sheet into dictionaries of people, then copies palette shapes from a template slide and moves them into an org chart. Drawing never computes a card’s pixels from scratch: a card is a copy of `CardOrdinaryMedium` (or Important / Outsource / Tbh, and a size suffix), and a connector is `Shapes.AddLine`.

Most modules have no `Option Explicit`. A name used inside a procedure and never declared is a procedure-local `Variant`. `Dim a, b As Double` types only `b`. `And` and `Or` evaluate both sides.

---

## 1. What happens, in order

### Ribbon load

`customUI14.xml` sets `onLoad="MyAddInInitialize"` (`utilInitialize`).

`MyAddInInitialize` stores the ribbon (`RibbonOnLoad` in `utilInitializeRibbon`), hooks `CEventClass` to `Application`, and creates the dictionaries:

| Dictionary | What it holds |
| --- | --- |
| `MyDictionaryPersonalInfo` | `FullNameKey` → `CEmployee` |
| `MyDictionaryDirectReports` | manager key → `Collection` of report keys |
| `MyDictionaryEmployeesPerCategory` | `"CEO"`, `"Managers"`, `"AZ"`, `"A"`…`"Z"` → array of keys |
| `MyDictionaryCardsCheckList` | shape-name key → collection of slide indexes |
| `MyDictionaryCeoCandidates` | key of a person with an empty solid-line manager |
| `MyDictionaryDepartmentsList` | declared, filled later by department code |
| `MyDictionarySubordinatesChecklist` | card names that should exist under a manager |

It also sets:

- `sheetHasTitlesFlag = True` (first data row is Excel row 2)
- `defaultInputNumberOfColumns = 13`
- `defaultHomeCity = "Russia, Moscow"`
- `defaultNumberOfShades = 3`
- `MyDictionaryCardSizes` = ExtraSmall, Small, Medium, Large
- `MyDictionaryTickLengths` = 5, 7, 10, 12, 15
- `maxWidthGuideline = 920`, `maxHeightGuideline = 350`
- `MAX_PEOPLE_FOR_2D_BRANCH = 6` (this constant selects the **1D** stack; see `makeTree`)
- `MAX_PEOPLE_FOR_4D_BRANCH = 14`
- `MAX_BRANCHES_HORIZONTALLY = 7`
- slide guide `const_slideTopBorder = 140`, right `945`, bottom `505`, left `15`
- `autodetectSizeAndWhiteSpaceFlag = True`
- card size and tick dropdowns enabled

Then `initPalette` / `findFullPalette` (`utilTestTemplate`) walk slides until one contains every required template shape. That slide’s number is `paletteSlideNumber`. Each template’s width is stored as collection item 1 and its height as item 2. A welcome `MsgBox` names that slide.

### Picking the workbook

`pickExcelFile` (`setupStage2LoadInputFile`) calls `browseFile` (`utilOpenDialog`).

- Windows: `Application.FileDialog(msoFileDialogFilePicker)`, one file.
- Mac: `AppleScriptTask("MyBoss-browse_files_on_mac.scpt", "openDialog", seedPath)`. A result that starts with `-` and is otherwise numeric is treated as a cancel (`browseFile = ""`). Otherwise `:` becomes `\`, and the text `Macintosh HD` is stripped.

The extension test keeps `.xls`, `.xlsx`, and `.xlsm`. The failure message says `.xls` or `.xlsm` and does not mention `.xlsx`. The procedure does not check that the filename is Latin.

Sheet names:

- Windows: open with `CreateObject("Excel.Application")`, copy `Worksheets(i).Name`, close without saving.
- Mac: `openExcelGetSheets`. `textInBraces` takes the text between the first `{` and `}`. A trailing comma is dropped. Names are split on `,`.

`mCurrentSheet` is `ReDim`’d and left empty. Invalidating `sheetName` makes the ribbon call `uiGetSheetId`, which stores `"sheetIndex1"` and `inputSheetName = MyVariantListOfSheets(0)`.

`sheetHasTitlesSwitch` toggles `sheetHasTitlesFlag`. The `pressed` argument is ignored.

### Import

`uploadExcelFile`:

- Titles on → `inputFirstRow = 2`. Titles off → `1`.
- Windows reads `A{inputFirstRow}:M{UsedRange.Rows.Count}` (`Chr(Asc("A") + 12)` is `M`). `Range.Value` replaces the array and is **1-based**.
- Mac reads the sheet in chunks of 100 through `openExcelGetSheetData` (`path|sheet|offset|100`). Rows are split on `;`, cells on `|`. A row is kept only when `UBound(eCells) >= 13`. A 0-based 13-cell split has `UBound` 12, so that test rejects a 13-field row and jumps to `askToCheckTheData` (“doesn’t have enough columns”). Cells 0..12 are copied, then `wfTransposeArray` swaps the dimensions into a **0-based** array.
- `buildHierarchy` then `checkConnectionsToCEO`. Both must return true.
- `CalculateSubordinates(fullNameTopLevelManagerKey, "number")` is stored on the CEO’s `HeadcountStats`.
- Keys are sorted alphabetically into `MyVariantListOfEmployees`.
- `getEmployeesListByCategory` fills CEO, Managers, AZ, and A–Z.
- `inputCategoryName = "CEO"`. Creation controls become visible (`fileUploadingSuccessful = True`).

### Drawing a chart

Every draw button calls `mainSwitch` (`createTree_Controller`) with a regime string. The anchor is the bottom-center of the selected card when that card is the person in the dropdown; otherwise `posX = SlideWidth / 4` and `posY = 70`.

| Button | Callback | Regime | What it draws |
| --- | --- | --- | --- |
| Single slide org structure | `createFullTreeClever` | `full_tree_clever` | `makeFullTreeAuto(..., False)` on a new slide after the current one |
| Full org structure | `createFullTreeInfinity` | `full_tree_infinity` | `makeFullTreeAuto(..., True)` on a new slide at the end; hidden managers get their own slides |
| Manager | `create0DKingBranch` | `element_head` | `makeHeadOfTree` plus the title/statistics group |
| Employee | `create0DBranch` | `element_single` | one card, tick pointing up |
| Direct reports | `createAutoDBranch` | `branch_clever` | subordinates, auto layout, managers treated as workers |
| 1-in-row | `create1DBranch` | `branch_1` | `make1DSubordinates` |
| 2-in-row | `create2DBranch` | `branch_2` | `makeNDSubordinates` with `"2"` |
| 4-in-row | `create4DBranch` | `branch_4` | `makeNDSubordinates` with `"4"` |
| Subordinates | `createFullTreeStraightforward` | `full_tree` | full tree of the selected person, on the current slide |
| Departments | `createDepartments` | — | shades, or erase them if they already exist |
| Links | `createLinks` | — | “See slide N” and “Up to …” |

A subordinate tree is drawn only when the person has direct reports, and when they do **not** have two managers or the card being drawn is the solid-line card:

```
(HasTwoManagers And managerType = "Solid") Or Not HasTwoManagers
```

A dotted card of a two-manager person does not grow a tree. The commented test dated 30/03/2018 used `"Dotted"` instead. `DefaultManagerType` and `ShouldBeShownUnderManager` were changed the same day so charts expand under the solid line.

---

## 2. The person

`CEmployee` is one Excel row. Fields match columns A–M:

| Offset | Column | Property |
| --- | --- | --- |
| 0 | A | `Surname` |
| 1 | B | `Name` |
| 2 | C | `JobTitle` |
| 3 | D | `Department` |
| 4 | E | `SurnameSolidLineManager` |
| 5 | F | `NameSolidLineManager` |
| 6 | G | `SurnameDottedLineManager` |
| 7 | H | `NameDottedLineManager` |
| 8 | I | `Location` |
| 9 | J | `TerminationDate` (kept as text) |
| 10 | K | `OutsourceFlag` |
| 11 | L | `ImportantFlag` |
| 12 | M | `CostCenter` (integer; non-digits stay 0) |

`FullName` is `Trim(Surname) & " " & Trim(Name)`. The space is always there.

`FullNameKey` is `LCase(Replace(FullName, " ", "")) & DuplicateId`. Spaces are removed, then the duplicate suffix is glued on with no separator.

Manager keys (`FullNameSolidLineManagerKey`, `FullNameDottedLineManagerKey`, `DepartmentKey`) are the same lowercase-and-strip, **without** `DuplicateId`.

`SurnameModerated` / `NameModerated` remove the text between the first `[` and the first `]`. The brackets themselves stay. `Smith [temp]` becomes `Smith []`.

Flags K and L are true only for the strings `"1"`, `"1.0"`, `"1,0"`, or the numbers `1` and `1#` (`utilStage1SaveInfoAboutEmployee`). Anything else, including blank, is false.

`IsToBeHired` is true when `FullName` starts with `TBH_` (case-sensitive, so the surname must start with it).

`IsManager` is `MyDictionaryDirectReports(FullNameKey).Count > 0`. A person who is only someone else’s dotted manager still counts, because `buildHierarchy` pushes the report onto both managers’ collections.

`HasTwoManagers` is true when the dotted key is not `""`. The solid key is not consulted.

`ManagerType(key)`:

- argument equals the solid key → `"Solid"` (this arm is first, so two identical keys, including both empty, return `"Solid"`)
- argument equals the dotted key → `"Dotted"`
- otherwise `""`

`FullNameManager` / `FullNameManagerKey` return the **dotted** manager when `HasTwoManagers` is true, and the solid manager otherwise. `DefaultManagerType` does not follow that. After 30/03/2018 it is the literal `"Solid"`. `ShouldBeShownUnderManager` is true only when `ManagerType(argument) = "Solid"`.

`DirectReports("sorted")` sorts with `sortUtility(..., "maturity", False)`: larger `HeadcountStats.SubordinatesNumber` first; a tie is uppercase key ascending. Any other regime, including `"Unsorted"`, keeps insertion order. The result is a dictionary of five collections: `All`, `Solid`, `Dotted`, `Workers`, `Managers`. A report is a worker when their own `IsManager` is false.

`ListOfDepartments(class)` walks `DirectReports("Sorted")` of that class. The key is `DepartmentKey`, or the literal `<no_department_specified>` when the department is blank. Order inside a department is the maturity order.

`nOrdinaryBranchesUtil(False)` is the manager-report count. Passed `True`, it is 0 (every report is treated as a worker).

`nDepartmentBranchesUtil(False)` is the number of distinct departments among workers. Passed `True`, it counts departments among all reports.

`nBranchesHorizontallyLimitExceeded` is true when those two counts sum to **more than** 7.

### Duplicates

`buildHierarchy` looks up the key before `DuplicateId` is set, so the first person keeps the unsuffixed key. Each later collision increments `Duplicates` on that first object and sets the new row’s `DuplicateId` to `CStr(Duplicates + 1)`. The second copy is `"2"`, the third `"3"`. `"1"` is never written. Later copies keep `Duplicates = 0`. An empty `FullNameKey` is counted as an empty row and is not inserted.

### Headcount

`CalculateSubordinates` (`utilCountPeople`) walks **solid** links only. A dotted report is not visited, and neither are the people under that dotted report.

For `outputFormat = "number"`:

- `SubordinatesNumber` counts the person themselves, plus everyone reached by solid links, except names whose `FullName` starts with `TBH_`.
- `SubordinatesInHomeCityNumber` is the same set whose `Location` equals `defaultHomeCity` exactly (`"Russia, Moscow"` unless something else assigned it).
- `OrdinaryNumber` is the same set with both `OutsourceFlag` and `ImportantFlag` false.
- `OutsourceNumber` and `ImportantNumber` are separate. A person with both flags increments both and is not ordinary.
- `SubordinatesCardsInFullTree` counts every visited person, including TBH.
- The result is stored on that person’s `HeadcountStats`. A later `"number"` call returns the stored object when `IsDefault` is false (`IsDefault` means all six counters are `-1`). `ResetToDefault` sets that sentinel; `ResetToZero` sets zeros.

`outputFormat = "checklist"` records `reportKey & ManagerType` for every direct report of a visited person (dotted reports included, their teams not walked) and `personKey & "Solid"` for the visited person. It does not update the counters.

`CSubordinatesStatistics.ByType` maps `"cards"`, `"all"`, `"allinhomecity"`, `"ordinary"`, `"important"`, `"outsource"`. Any other string returns `-1`.

### The three checks before a chart is allowed

`buildHierarchy` (`utilStage2`):

- 0 people with an empty solid-line manager: “Couldn't identify CEO… CEO should have emply Solid Line Manager fields.” (`emply` is the spelling in the file.) Returns false.
- exactly 1: that key becomes `inputCEOFullName`. `inputCompanySize` is the count of uploaded people who are not `TBH_`. Returns true.
- more than 1: lists them. The “line” number in that box is the **array index**, not `RowIndexInExcelFile`.

`checkConnectionsToCEO` (`utilStage3`):

1. `checkPersonExists` (`utilCheckWordCombinations`) splits the CEO string on spaces and tries every concatenation of those words against `MyDictionaryDirectReports`. The first hit wins. `fullNameTopLevelManagerKey` is that hit. The CEO row already has a key with no spaces, and `buildHierarchy` created a direct-report collection under it, so the identity permutation matches.
2. Check 1, everyone except the CEO: the solid manager key must exist. An empty key says “You didn't specify manager”. A missing key says “There is no such manager”. If there is a dotted key, it must differ from the solid key and must exist.
3. Check 2: if the solid manager’s `Duplicates > 0`, the upload stops and lists `Duplicates + 1` people (the unsuffixed key, then key+`"2"`, key+`"3"`, …). The dotted arm does the same, but the sentence inserts `FullNameDottedLineManagerKey` and the words “solid-line manager”.
4. Check 3: `verifyReportingSequence` for `"solid"` and for `"dotted"`. The walk stops when the next key is the CEO’s solid key (empty for the CEO) or when the next key is already in the chain. The last key in the chain must be `fullNameTopLevelManagerKey`. Otherwise a `MsgBox` lists the chain and says to check for loops or self-reporting. After the first hop, `extractReportingSequence` follows `FullNameManagerKey`, which is the dotted key when that person has two managers.

`showWhomIForgot` (`utilStage4`) is **not** part of import. It compares shapes already on slides with the solid-line checklist and shows “Number of forgotten … subordinates”. It is not wired to the ribbon (`checkPeople` exists and is not in `customUI14.xml`).

---

## 3. Card, tick, line

Palette names (`utilTestTemplate`):

- No size suffix: `DepartmentGroup`, `AttentionFrame`, `NavigationArrow`, `Shade1Department`, `Shade2Department`, `Shade3Department`.
- Plus Large / Medium / Small / ExtraSmall: `CardOrdinary`, `CardImportant`, `CardOutsource`, `CardTbh`, `BubbleOrdinary`, `BubbleImportant`, `BubbleOutsource`, `SlideN`, `CommentBox`.

`copyShape` picks the template from the person:

- `IsToBeHired` → `Tbh` (and does not look at the other flags)
- else `ImportantFlag` sets `Important`, then `OutsourceFlag` overwrites it with `Outsource`
- neither flag → `Ordinary`
- a call with an empty person name → `shapeName & inputBoxSize` and no type word

`getCardWidth` / `getCardHeight` always read the **Ordinary** card of the requested size, even when the drawn card is Important or Outsource. The numbers are the palette shape’s width and height, not literals in the VBA.

`getLineWeight` (`TreeManagement_GetDimensions`): ExtraSmall 2, Small 2, Medium 3, Large 4, anything else 3.

`getBubbleHeight` / `getBubbleWidth` are the Ordinary bubble’s palette size plus `2 * (getLineWeight() / 3)`.

`getCommentBoxHeight` reads `CommentBox` + size, capital C and B, with no line-weight padding. The comment box copied at draw time asks for `"commentBox" & inputBoxSize` (lowercase c). Those are different strings.

### `createOrgChartElement` (`buCreateCard`)

The card is built as if its top-left is `(0, 0)`, grouped with its tick, then moved so the **far end of the primary tick** lands on the `(X, Y)` passed in. Y grows downward. A negative line length goes up or left.

| `tickDirection` | Comment in the file | Line |
| --- | --- | --- |
| 1 | pointing UP | from `(cardWidth/2, 0)` upward by `inputWhiteSpaceY`. If the person has two managers, the length also adds `getBubbleHeight()`. |
| 2 | pointing RIGHT | from the right edge at mid-height, right by `inputWhiteSpaceX` |
| 3 | pointing DOWN | from bottom-center, down by `inputWhiteSpaceY` |
| 4 | pointing LEFT | from the left edge at mid-height, left by `inputWhiteSpaceX` |

Directions 1 and 4 place the apex at the tick’s `(Left, Top)`. Directions 2 and 3 place it at `(Left + Width, Top + Height)`. The second move is `X - (apexX - X)` and the same for Y, and the Y correction is stored in an `Integer` (`Dim tickApexDx, tickApexDy As Integer` types only `tickApexDy`).

Two managers, only when the dotted key is non-empty:

- A second tick, style `inverseType(ManagerType)`, from the top-center of the card upward by `inputWhiteSpaceY/2 + getBubbleHeight()/2`. On direction 1 its X is `cardWidth/2 + 5` (the only place the two upward ticks are separated). On directions 2, 3, and 4 it stays on the center.
- A bubble copied from the **other** manager’s template, centered on the card. Its top is at `-inputWhiteSpaceY/2 - getBubbleHeight()`, so it sits `inputWhiteSpaceY/2` above the card and its vertical center meets the extra tick. Text is the other manager’s raw `Surname`. Outline weight is `getLineWeight()/3`.

A non-empty `TerminationDate` adds a comment box flush under the card. `fillCommentBox` tries to parse strings that contain `"at"`. The search is case-sensitive, so `"Saturday"` matches. A string like `"Wednesday, January 15, 2020 at 3:00:00 PM"` becomes `Format(..., "d mmmm yyyy")` appended to the template text. If the cleaned text is not a date, the shape text is replaced by that raw string.

`fillCard`:

- To be hired: `"To Be Hired"`, then the job title. Name is not written.
- ExtraSmall: surname, name, and an empty third paragraph. Job title is omitted.
- Other sizes: surname, name, job title.
- Fonts: ExtraSmall and Small are 7 / 4, Medium 9 / 5, Large 11 / 6. Paragraphs 1–2 are bold at the name size. The job paragraph is not bold.
- Headcount text is `SubordinatesNumber`, or `SubordinatesNumber & " (" & OutsourceNumber & ")"` when the outsource count is greater than 0. `HeadCount.Visible = False` when `SubordinatesNumber <= 1`.
- An empty or whitespace `Location` hides `BottomBackground` and `Location`. Headcount text is then set to white for an ordinary card and for every important card. An outsource-only card does not get that color assignment. A non-empty location is written as stored; visibility is left as the template had it.

The grouped element is named `personFullName & ManagerType` (`…Solid` or `…Dotted`). Inner pieces are `Tick`, `Card`, `InfoBox`, `HeadCount`, `Location`, `BottomBackground`, `MainBackground`, plus the same suffix.

### Line style (`buCreateConnector`)

`applyLineStyle`, after `LCase`:

- `"solid"`: dash is set to round-dot and then immediately to `msoLineSolid`. Color `RGB(91, 155, 213)`.
- `"dotted"`: dash is set to round-dot and then to `msoLineSysDot`. Color `RGB(0, 32, 96)`.
- Weight is `getLineWeight() / widthCorrection`. Connectors pass 1. The bubble passes 3.

`inverseType("solid")` is `"Dotted"`. `inverseType("dotted")` is `"Solid"`.

### `sizeOfOrgChartElement`

Direction 1 (up): width is the card. Height is card + `inputWhiteSpaceY` + bubble height if two managers + comment height if terminated.

Direction 4 (left): width is card + `inputWhiteSpaceX`. Height is the card, plus `inputWhiteSpaceY/2 + getBubbleHeight()` if two managers, plus the comment height if terminated.

Directions 2 and 3 are not handled. The size stays at the bare card.

---

## 4. How a branch is drawn

The origin `(originX, originY)` is the point just under the parent card. `inputWhiteSpaceX` and `inputWhiteSpaceY` start equal. The dropdown `Tick size, pt.` sets both, and `inputWhiteSpaceExt`, to 5, 7, 10, 12, or 15. Default id is `whiteSpaceCoef10`. Auto-fit changes X and Y independently afterward; the dropdown is not read again.

### One card (`branch_auto` with exactly one report)

`createOrgChartElement(..., originX, originY, 1)`. Tick points up. No trunk.

### 1D stack — `make1DSubordinates` (`createBranch1D`)

Used for `branch_1`, and for `branch_auto` when the unsorted count is 2 through 6 (`<= MAX_PEOPLE_FOR_2D_BRANCH`).

```
OffsetX = -getCardWidth() / 2 - inputWhiteSpaceX
OffsetY = inputWhiteSpaceY
(+ getCommentBoxHeight() when the PARENT has a termination date)

trunk X  = originX + OffsetX
elbow Y  = originY + OffsetY
```

Reports are walked in maturity order (largest team first). For each one, `yIterator` moves down by `inputWhiteSpaceY + cardHeight/2`, plus `inputWhiteSpaceY/2 + bubbleHeight` when that report has two managers. The card is created with tick direction **4**. The left-tick apex is placed on the trunk at that Y, so the tick enters the middle of the card’s left edge and the card’s center is on `originX`. Then `yIterator` moves another `cardHeight/2`, plus the comment height when **that child** has a termination date.

Connectors, after the cards:

- Solid, when `Dotted.Count < All.Count` (at least one report is not dotted):
  - vertical from `(originX, originY)` down by `OffsetY` — the drop from the parent
  - horizontal from `(originX, elbow Y)` by `OffsetX` — left, to the trunk
  - vertical trunk from the elbow down to `lowestSolidLineDirectReportY`
- Dotted, when `Dotted.Count > 0`: the same three lines, ending at the last dotted apex.

`lowestSolidLineDirectReportY` is the apex Y of the **last** solid report, not the bottom of the card. The trunk stops there. Both styles share the same X. Dotted is drawn second.

Width used by `sizeOf1DSubordinates` is `inputWhiteSpaceX + getCardWidth()`. Height adds, per report in the **unsorted** collection, `inputWhiteSpaceY` plus the direction-4 element height. The parent `OffsetY` is added later by `sizeOfTree`, not here.

### Rows — `makeNDSubordinates` (`createBranchND`)

`createHorizontalShiftsMap`:

| Cards on the row | Shifts |
| --- | --- |
| 1 | `0` |
| 2 | `-0.5, 0.5` |
| 3 | `-0.5, 0, 0.5` |
| 4 | `-1.5, -0.5, 0.5, 1.5` |

An even count skips shift 0. A short last row uses the **leftmost** entries of that map. It is not recentered.

```
OffsetX = 0
OffsetY = inputWhiteSpaceY
(+ comment height when the parent is terminated)
trunk X = originX
first connector Y = originY + OffsetY
```

Each row is cut from the maturity-sorted list, up to N people. For person `i`:

```
hLineLength = shift(i) * (inputWhiteSpaceX + getCardWidth())
```

The card is created at `(originX + hLineLength, yIterator)` with tick direction **1**. The up-tick apex sits on the row’s horizontal. A stub is then drawn from `(originX, yIterator)` by `hLineLength`, in that person’s solid or dotted style.

If the row contains anyone with two managers, a person who does **not** have two managers gets `inputWhiteSpaceY` temporarily increased by `getBubbleHeight()` for the `createOrgChartElement` call only. Their up-tick grows by one bubble, so the card tops line up with the cards that already have a bubble in the tick. The global is restored before the stub and before the row advances.

The next connector is `inputWhiteSpaceY + getAlignedRowHeight` below. That height is `card + inputWhiteSpaceY`, plus one bubble height if anyone on the row has two managers, plus one comment height if anyone on the row has a termination date.

After all rows, dotted stubs are brought to front. Then the trunk: a drop from `(originX, originY)` by `OffsetY`, and a vertical at `originX` from the elbow down to the connector Y of the last row of that style. There is no left elbow; the per-person stubs are the only horizontals.

### Full tree — `makeFullTree` (`createBranchNNND`)

This is what “Subordinates” and the auto slides actually build.

`OffsetY` (the drop from the parent down to the bar) is `inputWhiteSpaceY` when `getNumberOfBranches >= 2`, plus the comment height when the parent is terminated. When there is only one branch, `OffsetY = 0`. The single branch brings its own top tick.

Two corner cases add `getCommentBoxHeight()` to `OffsetCornerCaseY` (a procedure-local `Variant`, so a recursive call has its own copy; `Empty` adds as 0):

- exactly one branch, and that branch is one manager, and the parent has a termination date
- exactly one branch, and that branch is one worker-department containing one worker, and the parent has a termination date

Department branches come first. The live direct-report collection is **replaced** with one department’s people, `makeTree(..., "Subordinates", "branch_auto")` is called at X = 0, then the original list is restored. `branch_clever` passes `considerAllDirectReportsAsWorkers = True`, so departments are taken from all reports and the manager loop is given an empty list.

If the whole tree needs a bubble alignment (`checkSolidLineDirectReportsNeedShiftDown`: any one-person worker department has two managers, or any manager-report has two managers) and this department is a single person who does **not** have two managers, that branch is shifted down by `getBubbleHeight()` and its reference tick is stretched up by the same amount so the apex stays put.

Manager branches come second. A manager is drawn as a full recursive tree when `ManagerType(parent) = DefaultManagerType` (that is `"Solid"`) **and** the visibility flag is 1. The child is called with regime `"Full"` and a new all-visible dictionary, at

```
originY + OffsetY + (cardHeight + inputWhiteSpaceY) + alignment + OffsetCornerCaseY
```

`alignment` is `getBubbleHeight()` for every manager branch when the shift flag is set. Otherwise the manager is one up-tick card. The Y of that card includes `OffsetDueToDifferentOriginY`, which this procedure assigns only later, and only when its own regime is `"Full"`. On the way down, that variable has not been assigned yet.

Regime `"Full"` then places the head: direction 1, X = `horizontalConnectorCenterX`, Y = `originY - sizeOfOrgChartElement(person, 1).tHeight`, with the comment height added back when the parent is terminated. The head and the subordinate group are grouped as `personFullName & "Full"`.

Branches are reordered so equal `DepartmentKey` values sit together, keeping the first-seen department order (worker departments were created first, so a worker column of a department is followed by the managers of that department).

They are packed from X = 0. The gap between boxes is `inputWhiteSpaceExt`.

For each box, the code reads `GroupItems(tickName).Left` after the move. Names containing `TickSolid` update the solid left/right. Names containing `TickDotted` update the dotted left/right. `SubordinatesTickMainSolid` does not contain the substring `TickSolid`; those main ticks are drawn after this loop anyway.

The center is the midpoint of the outermost solid-or-dotted tick left-edges. A horizontal is drawn at `Y = originY + OffsetY` from the outermost tick of that style to the center, and a main tick goes **up** from the center of the bar by `-OffsetY` back to `originY`. The connector is skipped when the only tick of that style is already on the center, or when the sentinels were never replaced (`Abs(left) + Abs(right) < 1000000` fails for the initial ±1000000).

`sizeOfFullTree` is the same structure, with one difference that matters: the draw test compares `ManagerType` with `DefaultManagerType`, while the size test compares `ManagerType` with the literal `"Solid"`. The commented line in the size function is the `DefaultManagerType` comparison. `OffsetDueToHeadY` is computed in `sizeOfFullTree` and never added there.

### Auto-fit — `makeFullTreeAuto` (`createBranchNNNDAuto`)

Only runs when the measured area is greater than 0.

Visibility:

- `SubordinatesNumber < 100`: start with everyone visible, then set flag 0 on every manager whose `ShouldBeShownUnderManager(parent)` is false (not the solid manager).
- 100 or more: start with everyone hidden.

The saved size is replaced with `Small` and all three whitespaces set to 5.

1. **Slides.** Under 100 people, while a visible manager remains and the tree is wider than 920 or taller than 350, hide the visible manager with the largest `SubordinatesNumber`. At 100 or more, reveal the smallest hidden solid-line manager while the tree still fits; if revealing them overflows, set the flag back to 0. That manager is still eligible on the next pass.
2. **Card size.** Try Small, Medium, Large (dictionary indexes 2, 3, 4). The first size that overflows reverts to the previous size.
3. **Tick.** Measure one step of `+1/3` on X and `+1` on Y. `nDMin` is the smaller of the two “steps until the guide”, rounded, minus 1. Apply that many steps. Caps: `inputWhiteSpaceX` at `cardWidth/2`, `inputWhiteSpaceY` at `cardHeight/3`. The height division is not guarded when the two heights are equal.
4. **Gap between branches.** `inputWhiteSpaceExt` grows by `Abs(treeWidth - 920) / (branchCount - 1)` when there is more than one recursive branch. The `Abs` adds gap whether the tree is narrower or wider than 920. Cap is `getCardWidth()`. The tree is not measured again after this.

`shrinkedTreeFlag` is true when no manager tree is visible, there is at least one manager, and the horizontal branch count exceeds 7. The draw is then `makeTree(..., SlideWidth/2, 140, "Subordinates", "branch_4")`. Otherwise `makeTree(..., 0, 140, "Subordinates", "full_tree", flags)`, and the finished group is centered: `Left = (SlideWidth - Width) / 2`. If there are no managers and the unsorted report count is `< 6`, the group is shifted so the first reference tick’s left edge sits on the old center. The comment in the file says the 1D tick is not in the middle of the group.

Every name from `getReferenceTickNames` is then hidden (`Visible = msoFalse`), and the group moves up by the height of the first of those ticks. `makeHeadOfTree` draws the head at `(SlideWidth / 2, 140)`. The hidden tick plus the upward shift is what makes the head’s 15-point tick meet the bar.

If `infinityRegimeFlag` is true, each direct report with flag 0 who `IsManager` and `ShouldBeShownUnderManager` gets `makeFullTreeAuto` on a new slide. `stopExecutionFlag` (set by the progress bar’s Cancel) leaves the loop before the saved card size and whitespace are restored.

### `makeTree` (`createBranch_Controller`)

| `treeType` | Draw |
| --- | --- |
| `branch_1` | `make1DSubordinates` |
| `branch_2`, `branch_4` | `makeNDSubordinates` with `"2"` or `"4"` |
| `branch_auto`, count = 1 | one up-tick card |
| `branch_auto`, count ≤ 6 | 1D |
| `branch_auto`, count ≤ 14 | ND, 2 per row |
| `branch_auto`, count > 14 | ND, 4 per row |
| `branch_clever`, `full_tree` | if no manager is visible and the branch count exceeds 7: ND with 4 per row. Otherwise `makeFullTree` in regime `"Subordinates"`. `branch_clever` passes `considerAllDirectReportsAsWorkers = True`. |

After `makeFullTree`, `removeComplexGrouping` ungroups nested groups whose names do not end in `Solid` or `Dotted`, then regroups under the same name. `.Left` is set so the first reference tick’s left edge is on `originX`.

Regime `"Full"` on `makeTree` itself (not the inner `makeFullTree` call) adds an up-tick head at the origin and, when more than one shape was created, shifts the shape named `personFullName & "Subordinates"` up by the direction-1 element height, plus another comment height when the parent is terminated. Direction-1 height already includes the comment. The result is grouped as `personFullName & "Full"`.

`sizeOfTree` uses the same thresholds, written as the literals `6` and `14` rather than the constants. It adds `inputWhiteSpaceY` (and, in regime `"Subordinates"`, the comment height) except for `branch_clever`, `full_tree`, and `branch_auto` with one sorted report.

After drawing, actual width and height are rounded to 1 decimal and compared with `sizeOfTree`. A gap larger than 25 shows a `MsgBox` with both pairs of numbers. It does not stop the draw.

### The head on an auto slide — `makeHeadOfTree` (`createBranch0D`)

Forces `inputBoxSize = "Large"` and `inputWhiteSpaceY = 15`, then restores both.

Places a direction-3 card so the down-tick apex is at `(SlideWidth / 2, 140)`. Ungroups and deletes `TickSolid` and `TickDotted`. Draws a new vertical from `(SlideWidth / 2, 139)` with length `-15` (up to Y = 124) for each of solid and dotted that has at least one direct report. The card, and the comment box when there is a termination date, are brought to front, and the pieces are grouped again under `personFullName & DefaultManagerType`.

### Reference tick names (`getReferenceTickNames`)

The prefix depends on `nOrdinaryBranches + nDepartmentBranches`:

- 2 or more, and no visible manager tree, and there is at least one manager, and the horizontal limit is exceeded: `person & "SubordinatesTick"`.
- 2 or more, otherwise: `person & "SubordinatesTickMain"`.
- exactly one department branch and more than one report: `person & "SubordinatesTick"`.
- exactly one department branch and one report: `thatReport & "Tick"`.
- exactly one manager branch whose first visibility flag is 1: `firstReport & "Tick"`.
- exactly one manager branch otherwise: `person & "SubordinatesTick"`.
- 0: a `MsgBox` saying the tree is not shown by default. No prefix is assigned.

`"Solid"` is appended when there is a solid report. `"Dotted"` is appended when there is a dotted report. The dotted count is copied to a variable before the `If`; the comment says that avoids an `If` that runs unconditionally.

---

## 5. Departments, links, title

### Title and statistics (`create_Decorations`)

`addDescriptionUtility` copies `DepartmentGroup`. When a person is passed, it writes:

- `DepartmentManagerFullName` = `FullName`
- `DepartmentName` = `Department`
- on `DepartmentStatistics`, `customReplace` of `<pname>` with `.Name` (the given name, not the full name), `<ndirect>` with the unsorted direct-report count, `<nall>`, `<nordinary>`, `<nimportant>`, `<noutsource>` with the headcount fields

The sentences themselves are the text of shape `DepartmentStatistics` on the palette slide, not VBA strings:

- `<pname> has <ndirect> direct reports`
- `Team Size – <nall> people` (the dash is U+2013)
- `<nordinary> ordinary employees`
- `<nimportant> non-ordinary employees`
- `<noutsource> outsourcers`

The legend text (`Indicators`, `— Ordinary employee`, `— Important employee`, `— Outsourcer`) is also template text. This module does not write it. The group is then ungrouped.

`addNavigationUtility` copies `NavigationArrow`, centers it horizontally, sets `Top = 15`, and replaces `<pname>` and `<sliden>`. The template text is `Up to <pname> (slide <sliden>)`. The hyperlink `SubAddress` is `SlideID & "," & slideNumber & "," & Slide.Name`.

`addSlideNUtility` copies `SlideN` + the size that matches the card height. It sits on the card’s bottom, plus `1.25 * commentBoxHeight` when the termination flag is set. If a target slide is passed, the text becomes `"See slide " & n` and the hyperlink points there.

### Shades (`create_Departments`)

`createDepartments` does nothing on the palette slide.

If the selection is exactly one shape whose name contains `"Shade"` and ends with `"Department"`, the procedure lightens that shade toward white (or toward the template shade) and returns. It does not rebuild.

Otherwise it collects every checklist entry whose parsed `ManagerType` is `"Solid"`, builds a `CVirtualShade` from `person & "MainBackgroundSolid"`, and skips a card whose top is not below `const_slideTopBorder` (140) or whose department key is empty.

`eraseAllShapesLike(..., "Department")` returns true when it deleted something. True means this click only erased. False means it builds.

For each department, a full-slide rectangle (top 140, right 945, bottom 505, left 15) is pulled in to the nearest card that `CVirtualShade.Location` reports as top, bottom, left, or right. The shade edge is then the average of the card edge and that neighbor, ±1. Those `±1` values are the only padding. The first pass uses the inverted `"default"` rectangle (top 1000000, right −1000000, bottom −1000000, left 1000000); `insertShapeInside` expands each side independently, so the first card replaces all four sentinels.

The shape is `Shade1Department`, `Shade2Department`, or `Shade3Department` (`shadeIterator` mod 3). Its text frame is the department name. It is sent to back. A second pass reorders shades when one fully contains another (`hasShapeInside`: all four corners).

`CVirtualShade.intersectsWith` is true only when some corner of one rectangle lies inside the other, edges included. It is not a general rectangle overlap test. `Location` overwrites in the order top, right, bottom, left; the last match wins.

### Links (`create_Links`)

First click that finds `SlideNSolid`, `SlideNDotted`, or `NavigationArrow` deletes them and stops. The next click, finding none, creates links.

A checklist key whose collection has count 2, and whose person `IsManager`, gets a pair. Count greater than 2 shows “too many cards… leave maximum two”. A non-manager with two cards is ignored.

```
If bufferSlideList(1) > bufferSlideList(2) Then
    slideNumberToBeLinked = bufferSlideList(2)
    slideNumberToBeLinked = bufferSlideList(1)   ' overwrites the previous line
Else
    slideNumberToBeLinked = bufferSlideList(1)
    slideNumberToLink = bufferSlideList(2)
End If
```

In the `Then` branch `slideNumberToLink` is never assigned. In the `Else` branch the smaller-or-equal index receives “See slide N” and the other receives the “Up to” arrow. The comment says the extended slide always comes after the shrunk slide.

### Align (`buAlignShapes`)

Operates on `ActiveWindow.Selection.ShapeRange`. Align needs more than one shape (`msoAlignTops`, `msoAlignLefts`, `msoAlignCenters`). Distribute needs more than two. The second argument is `msoFalse` (relative to the shapes, not the slide). Group requires more than one shape; otherwise the message is `It's already a Group object. :(`. Ungroup walks the selection and disables the button when no group remains. `alignCenter` has no ribbon button.

`CEventClass` sets the group, ungroup, and “add slide N” flags on `WindowSelectionChange`. The handler reads `ActiveWindow.Selection`, not the `Sel` argument. For a single shape whose parsed name has a `FullNameKey`, it copies that key into `inputPersonFullName`, sets the category to the uppercase first character, and, when sticky sizes are on and the element is a card, maps the `MainBackground` height to a box size (tolerance 5 points) and the tick’s longer side to 5, 7, 10, 12, or 15 (a difference of exactly 1 does not match).

---

## 6. File by file

### `customUI/customUI14.xml` (149 lines)

Ribbon. `startFromScratch="false"`. One tab, `OrgChartCreator`, label `Org chart`. Four groups, in order, with no `label` of their own: `setupStage`, then `createStage`, `decorationStageElements`, and `decorationStageManipulation`. The last three are visible only when `isExcelOK` (`fileUploadingSuccessful`).

`setupStage`: Select Excel file (`pickExcelFile`, hidden once a file is accepted), File / Sheet / Columns have titles, Import data (`uploadExcelFile`), then the Refresh split (Refresh calls `uploadExcelFile` again; the menu has Select another file and Select another sheet).

`createStage`: category, employee, Create org chart (Single slide org structure, Full org structure), Element size, Tick size, Sticky sizes, Manager / Employee, Direct reports / 1-in-row / 2-in-row / 4-in-row, Subordinates, Departments, Links.

`decorationStageElements`: Slide, Elements (Title, Shade, Frame, Card, Connector, Link).

`decorationStageManipulation`: Ungroup, Group, Arrange (Distribute Horizontally, Align Top, Distribute Vertically, Align Left).

`checkPeople` and `alignCenter` are not in this file.

### `utilInitialize` (205 lines)

API declares (`Sleep`, window properties, `CopyMemory`) and every public flag listed in section 1. `getRibbon` returns `RibbonObj()`. `MyAddInInitialize` is the ribbon `onLoad`. `stopExecutionFlag = False` inside it is commented out, so load does not clear that flag. `Dim X As New CEventClass` lives for the session.

### `utilInitializeRibbon` (13 lines)

Holds `m_oRibbon`. `RibbonOnLoad` stores it. Dropdowns are not filled here; their callbacks are in `setupStage3SelectTreeParameters`.

### `setupStage1PickExcelParameters` (85 lines)

Declares `mCurrentSheet`, `inputSheetIndex`, `inputSheetName`, `inputFirstRow`. Ribbon callbacks for the sheet dropdown (`getNumberOfSheets`, `getSheetName`, `getSheetIndex`, `uiGetSheetId`, `uiOnChangeSheet`) and the titles checkbox. Sheet ids are `"sheetIndex" & (index + 1)` — the ribbon index is 0-based, the id is 1-based. No worksheet is read here.

### `setupStage2LoadInputFile` (359 lines)

The file dialog, the sheet list, `uploadExcelFile`, and the stage flags `fileSelectionSuccessful` / `fileUploadingSuccessful`. Error labels: `askToSelectAnotherFile` (the message uses `inputFileNmae`, which is not declared or assigned anywhere, so the name in that sentence is empty), `askToSelectAnotherSheet` (Mac sheet check only; a bad Windows sheet name raises inside `Worksheets(inputSheetName)` and lands on the file message because `On Error GoTo askToSelectAnotherFile` is still active), `askToCheckTheData` (not enough columns).

### `setupStage3SelectTreeParameters` (253 lines)

Category and employee dropdowns, card size, tick length, sticky-sizes checkbox.

`getEmployeesListByCategory`:

- `"CEO"` → one entry, `inputCEOFullName`
- `"Managers"` → keys whose `IsManager` is true, in the alphabetically sorted list order
- `"AZ"` → the sorted list, capped at 1000
- any other string, the letter from the id → keys whose first character equals `LCase(category)`

Selecting `categoryAZ` always shows the 1000-person message. The cap itself is in `getEmployeesListByCategory`, and this `MsgBox` does not test the count.

`GetSelectedBoxSizeID` defaults to `boxSizeMedium`, so `inputBoxSize` becomes `"Medium"`. `GetSelectedWhiteSpaceCoef` defaults to `whiteSpaceCoef10`. Employee ids are `"employee" & index` with the index **not** increased by 1 (unlike sheet ids).

`autodetectSizeAndWhiteSpaceSwitch` toggles the flag and ignores `pressed`. Nothing in this file writes `dropDownBoxSizeEnabledFlag` or `dropDownWhiteSpaceEnabledFlag` after initialization.

### `utilStage1SaveInfoAboutEmployee` (116 lines)

`extractEmployeeInfo` / `fillEmployeeInfo`. An Excel error cell hides the progress form and says which line to check; the sheet piece of that message is `inputSheetIndex`, not the sheet name. A row whose blank-cell count equals 13 returns a new empty `CEmployee` and is later skipped as an empty key. A row with an empty surname or name returns `Nothing`, and `buildHierarchy` stops. The message says “Name of Surname”. The displayed line is `row + inputFirstRow - 1`. On the 1-based Windows array that is the Excel row. On the 0-based Mac array the same formula prints one less.

### `utilStage2BuildHierarchy` (99 lines)

`buildHierarchy`. Clears the three org dictionaries, walks every row, fills `MyDictionaryPersonalInfo`, `MyDictionaryDirectReports`, and `MyDictionaryCeoCandidates`. See section 2.

### `utilStage3CheckConnectionsToCEO` (189 lines)

`checkConnectionsToCEO`, `extractReportingSequence`, `verifyReportingSequence`. See section 2. `fullNameTopLevelManagerKey` is the public result.

### `utilStage4CheckAllPeopleCreated` (173 lines)

`showWhomIForgot` and the checklist scan. A shape counts when `parseShapeName` finds a non-empty person key, a manager type, and `ElementName = "InfoBox"`. The card key is `FullNameKey & ManagerType`. A manager’s second hit increments `numberOfPeopleSplitIntoSlides`; the third and later increment `numberOfPeopleDuplicates`. A non-manager’s second and later hits are duplicates. The duplicate `MsgBox` is inside a comment and does not run. Forgotten people are bucketed by the slide of their manager’s card, or by `"0"` when the manager card is missing or appears more than once. `sortUtility(..., "number", False)` orders the slide buckets. Bucket `"0"` is created unconditionally, so the `MsgBox` runs even when nobody is missing; its header is suppressed because `slideNumber <> 0` fails.

### `utilCheckWordCombinations` (91 lines)

`checkPersonExists` / `synthesizeFullNameCandidates`. Builds `N!` concatenations of the words and returns the first one that `MyDictionaryDirectReports.Exists`. No `MsgBox`.

### `utilCountPeople` (75 lines)

`CalculateSubordinates`, `refreshSubordinatesChecklist`. See section 2.

### `utilSortOrgElements` (74 lines)

`compare` / `sortUtility` / `makeLegibleAfterSort`.

- `"alphabet"`: `UCase(p1) < UCase(p2)`. `inverseOrderFlag` negates, including a tie (equal keys become true).
- `"number"`: the raw `Variant` `<`, then the same negation. It does not call `CLng`.
- `"maturity"`: larger `SubordinatesNumber` makes `compare` true. The inverse flag flips that, but a tie overwrites the result with the uppercase-key comparison, so the inverse flag does not apply to a tie. The stored number is `CInt`’d. `-1` (the “not calculated” sentinel) sorts as a small integer.
- Any other regime leaves the Boolean false.

`sortUtility` is an in-place quicksort. The pivot is copied from the middle index before partitioning. `makeLegibleAfterSort` replaces a key with `FullName`, or with `Name` alone when `IsToBeHired`.

### `CEmployee` (410 lines)

Section 2. No `Class_Initialize`. `HeadcountStats` is `Property Let` (not `Set`) and the body uses `Set`. The getter is typed `Object`.

### `CSubordinatesStatistics` (112 lines)

Six integer counters, `ResetToZero`, `ResetToDefault`, `IsDefault`, `add`, `ByType`. A new instance is all zeros, and `IsDefault` is false for zeros.

### `CTreeSize` (16 lines)

`tWidth` and `tHeight`, both `Double`, starting at 0.

### `CVirtualShade` (233 lines)

Axis-aligned rectangle, Y downward. `dimensionsVBAStyle(x, y, width, height)` stores top = y, left = x, right = x + width, bottom = y + height. `Width` and `Height` are declared `As Integer` and return a `Double` subtraction. `visualizeColorful` drops a random-colored rectangle on the slide; nothing in the draw path calls it.

### `CTimer` (63 lines)

Mac: `Timer` deltas in milliseconds. Windows: `QueryPerformanceCounter`. `TimeElapsed` samples the counter on every read.

### `CEventClass` (85 lines)

Selection change. See the align section. Errors before the inner handler, and the inner name-parse failure, exit without invalidating the ribbon. The outer comment says errors are ignored for Mac compatibility.

### `Dictionary` (454 lines)

Tim Hall’s VBA-Dictionary 1.4.1, compiled as the custom class on both platforms (`UseScriptingDictionaryIfAvailable = False`). Not `Scripting.Dictionary`.

- Default compare mode is binary.
- `Exists` is false for a missing key and does not add it. A stored `Empty` value still exists, because the entry is a three-element array.
- `Item` get of a missing key returns `Empty` and does not add the key. `Item` let/set of a missing key adds it.
- `add` of an existing key raises 457. `Remove` of a missing key raises 32811.
- `keys` and `Items` are 0-based insertion order. An empty dictionary returns `Split("")`.
- String keys under binary compare are stored with a case mask (`abc` → `abc__abc`, `ABC` → `ABC_____`) because `Collection` keys are not case-sensitive.
- `RemoveAll` does not clear the side list of object keys.

### `buCreateCard` (323 lines)

Section 3.

### `buCreateConnector` (40 lines)

Section 3.

### `TreeManagement_GetDimensions` (97 lines)

`mapBoxSizeToHeight` returns the first card size whose Ordinary height is within 5 points, otherwise `MyDictionaryCardSizes(1)` (`ExtraSmall`, the first item added). The other getters are in section 3. No card point-size literal lives in this file.

### `TreeManagement_SizeVisibility` (192 lines)

`getNumberOfBranches`: `"ordinary"` manager count, `"department"` worker-department count, `"all"` the sum, `"allrecursive"` worker departments plus, for each manager, a recursive call when the flag is 1 and 1 otherwise. Any other regime returns 0.

`getReferenceTickNames`: section 4.

`getAlignedRowHeight`: section 4.

`synthesizeInputAllDirectReportsVisible` / `Invisible` / `FixedVisibility`: a new dictionary, every unsorted direct report set to 1 or 0. The person themselves is not an entry.

`countVisibleTrees` / `countInvisibleTrees`: entries whose value is 1 or 0, the key is not `""`, and `IsManager` is true.

`checkSolidLineDirectReportsNeedShiftDown`: section 4.

### `VBA_CopyShape` (95 lines)

`copyShapeErrorFree` copies a palette shape through the clipboard and pastes it. A name collision appends a random integer from 1 to 100. Mac waits `Timer + 0.5`. Windows polls `GetClipboardSequenceNumber` every 50 ms, up to 200 times. Paste is retried up to 10 times. Failure shows “Something seriously went wrong during copy process…” and executes `End`. The pasted shape is not moved.

### `VBA_Group` (68 lines)

`groupAllShapesInIndexList` groups shapes `(firstIndex + 1)` through `lastIndex`. Callers pass the count before the new shapes and the count after, so the range is exactly the shapes just added. One shape is left with its own name (the rename is commented out). Two or more are grouped under `groupName`.

`removeComplexGrouping` ungroups nested groups whose names do not end in `Solid` or `Dotted`, then regroups the span and sends it to back.

### `VBA_EraseShapes` (187 lines)

`eraseAllShapesLike` deletes every shape, including inside groups, whose name ends with the tag. The tag `"<TEXTBOX>"` matches `msoTextBox` instead. `eraseShape` deletes every shape of that exact name, repeating the lookup. `eraseCardAttributes` / `eraseManagerAttributes` / `eraseSingleElement` / `eraseSubordinates` delete by constructed name. `eraseSubordinates` keeps a subordinate when `extractReportingSequence` contains the manager. The root card is not deleted there.

### `VBA_Utilities` (41 lines)

`wfRandomIntegerInRange`, `wfFactorial` (returns 1 for `n <= 0`; the return type is `Integer`), `wfTransposeArray` (0-based result; a 1-based source is indexed from 0), `cce` (counts a collection by walking it), `checkClassToCodeCommunication` (a `MsgBox` of `inputPersonFullName` and a ribbon invalidate).

### `VBA_RegularExpression` (104 lines)

No regular-expression engine. `textInBraces`, `textInSquareBrackets`, `capitalLetters` (A–Z only), `isNumeric` (digits only; a sign or a dot fails), `startsWith`, `endsWith`, `multReplace`.

`parseShapeName` walks capital letters and splits the name into blocks. The first block is `FullNameKey` only when `MyDictionaryPersonalInfo.Exists` says so. A block equal to `Solid` or `Dotted` is `ManagerType`. Every other block is appended to `ElementName`. `extractEmployeeName`, `extractManagerType`, `extractElementName` are one-field wrappers.

### `VBA_ReplacePreserveFormatting` (12 lines)

`customReplace` finds the first hit, inserts the new text before it, and deletes the old hit. One replacement. Used for the statistics tokens and the navigation tokens.

### `createBranch0D` (47 lines)

`makeHeadOfTree`. Section 4.

### `createBranch1D` (82 lines)

`make1DSubordinates`, `sizeOf1DSubordinates`. Section 4.

### `createBranchND` (156 lines)

`createHorizontalShiftsMap`, `makeNDSubordinates`, `sizeOfNDSubordinates`. Section 4. When the report count is at most half the row width, `sizeOfNDSubordinates` sets `hLineRightAbs = -getCardWidth() / 2` and adds that negative number into the width.

### `createBranchNNND` (489 lines)

`makeFullTree`, `sizeOfFullTree`. Section 4.

### `createBranchNNNDAuto` (236 lines)

`makeFullTreeAuto`. Section 4. The fourth priority is spelled `FORTH PRIORITY` in the source.

### `createBranch_Controller` (172 lines)

`makeTree`, `sizeOfTree`. Section 4.

### `createTree_Controller` (243 lines)

`mainSwitch` and the ten ribbon wrappers, including `checkPeople`, which is not on the ribbon. Section 1. If `inputPersonFullName` is empty the message is “Employee's full name undefined. Might be you clicked on a person from another Excel sheet.” No slide selected raises “Please select the slide before creating orgtree.” More than one selected shape, or a shape that is neither the person group nor the card group, stops with a message. The anchor taken from a selected card is the `MainBackground` bottom-center; if the tick hangs below that bottom, `posY` becomes the tick’s bottom.

### `create_Decorations` (163 lines)

Empty slides (`"insert"` is the current index + 1, `"add"` is the slide count + 1, using slide 1’s custom layout), navigation, slide-N badges, the title group, the attention frame, one manual shade, four empty cards, and a random horizontal named `connectorRandom` plus an integer from 1 to 100, drawn at (100, 100) with length 100.

### `create_Departments` (294 lines)

Section 5.

### `create_Links` (71 lines)

Section 5.

### `buAlignShapes` (83 lines)

Section 5.

### `UI_UpdateProgressForm` (120 lines)

Two modeless forms.

- Digit (`openProgressDigit`, `updateTaskDigit`, `updateProgressDigit`, `closeProgressDigit`): the caption is `Round(percent * 100) & "%"`. A negative percent shows “This may take a while”. Values above 1 are not clamped.
- Bar (`openProgressBar`, `updateTaskBar`, `updateProgressBar`, `updateCancelButtonVisibility`, `closeProgressBar`): percent above 1 is clamped to 1. The caption is `"N% completed"`. The bar width is `percent * 200`. Negative percent clears the bar. `cancelProgress` sets `stopExecutionFlag = True` and hides the form. `updateProgress` is only `DoEvents`; `UFProgressBar_Activate` calls it.

Both are centered on `Application` plus half the active window.

### `Forms/UFProgressBar` (6 lines)

`Activate` → `updateProgress`. `CancelButton_Click` → `cancelProgress`.

### `Forms/UFProgressDigit` (9 lines)

`statusLabel_Click` and `UProgressBar_Click` contain only the comment `'Unload Me`. A click does nothing.

### `utilOpenDialog` (53 lines)

`browseFile`, described above. `CheckAppleScriptTaskPowerPointScriptFile` looks for the script under `~/Library/Application Scripts/com.microsoft.PowerPoint/` by taking the desktop POSIX path, deleting `"/Desktop"`, and appending that library path. A missing file produces “Sorry, the MyBoss-browse_files_on_mac.scpt is not in the correct location”.

### `utilTestTemplate` (83 lines)

`initPalette`, `verifyPaletteOnSlide`, `findFullPalette`, `findPaletteCandidateForTemplate`. `initPalette` is declared `As Boolean` and never assigns the return value. `findFullPalette` sets `paletteSlideNumber` only when a slide contains every template; otherwise it leaves the variable as it was.

### `util_Macros_EraseTextBoxes` (3 lines)

`eraseTextBoxesOnSlide` calls `eraseAllShapesLike` on the selected slide with `"<TEXTBOX>"`.

### `util_Macros_ExamineCostCenters` (132 lines)

`showCostCenters` calls `hideCostCenters`, refreshes the card checklist, and adds a green text box (`RGB(0, 255, 0)`, 50×30, font 10 bold, centered) at the card’s bottom-right (`Left + Width - 50`, `Top + Height - 20`). The text is `CostCenter`, or `"?"` when it is blank or `"0"`. The shape is named `ccTextBox` plus a counter. `hideCostCenters` deletes every shape whose name contains `ccTextBox`, repeating the scan until none remain.

### `MyBoss-browse_files_on_mac.scpt`

Compiled AppleScript (`FasdUAS 1.101.10`), not source text. A byte-identical copy is `misc/MyBoss-browse_files_on_mac.scpt`. Five handlers:

| Handler | What the bytecode does |
| --- | --- |
| `openDialog(seedPath)` | `choose file` with prompt “Please select a file”, default location `seedPath`, multiple selection off. Returns the alias as text. On error returns `errorNumber`. |
| `openExcelGetSheets(filePath)` | Tells Microsoft Excel to open the file, hides the Excel process, appends each sheet name plus a comma, saves and closes, returns `filePath & "|{" & names & "}"`. The trailing comma is why VBA drops the last character. The strings `xls` / `xlsx` / `xlsm` appear in stored source text and are not tested by the bytecode. |
| `openExcelGetSheetData(filePath\|sheet\|offset\|count)` | Opens the workbook, reads the used range, appends each row whose counter is past `offset` and within `offset + count`, separated by `;`. |
| `checkExcelExists(filePath)` | Opens, saves, closes. Returns `"success"` or `errorNumber`. |
| `checkSheetExists(filePath\|sheet)` | Same, and looks up the sheet by name on the opened workbook. |

### `VBA_MAC_COMPATIBILITY.md` (459 lines)

A note dated January 2026 about Windows-versus-Mac VBA. It is a claim about Office, not a trace of this add-in running. It contradicts itself on whether a `Declare PtrSafe` at module level crashes on load. The Mac branches that this add-in actually compiles are the `#If Mac` blocks in `utilOpenDialog`, `setupStage2LoadInputFile`, `VBA_CopyShape`, and `CTimer`.

---

## 7. Easy to misread

These are in the source as written. They are listed so a later port does not “correct” them silently.

1. `MAX_PEOPLE_FOR_2D_BRANCH` is 6 and selects the **vertical** stack (`make1DSubordinates`), not a 2-column layout. Two columns start at 7 and run through 14.
2. `DefaultManagerType` is the literal `"Solid"`. `FullNameManagerKey` still returns the dotted manager when the person has one.
3. Headcount includes the person. `TBH_` is excluded from the five person counters and included in `SubordinatesCardsInFullTree`. Dotted subtrees are not walked.
4. The 1D trunk is to the **left** of the cards (`originX - cardWidth/2 - inputWhiteSpaceX`) and the tick enters the left edge at mid-height (direction 4). It is not a line through the card centers.
5. On an auto slide the reference ticks are hidden and the group is shifted up by that tick’s height. The visible drop under the head is the 15-point tick from `makeHeadOfTree`.
6. Solid connectors are `RGB(91, 155, 213)`. Dotted connectors are `RGB(0, 32, 96)` with `msoLineSysDot`.
7. `create_Links` assigns `slideNumberToBeLinked` twice in the branch where the first slide index is greater, and never assigns `slideNumberToLink` there.
8. The Mac importer’s column test is `UBound(eCells) >= 13`. A 0-based 13-cell row has `UBound` 12.
9. `sizeOfFullTree` tests `ManagerType = "Solid"`. `makeFullTree` tests `ManagerType = DefaultManagerType`. The commented line in the size function is the draw’s comparison.
10. `And` evaluates both sides. `verifyReportingSequence` is called for solid and dotted even when the first one already failed.
11. Statistics sentences, the legend, and “Up to …” are text on the palette slide. The VBA only replaces tokens.
12. Card point sizes are the palette shape’s width and height. `getLineWeight` is the only size table written as literals (2, 2, 3, 4).
