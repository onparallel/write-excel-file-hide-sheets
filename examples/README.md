# Examples

Runnable TypeScript scripts. Each one writes its own `.xlsx` next to the script.

| Example                                          | Demonstrates                                                                                              |
| ------------------------------------------------ | --------------------------------------------------------------------------------------------------------- |
| [`hide-second-sheet.ts`](./hide-second-sheet.ts) | Hide a sheet other than the first. The workbook opens on sheet 1; the hidden one is reachable via Unhide. |
| [`hide-first-sheet.ts`](./hide-first-sheet.ts)   | Hide the first sheet. `<workbookView/>` is pointed at the first visible sheet so Excel opens cleanly.     |

## Run

```sh
# Single example
npx tsx examples/hide-second-sheet.ts

# All examples
npm run examples
```
