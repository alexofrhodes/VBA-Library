# one click export/import system

## working on big update, almost ready to share



| Macro  | Description |
| ------------- | ------------- |
| addLinkedLists | Insert in Procedure-to-Export a list of linked procedures, userforms, classes, declarations |
| ExporProcedure  | implements addLinkedLists and exports recursively the LinkedProcedures and other elements to folders <br /> (to be synced on github - manually atm, todo)  |
| ImportProcedure  | Import target procedure and its linked elements from local folder or download it from github | 
| Update | implements ImportProcedure |
|GetMissingDependencies | implements ImportProcedure |
