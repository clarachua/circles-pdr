**About this Project**

This project provides Google App Scripts for a company's Performance Management Review.  Using git for versioning, to dos, and bug fixes.


**TO DO**
- [X] Remove var findx and replx, only declare 1 var and +rows from that
- [X] Remove parentFolder and childFolder in each function
- [X] Extract FolderIds and number of files per folders
- [X] Move resigned employees to Archive Folders
- [X] Move employees who have changed supervisors into sup folder
- [X] Set Folder Permissions by FolderId and Email tied to supervisor
- [ ] Dashboard of completion status over time
  - [ ] Capture daily completion status for all departments
  - [ ] Create dashboard status
- [ ] Modularize find and replace text as a separate function that can be called separately
- [ ] Macro to send emails to supervisors
- [ ] Debug ready()
- [ ] Check protectTab works before CopyAllFolders, changeFormulas
- [X] Remove code that isn't required
- [ ] Refactor code to smaller modules (e.g. File/FolderIters)


**Potential Features / Macros**
- [ ] Protect key ranges if possible while allowing editing access (e.g. adding rows)
- [ ] Set OKR weights as percentages
- [ ] Split codes into 2 - One with all the code necessary to copy a template (Q1), one with code to duplicate sheets (Q2-Q4)
- [ ] Consider coding changing evaluation in find and replace to prevent spelling errors?
