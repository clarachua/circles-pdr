**About this Project**

This project provides Google App Scripts for a company's Performance Management Review.  Using git for versioning, to dos, and bug fixes.


**TO DO**
- [X] Remove var findx and replx, only declare 1 var and +rows from that
- [X] Remove parentFolder and childFolder in each function to see if it still works
- [X] Extract FolderIds and number of files per folders
- [ ] Move resigned employees to Archive Folders
- [ ] Move employees who have changed supervisors into sup folder
- [ ] Set Folder Permissions by FolderId and Email tied to supervisor
- [ ] Set up macro to send emails to supervisors
- [ ] Check ready() works (it's not working now)
- [ ] Check protectTab works before CopyAllFolders, changeFormulas
- [ ] Remove code that isn't required
- [ ] Refactor repeated code to smaller modules (e.g. File/FolderIters)


**Potential Features / Macros**
- [ ] Protect key ranges if possible while allowing editing access (e.g. adding rows)
- [ ] Set OKR weights as percentages if possible
- [ ] Split codes into 2?  One with all the code necessary to copy a template, one with code to duplicate sheets
- [ ] Consider hard coding changing evaluation in find and replace to prevent spelling errors?
