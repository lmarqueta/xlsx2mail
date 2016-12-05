# xlsx2mail
Parses an Excel (xlsx) spreadsheet and creates a text list with urgent tasks
(those with due date in a week or less) for a particular user, so it can be
sent by email.

Excel format should look like this:
--------------------------------------------------------
| Task | Status | Due | Owner | Effort | Pct | Comment |
--------------------------------------------------------
| ...
