# Hold Form — In-Contract Student Handling

## Goal

Add support for students who are still within their initial enrollment contract. These students continue to be billed during a hold rather than having billing paused. The email and web app acknowledgment page must reflect this with different wording so parents understand they'll be charged during the hold and credited later.

## Background

When a student goes on hold, billing is normally paused. But if the student is still within their initial enrollment contract, Mathnasium continues charging and applies a credit at the end of the contract term. The current hold workflow has no way to distinguish these two cases — both get the same "billing paused" language, which is incorrect for in-contract students.

## New INPUT Fields

Three new columns are inserted into the CD-editable zone of the INPUT tab, before "Send Email?":

| Col | Field | Type | Notes |
|---|---|---|---|
| H | In Contract? | YES/NO dropdown | Required when sending |
| I | Charged Month(s) | Free text | e.g. "June and July" — only used when In Contract? = YES |
| J | Credit Month(s) | Free text | e.g. "August and September" — only used when In Contract? = YES |
| K | Send Email? | YES/NO dropdown | Shifted from H |
| L | Radius Updated | YES/NO dropdown | Shifted from I |
| M–U | Auto-fill (greyed) | — | All shifted 3 columns right |

Charged Month(s) and Credit Month(s) are only used when In Contract? = YES. If In Contract? = NO, those fields are ignored even if populated.

## Column Index Changes

All IC (INPUT) and MC (MASTER_HOLDS) indices shift to accommodate 3 new fields:

### IC (INPUT tab, 0-based)
```
STUDENT_NAME:    0  // A
PARENT_NAME:     1  // B
PARENT_EMAIL:    2  // C
HOLD_START:      3  // D
EXPECTED_RETURN: 4  // E
BILLING_RESUME:  5  // F
NOTES:           6  // G
IN_CONTRACT:     7  // H  ← NEW
CHARGED_MONTHS:  8  // I  ← NEW
CREDIT_MONTHS:   9  // J  ← NEW
SEND_EMAIL:      10 // K  (was 7)
RADIUS_UPDATED:  11 // L  (was 8)
// auto-fill (greyed) ─────────────
STATUS:          12 // M  (was 9)
FORM_LINK:       13 // N  (was 10)
EMAIL_SENT_TS:   14 // O  (was 11)
SIGNED:          15 // P  (was 12)
SIGNATURE_TS:    16 // Q  (was 13)
REMINDER_1:      17 // R  (was 14)
REMINDER_2:      18 // S  (was 15)
REMINDER_3:      19 // T  (was 16)
HOLD_ID:         20 // U  (was 17)
```

### MC (MASTER_HOLDS tab, 0-based)
Adds IN_CONTRACT, CHARGED_MONTHS, CREDIT_MONTHS after NOTES (index 10):
```
HOLD_ID:          0
CENTER:           1
SOURCE_TAB:       2
SOURCE_ROW:       3
STUDENT_NAME:     4
PARENT_NAME:      5
PARENT_EMAIL:     6
HOLD_START:       7
EXPECTED_RETURN:  8
BILLING_RESUME:   9
NOTES:            10
IN_CONTRACT:      11  ← NEW
CHARGED_MONTHS:   12  ← NEW
CREDIT_MONTHS:    13  ← NEW
FORM_LINK:        14  (was 11)
EMAIL_SENT:       15  (was 12)
EMAIL_SENT_TS:    16  (was 13)
SIGNED:           17  (was 14)
SIGNATURE_TS:     18  (was 15)
REMINDER_1:       19  (was 16)
REMINDER_2:       20  (was 17)
REMINDER_3:       21  (was 18)
RADIUS_UPDATED:   22  (was 19)
STATUS:           23  (was 20)
```

## Email Behavior

### In Contract? = NO (standard hold)
No change. Email body shows "Enrollment Billing Resumes" row as today.

### In Contract? = YES
The "Enrollment Billing Resumes" row is hidden from the details table. In its place, a billing notice paragraph appears:

> As **[Student Name]** is currently within their enrollment contract, monthly billing will continue during the hold period. They will be charged for **[Charged Month(s)]**, and a corresponding credit will be applied to **[Credit Month(s)]** at the end of the initial contract.

Email subject line is unchanged.

## Web App Acknowledgment Page Behavior

### In Contract? = NO (standard hold)
No change. Details table shows "Enrollment Billing Resumes" row. Checkbox text is the existing standard language.

### In Contract? = YES
The "Enrollment Billing Resumes" row is hidden. The billing notice paragraph (same as email) appears in the details section. The acknowledgment checkbox text changes to:

> I understand that **[Student Name]**'s monthly enrollment fee will continue during the hold period, and that a credit will be applied to **[Credit Month(s)]** at the end of the initial contract.

## Reminder Emails

No change to reminder timing or structure. Reminders inherit the in-contract flag from MASTER_HOLDS and render the same in-contract wording as the initial email if applicable.

## Validation

- In Contract? is required (YES/NO dropdown). If blank and Send Email? = YES, row is skipped with an error logged.
- Charged Month(s) and Credit Month(s) are required when In Contract? = YES. If either is blank, row is skipped with an error logged.
- If In Contract? = NO, Charged Month(s) and Credit Month(s) are ignored.

## Dashboard

No changes to the dashboard. Contract status is stored in MASTER_HOLDS but not surfaced in dashboard views (out of scope for this change).

## Files Changed

- `/hold-form/Code.gs` — all changes contained here:
  - Section 0: IC and MC constant indices updated
  - Section 3: INPUT tab headers, column formatting, dropdown validation, CD-editable zone (A–L), auto-fill zone (M–U)
  - Section 4: MASTER_HOLDS headers updated
  - Section 5: `processDirectorInputs()` — read 3 new fields, validate, pass in holdData
  - Section 6: `getInitialEmailBody()` — branch on `holdData.inContract`
  - Section 9: `doGet()` — pass inContract, chargedMonths, creditMonths to template
  - Section 10: `getAcknowledgmentPageHtml()` — branch on inContract for details row and checkbox text
  - Section 11: Reminder functions — pass through in-contract fields from MASTER_HOLDS
