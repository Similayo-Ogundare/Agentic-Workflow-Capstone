# Instructions.

## 1. Goal
Review the tithe and offering spreadsheets for the January to April review period and produce monthly summaries and a final cumulative summary for all 19 members.

## 2. Source Files
Use all relevant spreadsheet files contained in the January, February, March, and April folders.

Ignore the `church.eml` file.

Do not modify the original spreadsheets.

## 3. Monthly Review Process
For each month:
1. Review every sunday spreadsheet in that month's folder.
2. Identify each member's tithe and offering payments.
3. Count the number of tithe payment occurences for each member.
4. Count the number of offerinf payment occurences for each member.
5. Calculate the total amount of tithe paid by each member for that month.
6. Calculate the total amount of offering paid by each member for that month.
7. Include all 19 members, even if a member made no payment during that month.

## 4. Monthly Output
Create one summary spreadsheet for each month.

Each monthly spreadsheet should contain:

- First Name
- Last Name
- Total Tithe Paid
- Total Offering Paid
- Number of Tithe Payments
- Number of Offering Payments

Create:
- January_Summary.xlsx
- February_Summary.xlsx
- March_Summary.xlsx
- April_Summary.xlsx

Save all monthly summaries in the `outputs` folder.

## 5. Final Summary

After completing all four monthly summaries, create one final spreadsheet.

The final spreadsheet should contain:
- First Name
- Last Name
- Total Tithe Paid Across the Review Period
- Total Offering Paid Across the Review Period
- Number of Tithe Payments
- Tithe Consistency Rank
- Number of Offering Payments
- Offering Consistency Rank

Save the file as:
`Final_Summary.xlsx`

## 6. Consistency Ranking
Tithe consistency is based on how many times each member made a tithe payment during the review period.
Offering consistency is based on how many times each member made an offering payment during the review period.
A payment occurence is counted once for each Sunday on which a payment is recorded, regardless of the amount paid.
Do not use the total monetary amount as the primary basis for consistency.
Rank tithe consistency separately from offering consistency.
The member with the highest number of payment occurences should receive Rank 1, then continue in descending order until all 19 members are ranked.

## Tie-Breaking Rule
If two or more members have the same number of tithe payment occurences, the member with the higher total tithe amount paid during the review period should receive the higher tithe consistency rank.

If two or more members have the same number of offering payment occurences, the member with the higher total offering amount paid during the review period should receive the higher offering rank.

The payment should only be used as a tie-breaker when the number of payment occurences is equal.

## Added One More Lever (Instrutions) For The Second Prompt
If two or more members have the same consistency rank, break the tie using the combined total amount of tithe and offering paid by each tied member during the review period. The member with the higher combined total should receive the higher rank. This tie-breaking rule applies to both tithe consistency ranking and offering consistency ranking. In either case, compare the tied members total tithe plus total offering payments. The member with the higher combined amounts ranks higher.

## Additional Instructions For The Third Run
Add a column titled Average and a separate column titled Overall Rank to the final summary. Calculate it using the average of each member's tithe consistency rank and offering consistency rank. The member with the lowest average rank should receive the highest overall rank.

## Additional Instructions For The Fourth and Final Run
If two or more members have the same average rank, apply the existing tie-breaker rule using the combined total amount of tithe and offering paid.

## 7. Quality Checks.

Before completing the task:

- Confirm that every relevant Sunday spreadsheet has been reviewed.
- Confirm that all 19 members appear in every monthly summary.
- Verify monthly totals against the source spreadsheets.
- Verify that the final totals equals the sum of the monthly totals.
- Verify that payment counts match the number of Sundays on which payments are recorded.
- Do not invent missing values.
- Clearly state any uncertainty or conflicting information.
- Do not modify the original files.

## 8. Output Location

 Save all generated files in the `outputs` folder.
 Do not save generated files inside the source month folders.

 ## Documentation and Audit Trail
 Create a `logs` folder if one does not already exist.
 Inside it, create and maintain a file named `WORK_LOG.md`.
 Update `WORK_LOG.md` as the task progresses.

 The log should record:
 - files reviewed
 - major actions completed
 - assumptions made
 - decisions made
 - unclear or conflicting information
 - corrections made
 - files created
 - issues requiring human review if any.


## 9. Quality Standard
The work is considered complete only when every source file has been reviewed, all calculations and rankings have been checked for accuracy, the same rules have been applied consistently to all members, any uncertain or conflicting information has been clearly documented, and all required outputs have been created in the correct location without modifying the original source files.