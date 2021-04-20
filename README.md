# Automate_Forms
To appoint an employee for the year, three documents must be completed with similar information:
  Excel sheet of their employment history, "History Record", with new appointment overview added
  Word Document with the appointment offer details
  PDF Form with appointment details
The program first reads a list of all assignments for the year by employee and makes some initial determinations about each appointment (start date for example). It then looks for a preexisting History Record for each appointment listed and pulls past data to help calculate additional appointment information. If no History Record exists it uses base information. It then writes the new appointment at the end of the history record and makes calculations/determinations throughout. Certain milestones can be hit throughout the year based on the history. The algorithm catches these and makes adjustments as necessary.

Next steps: Use appointment to fill out Offer Letter and PDF form.

Explainers: In course assignment documents. Under each quarter is the courses for that year and often looks like "# xxx #". The first number is the total instances of the specific course, the set of letters is an abbreviation for the course type, and the third number is the identifying course number. Example: "2 MATH 101" is 2 sessions of "Math 101" 

Limitations: Cannot calculate salary increases since they may very, but does highlight when salary increases will occur and creates a list for the user to go back and adjust. 
