# TimeCalc
TimeCalc is an Excel spreadsheet file that makes extensive use of formulas, VBA macros, and formatting to ease the task of keeping extremely detailed track of time spent on tasks and projects.

At the end of every billing period, TimeCalc gives you the abilities to send invoices on time spent, and receive payment on outstanding billed items.

# Sheets
The workbook starts with 21 sheets, enumerated below, each used for specialized purposes.

  > [Config](#Config) &bull; [Work](#Work) &bull; [Timesheet](#Timesheet) &bull; [01](#Sheets-01-to-12) &bull; [02](#Sheets-01-to-12) &bull; [03](#Sheets-01-to-02) &bull; [04](#Sheets-01-to-12) &bull; [05](#Sheets-01-to-12) &bull; [06](#Sheets-01-to-12) &bull; [07](#Sheets-01-to-12) &bull; [08](#Sheets-01-to-12) &bull; [09](#Sheets-01-to-12) &bull; [10](#Sheets-01-to-12) &bull; [11](#Sheets-01-to-12) &bull; [12](#Sheets-01-to-12) &bull; [Summary](#Summary) &bull; [LogReport](#LogReport) &bull; [ServiceInvoice](#ServiceInvoice) &bull; [Contacts](#Contacts) &bull; [Services](#Services) &bull; [Training](#Training)

Following are descriptions of each of the sheets and their contents.
- ## Config
  All of the settings of the worksheet are configured in this sheet. The following settings are currently supported.
  - **Invoice Name**. Name of the party to print on the invoice. This will be the name of your business, if doing business under a company name. If you are an individual freelancer, this will be your own name.
  - **Invoice Address**. Address to which payment and correspondence should be sent.
  - **Hourly Rate**. The rate you charge for your time.
  - **Commission**. Commission that you pay to another party, such as a service host, for any time you bill. This setting is depreciated and will be set on individual contacts in the next version.
  - **Time Format**. Format you use for expressing time and date.
  - **Last Invoice**. The number of the last invoice generated.
<br />
- ## Work
  Temporal calculations for checking what-if scenarios and one-off decisions.
  - **Rate Projector**. If you have a net (take-home) amount in mind that you need to establish, and have to take out a percentage of all income for taxes or commission, then use the Rate projector calculator on the Work sheet. For example, if I need to make $50.00/hr after taxes of 20%, I'll need to charge a rate of $62.50/hr to personally receive $50.00.
  - **User Calculations - Weekly**. Using the time accumulated, the amount of time in the week, and current check-in time, show the ideal check-out time.
  - **User Calculations - Daily**. Using the time in, the number of hours needed, and length of break, display the ideal check-out time for today.
  - **General Calculations - Hours to Decimal**. Convert time-formatted hours to decimal fraction.
  - **General Calculations - Decimal to Hours**. Convert decimal fractional hours to time-formatted hours, minutes, and seconds.
  - **General Calculations - Range**. Convert a range of time-formatted values to decimal hours.
  - **General Calculations - Forecast (From Beginning of Day)**. Given the current number of accumulated hours and the name of the current day, calculate the number of days remaining and the check-out time on the equivalent of Friday.
  - **General Calculations - Award from Cyclical Time Spent**.  Given the payment per unit and time spent per unit, show the about of payment for all major time periods from second to year.
  - **General Calculations - Calculate Award Needed**. Given the payment required per time unit and the time required to complete the task, in a specified time unit, display the amount of payment needed per task iteration.
  - **General Calculations - Holidays**. Keep track of the number allotted and up to five uses each of Vacation, Personal Days, and Sick Days.
  - **General Calculations - Day Track**. Given the entered decimal hours for each of the seven days of the week, display the total number of hours spent in the week.
  - **General Calculations - Hour Track**. Given the in and out times entered for each of the seven days of the week, display the total number of decimal hours spent in the week.
<br />
- ## Timesheet
  52 individual week grids representing the clock-in and clock-out times of each day, summarized with the total decimal hours of each day and week.
<br />
- ## Sheets 01 to 12
  Individual hourly expenditures for each month of the year. When starting a new task, select the project, and task, then type **Now** in the Start column. After completing work, either through completion or interruption, select the End column and type **Now**. Man-hours, Billable amount, Charge, Invoiced, Received, and Due are all calculated automatically. Note that billing is not very dependent on the monthly timeline. You can send out as many as multiple invoices per month, as few as one invoice after the completion of an entire project, or anywhere in between.

  Each entry on the month sheet represents a single span of uninterrupted time. A basic example is illustrated in the following table.
  |Task|Start|End|
  |----|-----|---|
  |Product Testing|8:25AM|10:13AM|
  |Product Testing|10:37AM|12:45PM|

  Following is a list of the columns found in each of these sheets.
   - **Active**. Value indicating whether this item will be invoiced. If **1**, the item is calculated on invoices.
   - **Sent**. Value indicating whether the item has been sent to an invoice. If **1**, the item has been invoiced.
   - **Service**. Drop-down list of service records, formatted as *ContactCode*_*ServiceName*. If no services have been defined, this list will be empty. In this version, you can either leave this item blank, or select an existing item. Ad-hoc entries are not yet supported.
   - **Project**. Freeform name of a project or assignment.
   - **Task**. Name of the task being performed.
   - **Start**. Date and time upon which this effort was started. This cell accepts the word **now**, time-only, and full date and time formatted values.
   - **End**. Date and time upon which this effort was stopped. This cell accepts the word **now**, time-only, and full date and time formatted values.
   - **MH**. (Calculated). Number of man-hours elapsed between Start and End times.
   - **Billable**. (Calculated). Number of billable man-hours in MH. In this version, if Active = 1, then Billable = MH, otherwise, Billable = 0.
   - **Charge**. (Calculated). The charge to be invoiced to the customer. If Service is blank, this will be equal to Billable * Config.Rate. With a service record selected, the value is Billable * Service.Rate.
   - **Invoiced**. (Calculated or user entry). The amount charged for this entry on the invoice.
   - **Received**. (Calculated or user entry). The payment amount received for this entry.
   - **Due**. (Calculated). The result of Invoiced - Received.
<br />
- ## Summary
  Invoices and receipts over the course of the year. Informational only. All values are calculated from other areas in the workbook.

  The rows are labeled January through December to capture the statistics for each month of the year. Following is a brief description of columns.

   - **Billable**. Number of active decimal man-hours spent in each month.
   - **Invoiced**. Total amount invoiced to the customer for each month.
   - **Received**. Total amount received from customers in each month.
   - **Due**. Current amount still outstanding for each month.
<br />
- ## Log Report
  Separate log of activities performed for a specific customer over a specified period of time. This sheet is similar to an invoice, but is used more often as an intermediary report sent to the customer in a specific time interval between invoices.

  Following are the columns found in this sheet.
   - **Project**. Name of the project to which the effort was applied.
   - **Task**. Name of the task.
   - **Start**. Starting date and time of the task.
   - **End**. Ending date and time of the task.
   - **MH**. Decimal man-hours spent on the entry.
   - **Charge**. The charge that will be applied in the invoice.
<br />
- ## Service Invoice
  Last printed invoice for the selected customer and for a specified period of time.

  Following is a brief description of the fields present on this sheet.

   - **Title**. (Static). Service Invoice. You can rename this title if needed.
   - **Invoicer Name**. The name of the person or company creating the invoice. The source of this value is Config[Invoice Name].
   - **Invoicer Address**. The address of the person or company creating the invoice. Source found at Config[Invoice Address].
   - **Bill To**. Customer name and address to which the bill will be sent.
   - **Ship To**. Customer name and address to which products and project results will be sent.
   - **Invoice \#**. The unique number of the invoice. In this version the number is auto-generated, and follows the format *yyyymmddNN*, where *yyyy* is the four digit year, *mm* is the two digit month, *dd* is the two digit day of the month, and *NN* is the number of invoices generated on this day, including this one.
   - **Invoice Date**. Date upon which the invoice was generated.
   - **Project**. Area where projects listed in the included time entries will be cited.
   - **Due Date**. Date upon which full payment for this invoice is due.
   - **Details**. Multiple rows, each with the following columns.
     - **Date**. Date upon which the listed effort was completed. This value is equal to the corresponding End cell on the month sheet.
     - **Project**. Name of the project to which the effort was dedicated.
     - **Task**. Name of the task performed.
     - **Man Hours**. Decimal man-hours accrued on this entry.
     - **Rate**. Rate per hour assigned to this entry, either by Config[Rate] setting, or selected Services[Customer_ServiceName].
     - **Charge**. Line-item charge for this entry.
   - **Total Time**. The total number of decimal man-hours on this invoice.
   - **Total Due**. The total monetary charge on this invoice.
   - **Terms and Conditions**. Payment is due within 15 days of this invoice. In this version, the terms are NET 15 days. Dynamically selectable payment terms are scheduled for an upcoming version.
   - **Thank you**. Thank you message to the customer.
<br />
- ## Contacts
  Table of customers and customer-specific settings.

  The following columns are found on the Contacts table.
   - **Code**. Short code, typically four characters, used to abbreviate the contact name for use as a key in various contexts.
   - **Bill To Name**. Title of the customer as currently displayed in most areas of this file.
   - **Bill To Address**. Street number, street name, and suite number of customer.
   - **Bill To City State Zip**. City name, state code, and zip code of the customer.
   - **Ship To Name**. Title of the customer for shipping purposes.
   - **Ship To Address**. Street number, street name, and suite of the customer.
   - **Ship To City State Zip**. City name, state code, and zip code of the customer.
<br />
- ## Services
  Table of services defined per contact.

  Following is a brief description of each column in this table.
   - **Customer**. Drop-down list of defined contact codes from the Contacts table.
   - **Service**. Name of the service to provide to the selected contact.
   - **Rate per hr**. Hourly rate to charge to the selected contact for the selected service.
   - **Commission**. Commission percentage to be paid for work performed in this service.
<br />
- ## Training
  Reminders of training completed over the course of the year.

  Following are the columns found in the Training sheet.
   - **Date**. Date of training completion.
   - **Conference**. Name of the seminar, conference, or class.
   - **Hours**. Decimal number of hours spent in the course.
   - **Type**. Type of class.
   - **Description**. Brief synopsis of the course or seminar.
<br />
